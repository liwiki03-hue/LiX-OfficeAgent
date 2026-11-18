import json
import os
from pathlib import Path
import subprocess
import sys
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def measured_conversion(command, environment, destination, source, report_path, prepare_seconds):
    started = time.monotonic()
    process = None
    report = {'pdf_first_seen_seconds': None, 'worker_prepare_seconds': prepare_seconds}
    sources = source if isinstance(source, tuple) else (source,)
    if len(sources) == 2:
        report['second_pdf_first_seen_seconds'] = None
    code = 3
    try:
        process = subprocess.Popen(command, env=environment, stdin=subprocess.DEVNULL,
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                   creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        report['libreoffice_launch_seconds'] = time.monotonic() - started
        waiting = time.monotonic()
        deadline = started + 25
        pdfs = [destination / (item.stem + '.pdf') for item in sources]
        while True:
            for index, pdf in enumerate(pdfs):
                name = 'pdf_first_seen_seconds' if index == 0 else 'second_pdf_first_seen_seconds'
                if report[name] is None and pdf.is_file():
                    report[name] = time.monotonic() - started
            returned = process.poll()
            if returned is not None:
                code = returned
                break
            if time.monotonic() >= deadline:
                break
            time.sleep(.01)
        report['libreoffice_wait_seconds'] = time.monotonic() - waiting
    finally:
        cleanup = time.monotonic()
        if process is not None:
            if process.poll() is None:
                process.kill()
            process.wait(timeout=2)
        report['libreoffice_cleanup_seconds'] = time.monotonic() - cleanup
        report['worker_conversion_seconds'] = time.monotonic() - started
        try:
            with report_path.open('x', encoding='utf-8') as stream:
                json.dump(report, stream)
        except (OSError, ValueError):
            pass
    return code

def main():
    pair = len(sys.argv) in {6, 7} and sys.argv[1] == '--convert-pair'
    if pair or len(sys.argv) in {5, 6} and sys.argv[1] in {"--convert", "--legacy", "--legacy-fixture", "--recalculate"}:
        prepared = time.monotonic()
        if pair:
            sources = tuple(map(Path, sys.argv[2:4]))
            destination, executable = map(Path, sys.argv[4:6])
            if any(item.suffix.lower() != '.docx' for item in sources) or sources[0].name.casefold() == sources[1].name.casefold():
                return 2
        else:
            source, destination, executable = map(Path, sys.argv[2:5])
            sources = (source,)
        source = sources[0]
        temporary = destination / "temp"
        temporary.mkdir()
        environment = dict(os.environ, TEMP=str(temporary), TMP=str(temporary))
        export_filter = {".docx": "writer_pdf_Export", ".xlsx": "calc_pdf_Export", '.pptx': 'impress_pdf_Export'}.get(source.suffix)
        target_format = 'pdf'
        if sys.argv[1] == '--recalculate' and source.suffix == '.xlsx':
            target_format, export_filter = 'xlsx', 'Calc MS Excel 2007 XML'
        if sys.argv[1] == '--legacy':
            target_format, export_filter = {'.doc': ('docx', 'Office Open XML Text'),
                                            '.xls': ('xlsx', 'Calc MS Excel 2007 XML')}.get(source.suffix, (None, None))
        if sys.argv[1] == '--legacy-fixture':
            target_format, export_filter = {'.docx': ('doc', 'MS Word 97'),
                                            '.xlsx': ('xls', 'MS Excel 97')}.get(source.suffix, (None, None))
        if export_filter is None:
            return 2
        command = [str(executable), f"-env:UserInstallation={(destination / 'profile').as_uri()}",
                   "--headless", "--nologo", "--nodefault", "--nofirststartwizard", "--norestore",
                   "--convert-to", f"{target_format}:{export_filter}", "--outdir", str(destination), *[item.as_uri() for item in sources]]
        timing_index = 6 if pair else 5
        if len(sys.argv) > timing_index and sys.argv[1] in {'--convert', '--convert-pair'}:
            return measured_conversion(command, environment, destination, sources, Path(sys.argv[timing_index]), time.monotonic() - prepared)
        try:
            return subprocess.run(command, env=environment, stdin=subprocess.DEVNULL,
                                  stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                  timeout=25, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).returncode
        except subprocess.TimeoutExpired:
            return 3
    if len(sys.argv) not in {4, 5}:
        return 2
    from office_agent.fidelity import FidelityError, compare_pdf_snapshots, pdf_snapshot
    left, right, target = map(Path, sys.argv[1:4])
    try:
        snapshots = (pdf_snapshot(left,reading_order=len(sys.argv)==5), pdf_snapshot(right,reading_order=len(sys.argv)==5))
        report = compare_pdf_snapshots(*snapshots)
        if len(sys.argv) == 5:
            import re
            from office_agent.fidelity import bounded_bytes
            data = bounded_bytes(Path(sys.argv[4]))
            if len(data) > 60000:
                raise FidelityError('anchor_capacity')
            anchors = json.loads(data)
            if type(anchors) is not list or len(anchors) > 12:
                raise FidelityError('anchor_capacity')
            from office_agent.range_alignment import align, VERSION
            report['anchor_alignment'] = align(anchors, snapshots)
            report['alignment_policy'] = VERSION + '; unique exact NFC/whitespace-normalized context, at most two adjacent pages; uncertain matches require review'
            report['visual_checks'] = {'missing_text': 'deterministic package changes separately verified',
                                       'table_overflow_images_heading_orphans_headers_footers': 'requires visual review; pixel differences are observations',
                                       'pagination_change': 'allowed; not automatically a failure'}
    except FidelityError as error:
        report = {"status": "error", "equivalent": None, "code": error.code}
    with target.open("x", encoding="utf-8") as stream:
        json.dump(report, stream, ensure_ascii=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
