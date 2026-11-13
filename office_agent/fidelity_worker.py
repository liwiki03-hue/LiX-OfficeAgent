import json
import os
from pathlib import Path
import subprocess
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

def main():
    if len(sys.argv) == 5 and sys.argv[1] in {"--convert", "--legacy", "--legacy-fixture", "--recalculate"}:
        source, destination, executable = map(Path, sys.argv[2:])
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
                   "--convert-to", f"{target_format}:{export_filter}", "--outdir", str(destination), source.as_uri()]
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
        snapshots = (pdf_snapshot(left), pdf_snapshot(right))
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
