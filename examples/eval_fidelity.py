import argparse
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx import Document
from docx.shared import Pt
from openpyxl import Workbook

from office_agent.fidelity import compare_documents


def sample_suite(directory: Path, *, executable=None):
    directory.mkdir(parents=True, exist_ok=False)
    word = directory / "source.docx"
    document = Document()
    document.add_heading("Fidelity baseline", 0)
    document.add_paragraph("Approved budget: 7319")
    document.add_paragraph("中文预算：7319 元。")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Item"
    table.cell(0, 1).text = "Amount"
    table.cell(1, 0).text = "Alpha"
    table.cell(1, 1).text = "7319"
    document.save(word)
    identical = directory / "identical.docx"
    identical.write_bytes(word.read_bytes())
    changed_text = directory / "changed-text.docx"
    document.paragraphs[1].runs[0].text = "Approved budget: 7320"
    document.save(changed_text)
    changed_style = directory / "changed-style.docx"
    document = Document(word)
    document.paragraphs[1].runs[0].font.size = Pt(22)
    document.save(changed_style)
    excel = directory / "source.xlsx"
    workbook = Workbook()
    workbook.active.append(["Item", "Amount"])
    workbook.active.append(["Alpha", 7319])
    workbook.save(excel)
    workbook.close()
    excel_copy = directory / "identical.xlsx"
    excel_copy.write_bytes(excel.read_bytes())
    cases = [("word_identical", word, identical), ("word_text_changed", word, changed_text),
             ("word_style_changed", word, changed_style), ("excel_identical", excel, excel_copy)]
    results = {name: compare_documents(directory, source, output, executable=executable)
               for name, source, output in cases}
    passed = all(report["inputs_unchanged"] and report["render"]["status"] in {"equivalent", "different"}
                 for report in results.values())
    passed &= all(results[name]["structure"]["equivalent"] is True and results[name]["render"]["equivalent"] is True
                  for name in ("word_identical", "excel_identical"))
    passed &= results["word_text_changed"]["render"].get("text_equal") is False
    passed &= results["word_style_changed"]["render"].get("text_equal") is True
    passed &= results["word_style_changed"]["render"].get("pixels_equal") is False
    return {"schema_version": 1, "passed": passed, "real_libreoffice": passed,
            "real_microsoft_office": False, "cases": results}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path)
    parser.add_argument("--source", type=Path)
    parser.add_argument("--copy", type=Path)
    parser.add_argument("--samples", type=Path)
    parser.add_argument("--libreoffice", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.samples:
        if any((args.root, args.source, args.copy)):
            parser.error("--samples cannot be combined with file inputs")
        report = sample_suite(args.samples.resolve(), executable=args.libreoffice)
        passed = report["passed"]
    else:
        if not all((args.root, args.source, args.copy)):
            parser.error("--root, --source and --copy are required without --samples")
        report = compare_documents(args.root.resolve(), args.source.absolute(), args.copy.absolute(),
                                   executable=args.libreoffice)
        passed = report["inputs_unchanged"] and report["render"]["status"] in {"equivalent", "different"}
    with args.output.open("x", encoding="utf-8") as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
        stream.write("\n")
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
