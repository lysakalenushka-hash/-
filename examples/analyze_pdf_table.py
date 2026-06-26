#!/usr/bin/env python3
"""Hello-world document analysis: PDF -> structured table -> Excel + CSV.

This is a minimal, self-contained template for the document-analysis workspace.
Run it with no arguments to generate a sample PDF and analyze it, or pass a PDF
path to analyze your own document.

    .venv/bin/python examples/analyze_pdf_table.py [input.pdf]

Each real analysis should live on its own branch ("separate space"); see AGENTS.md.
"""

import sys
import re
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd

OUT_DIR = Path("outputs")

SAMPLE_ROWS = [
    ("1001", "Охрана труда для руководителей", "40", "повышение квалификации"),
    ("1002", "Пожарная безопасность", "16", "повышение квалификации"),
    ("1003", "Безопасность дорожного движения", "24", "проф. переподготовка"),
    ("1004", "Промышленная безопасность", "72", "проф. переподготовка"),
]


# Base-14 PDF fonts (e.g. "helv") have no Cyrillic glyphs, so embed a Unicode
# TTF if one is available on the system for a clean demo.
_CYRILLIC_FONT = next(
    (
        p
        for p in (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        )
        if Path(p).exists()
    ),
    None,
)


def make_sample_pdf(path: Path) -> None:
    """Render a small one-page 'programs catalog' PDF for the demo."""
    doc = fitz.open()
    page = doc.new_page()
    fontname, fontfile = "helv", None
    if _CYRILLIC_FONT:
        fontname, fontfile = "F0", _CYRILLIC_FONT
        page.insert_font(fontname=fontname, fontfile=fontfile)
    text = "Каталог программ обучения\n\nID    Наименование программы    Часов    Вид обучения\n"
    for cid, name, hours, kind in SAMPLE_ROWS:
        text += f"{cid}    {name}    {hours}    {kind}\n"
    page.insert_text((50, 72), text, fontsize=11, fontname=fontname, fontfile=fontfile)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    doc.close()


def extract_rows(path: Path) -> list[dict]:
    """Extract program rows (lines starting with a 4-digit ID) from the PDF."""
    doc = fitz.open(str(path))
    rows = []
    for page in doc:
        for line in page.get_text("text").splitlines():
            line = line.strip()
            m = re.match(r"^(\d{4})\s+(.*?)\s+(\d{1,4})\s+(.+)$", line)
            if m:
                cid, name, hours, kind = m.groups()
                rows.append(
                    {
                        "ID": cid,
                        "Наименование программы": name.strip(),
                        "Часов": int(hours),
                        "Вид обучения": kind.strip(),
                    }
                )
    doc.close()
    return rows


def main() -> int:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    if len(sys.argv) > 1:
        pdf_path = Path(sys.argv[1])
        if not pdf_path.exists():
            print(f"Файл не найден: {pdf_path}")
            return 1
    else:
        pdf_path = OUT_DIR / "sample_programs.pdf"
        print(f"Входной PDF не задан — генерирую пример: {pdf_path}")
        make_sample_pdf(pdf_path)

    print(f"Анализирую документ: {pdf_path}")
    rows = extract_rows(pdf_path)
    print(f"  Найдено программ: {len(rows)}")

    if not rows:
        print("  Не удалось извлечь строки таблицы из документа.")
        return 1

    df = pd.DataFrame(rows)
    xlsx_path = OUT_DIR / "programs_table.xlsx"
    csv_path = OUT_DIR / "programs_table.csv"
    df.to_excel(xlsx_path, index=False)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    print("\n--- Извлечённая таблица ---")
    print(df.to_string(index=False))
    print(f"\nСредняя нагрузка (часов): {df['Часов'].mean():.1f}")
    print(f"Результат сохранён:\n  Excel: {xlsx_path}\n  CSV:   {csv_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
