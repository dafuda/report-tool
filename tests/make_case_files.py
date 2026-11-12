"""產生測試用的新案資料檔（虛構內容）：簡報 .pptx 與試算表 .xlsx。

用法：python tests/make_case_files.py tests/fixtures
"""
import sys
from pathlib import Path

import openpyxl
from pptx import Presentation
from pptx.util import Inches


def make_pptx(path):
    prs = Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[1])
    s1.shapes.title.text = "設計說明簡報"
    s1.placeholders[1].text = "工程名稱：新竹縣竹北市虛構活動中心新建工程\n工期：540日曆天"
    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    s2.shapes.title.text = "建築概要"
    rows = [("項目", "內容"), ("構造形式", "SC"), ("地上層數", "3"), ("地下層數", "0")]
    tbl = s2.shapes.add_table(len(rows), 2, Inches(1), Inches(2), Inches(6), Inches(2)).table
    for r, (a, b) in enumerate(rows):
        tbl.cell(r, 0).text = a
        tbl.cell(r, 1).text = b
    s2.notes_slide.notes_text_frame.text = "基地面積：2,500 m²"
    prs.save(path)


def make_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "基本資料"
    for row in [("項目", "內容"), ("主辦機關", "新竹縣政府"), ("建蔽率", "60%")]:
        ws.append(row)
    wb.save(path)


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "tests/fixtures")
    out.mkdir(parents=True, exist_ok=True)
    make_pptx(out / "新案_設計說明簡報.pptx")
    make_xlsx(out / "新案_基本資料.xlsx")
    (out / "舊版簡報.ppt").write_bytes(b"not a real ppt")
    print("written", out)
