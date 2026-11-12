"""產生一份模擬母本結構的測試用 docx（不含真實母本內容）。

用法：python tests/make_sample_docx.py tests/fixtures/sample_master.docx
      python tests/make_sample_docx.py tests/fixtures/sample_master_v2.docx --v2
--v2 產生「改版」母本：第一章開頭多一段、一段文字略改、一段刪除，用來測試依文字內容記憶。
"""
import struct
import sys
import zlib
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls
from docx.oxml.ns import qn
from docx.shared import Inches


def tiny_png(rgb=(255, 0, 0)):
    """1x1 PNG；不同顏色產生不同的圖片檔（用來測試無圖名圖片依檔案內容記憶）。"""
    raw = b"\x00" + bytes(rgb)
    def chunk(t, d):
        c = struct.pack(">I", len(d)) + t + d
        return c + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def add_field(par, instr, result):
    def fld(kind):
        r = OxmlElement("w:r")
        f = OxmlElement("w:fldChar")
        f.set(qn("w:fldCharType"), kind)
        r.append(f)
        par._p.append(r)
    fld("begin")
    r = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    r.append(it)
    par._p.append(r)
    fld("separate")
    par.add_run(result)
    fld("end")


def add_textbox(par, text):
    """在段落中加一個 VML 文字方塊（封面常見）。"""
    xml = (
        '<w:r %s><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:320pt;height:40pt">'
        '<v:textbox><w:txbxContent><w:p><w:r><w:t>%s</w:t></w:r></w:p></w:txbxContent></v:textbox>'
        '</v:shape></w:pict></w:r>' % (nsdecls("w"), text)
    )
    par._p.append(parse_xml(xml))


def caption(doc, kind, ch, n, name):
    p = doc.add_paragraph(style="表名" if kind == "表" else "圖名")
    p.add_run(f"{kind}{ch}-")
    add_field(p, f" SEQ {kind}{ch} \\* ARABIC ", str(n))
    p.add_run(f" {name}")
    return p


def table(doc, rows):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            t.cell(i, j).text = v
    return t


def main(out, v2=False):
    doc = Document()
    for name in ["章節", "1-1", "表名", "圖名"]:
        doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)

    # 頁首（含工程名稱）
    doc.sections[0].header.paragraphs[0].text = "臺中市北屯區第28公墓新建納骨塔工程　設計階段施工風險評估報告"

    # 前置
    add_textbox(doc.add_paragraph(), "臺中市北屯區第28公墓新建納骨塔工程")
    doc.add_paragraph("臺中市北屯區第28公墓新建納骨塔工程")
    doc.add_paragraph("設計階段施工風險評估報告")
    if not v2:
        # 封面機關標誌：沒有圖名的圖片
        doc.add_paragraph().add_run().add_picture(BytesIO(tiny_png((0, 0, 255))), width=Inches(0.5))
    doc.add_paragraph("中華民國115年6月")
    if v2:
        # 改版：標誌移到日期下方（位置不同、圖片檔相同），另加一張沒有圖名的新照片
        doc.add_paragraph().add_run().add_picture(BytesIO(tiny_png((0, 0, 255))), width=Inches(0.5))
        doc.add_paragraph().add_run().add_picture(BytesIO(tiny_png((0, 160, 0))), width=Inches(0.5))
    doc.add_paragraph("代辨機關：臺中市政府")
    doc.add_section(WD_SECTION.NEW_PAGE)

    # 第一章
    doc.add_paragraph("第一章 工程計畫概要", style="章節")
    doc.add_paragraph("一、計畫緣起與目標", style="1-1")
    if v2:
        doc.add_paragraph("本計畫配合市府殯葬設施整體規劃辦理，期提升納骨空間之使用效率。")
    doc.add_paragraph("本案為臺中市北屯區第28公墓新建納骨塔工程，地下1層、地上4層RC造建築物。")
    if not v2:
        doc.add_paragraph("地下停車場出入口由國豐街253巷進、出。")
    doc.add_paragraph("二、計畫範圍及環境", style="1-1")
    doc.add_paragraph("基地位於松竹路旁，基地面積12,345 m²（約1.23公頃），年雨量約1,700 毫米以上。")
    doc.add_paragraph("本地區年平均雨量 127.2mm，相對濕度 78%。")
    caption(doc, "表", 1, 1, "工程基本資料表")
    table(doc, [["類別", "項目", "內容"], ["工址環境現況", "工程名稱", "臺中市北屯區第28公墓新建納骨塔工程"], ["工址環境現況", "構造", "地上4層、地下1層 RC造"]])
    doc.add_paragraph("")
    p = doc.add_paragraph()
    p.add_run().add_picture(BytesIO(tiny_png()), width=Inches(1))
    caption(doc, "圖", 1, 1, "基地位置圖")
    doc.add_section(WD_SECTION.NEW_PAGE)

    # 第二章
    doc.add_paragraph("第二章 準備作業", style="章節")
    doc.add_paragraph("一、施工風險評估小組之組成", style="1-1")
    doc.add_paragraph("本小組由設計單位組成，評估人員：吳嘉栩。")
    doc.add_paragraph("二、工址環境現況調查分析", style="1-1")
    # 表格以截圖呈現：表名下方是圖片段落
    caption(doc, "表", 2, 1, "基礎容許承載力")
    doc.add_paragraph().add_run().add_picture(BytesIO(tiny_png()), width=Inches(2))
    # 表格以截圖呈現：1 格表格內放圖片
    caption(doc, "表", 2, 2, "地下水位觀測紀錄")
    t = doc.add_table(rows=1, cols=1)
    t.cell(0, 0).paragraphs[0].add_run().add_picture(BytesIO(tiny_png()), width=Inches(2))
    p2 = doc.add_paragraph()
    p2.add_run().add_picture(BytesIO(tiny_png()), width=Inches(1))
    caption(doc, "圖", 2, 1, "施工風險評估作業流程示意")
    doc.add_paragraph("四、法規及施工規範彙整分析", style="1-1")
    if v2:
        doc.add_paragraph("營造安全衛生設施標準第 71 條規定，擋土支撐應由專任工程人員妥為設計。")
    else:
        doc.add_paragraph("營造安全衛生設施標準第 71 條規定，擋土支撐應由專任工程人員設計。")
    doc.add_section(WD_SECTION.NEW_PAGE)

    # 第三章
    doc.add_paragraph("第三章 設計方案評選", style="章節")
    doc.add_paragraph("一、主要施工項目設計方案評選", style="1-1")
    caption(doc, "表", 3, 1, "擋土設施設計方案評選表")
    table(doc, [["工程方案", "潛在危害", "可能之風險狀況", "備註"], ["擋土柱+斜撐", "倒塌", "擋土柱變形造成崩塌", ""]])
    doc.add_paragraph("評估日期：112年5月1日")
    caption(doc, "表", 3, 2, "電梯組立設計方案評選表")
    table(doc, [["工程方案", "潛在危害", "可能之風險狀況", "備註"], ["電梯井道組立", "墜落", "電梯井開口墜落", ""]])
    doc.add_section(WD_SECTION.NEW_PAGE)

    # 第四章
    doc.add_paragraph("第四章 設計成果施工風險評估", style="章節")
    doc.add_paragraph("二、施工風險評估", style="1-1")
    caption(doc, "表", 4, 1, "設計階段施工風險評估表")
    table(doc, [
        ["編號", "第一階作業", "危害類型", "可能之風險狀況", "可能性", "嚴重度", "風險值", "等級"],
        ["分項工程：A 假設工程", "", "", "", "", "", "", ""],
        ["A1", "安全設施", "跌倒", "設置生方向指示燈不足", "2", "2", "4", "M"],
        ["分項工程：B 擋土支撐工程", "", "", "", "", "", "", ""],
        ["B1", "擋土柱打設", "倒塌", "擋土柱傾倒", "3", "3", "6", "M"],
        ["分項工程：C 骨櫃工程", "", "", "", "", "", "", ""],
        ["C1", "骨櫃安裝", "被撞", "骨櫃吊運撞擊", "1", "2", "2", "L"],
    ])
    doc.add_section(WD_SECTION.NEW_PAGE)

    # 第五章
    doc.add_paragraph("第五章 設計階段施工風險評估成果之運用", style="章節")
    doc.add_paragraph("一、設計過程風險資訊傳遞", style="1-1")
    caption(doc, "表", 5, 1, "設計階段施工風險資訊傳遞")
    table(doc, [
        ["辦理事項", "產出之風險資訊", "傳遞對象", "後續辦理事項"],
        ["設計方案評選", "表3-1 擋土設施設計方案評選表", "主辦機關", "納入契約"],
        ["設計方案評選", "表3-7 吊掛作業設計方案評選表", "主辦機關", "納入契約"],
        ["風險評估", "表4-1 施工規劃階段施工風險評估表", "施工廠商", "納入施工計畫"],
    ])
    doc.add_section(WD_SECTION.NEW_PAGE)

    # 第六章
    doc.add_paragraph("第六章 結論", style="章節")
    doc.add_paragraph("本報告依據公共工程施工階段職業安全衛生設施規劃之原則撰寫。")

    Path(out).parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    print("written", out)


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    main(args[0] if args else "tests/fixtures/sample_master.docx", v2="--v2" in sys.argv)
