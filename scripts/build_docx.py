#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_docx.py —— 按"教案模板.doc"的固定表格结构，渲染 N 次课时的教案 docx。

数据格式（JSON，--data）：
{
  "course": "传感器应用技术",          # 课程名称（通常=书名）
  "lessons": [
    {
      "index": 1, "total": 8,
      "chapter": "模块一 认识传感器／1.1 传感器基础知识",
      "objectives": "知识目标：...\\n能力目标：...\\n素质目标：...",
      "keypoints": "重点：...\\n难点：...",
      "methods": "讲授法、演示法",
      "prep": "学生预习...；教师准备...",
      "stages": [                       # 固定 4 环节
        {"name":"导入",      "content":"...", "intent":"..."},
        {"name":"新授内容",  "content":"...", "intent":"..."},
        {"name":"归纳小结",  "content":"...", "intent":"..."},
        {"name":"作业",      "content":"...", "intent":"..."}
      ],
      "notes": ""
    }, ...
  ]
}

用法：
  python3 build_docx.py --data data.json --out 教案.docx
  python3 build_docx.py --data data.json --out 教案.docx --per-lesson   # 每课时一个文件
"""
import argparse
import json
import os
import sys

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
except ImportError:
    sys.stderr.write("缺少 python-docx，正在尝试安装...\n")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT

FONT = "宋体"
SIZE = Pt(10.5)

# 列宽（cm）：课程名区 / 章节区 / 教学内容区(3列) / 设计意图区(3列)
COL_WIDTHS = [1.6, 2.0, 3.0, 3.0, 3.0, 1.3, 1.3, 1.3]  # 合计 15.5cm


def set_run_font(run, bold=False):
    run.font.name = FONT
    run.font.size = SIZE
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), FONT)


def write_cell(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, valign="center"):
    cell.text = ""
    # 垂直居中
    tcPr = cell._tc.get_or_add_tcPr()
    vA = OxmlElement("w:vAlign")
    vA.set(qn("w:val"), valign)
    tcPr.append(vA)
    lines = (text or "").split("\n")
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        if line == "":
            continue
        run = p.add_run(line)
        set_run_font(run, bold=bold)


def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "6")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)


def set_col_widths(table):
    tblW = table._tbl.tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        table._tbl.tblPr.append(tblW)
    tblW.set(qn("w:w"), str(int(sum(COL_WIDTHS) * 567)))
    tblW.set(qn("w:type"), "dxa")
    for r in table.rows:
        for i, w in enumerate(COL_WIDTHS):
            cell = r.cells[i]
            tcW = cell._tc.get_or_add_tcPr().find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                cell._tc.get_or_add_tcPr().append(tcW)
            tcW.set(qn("w:w"), str(int(w * 567)))
            tcW.set(qn("w:type"), "dxa")


def build_lesson_table(doc, lesson, course):
    stages = lesson.get("stages") or []
    while len(stages) < 4:  # 不足 4 环节补空，保证结构
        stages.append({"name": "", "content": "", "intent": ""})

    t = doc.add_table(rows=11, cols=8)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t)
    set_col_widths(t)
    c = lambda r, col: t.cell(r, col)  # 0-based

    # ---- 合并 ----
    c(0, 1).merge(c(0, 2))   # 课程名值 跨 c2-c3
    c(0, 4).merge(c(0, 5))   # 章节值 跨 c5-c6
    c(1, 1).merge(c(1, 7))   # 教学目标 值 跨 c2-c8
    c(2, 1).merge(c(2, 7))   # 重点难点
    c(3, 1).merge(c(3, 7))   # 教学方法
    c(4, 1).merge(c(4, 7))   # 课前准备
    c(5, 0).merge(c(9, 0))   # 教学程序 纵向合并 r6-r10
    c(5, 2).merge(c(5, 4))   # 教学内容 表头 跨 c3-c5
    c(5, 5).merge(c(5, 7))   # 设计意图 表头 跨 c6-c8
    c(6, 2).merge(c(6, 4)); c(6, 5).merge(c(6, 7))  # 导入
    c(7, 2).merge(c(7, 4)); c(7, 5).merge(c(7, 7))  # 新授内容
    c(8, 2).merge(c(8, 4)); c(8, 5).merge(c(8, 7))  # 归纳小结
    c(9, 2).merge(c(9, 4)); c(9, 5).merge(c(9, 7))  # 作业
    c(10, 1).merge(c(10, 7))  # 教学后记 值 跨 c2-c8

    # ---- 写入 ----
    chapter = lesson.get("chapter") or ""
    idx = lesson.get("index", 1)
    total = lesson.get("total", 1)

    write_cell(c(0, 0), "课程名称", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(c(0, 1), course)
    write_cell(c(0, 3), "授课章节(教学课题)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(c(0, 4), chapter)
    write_cell(c(0, 6), "课时", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(c(0, 7), f"{idx}／{total}")

    write_cell(c(1, 0), "教学目标", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(c(1, 1), lesson.get("objectives") or "")
    write_cell(c(2, 0), "重点难点", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(c(2, 1), lesson.get("keypoints") or "")
    write_cell(c(3, 0), "教学方法", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(c(3, 1), lesson.get("methods") or "")
    write_cell(c(4, 0), "课前准备", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(c(4, 1), lesson.get("prep") or "")

    write_cell(c(5, 0), "教学程序", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(c(5, 2), "教学内容", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(c(5, 5), "设计意图", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, st in enumerate(stages):
        r = 6 + i
        write_cell(c(r, 1), st.get("name", ""), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        write_cell(c(r, 2), st.get("content", ""))
        write_cell(c(r, 5), st.get("intent", ""))

    write_cell(c(10, 0), "教学后记", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(c(10, 1), lesson.get("notes") or "")
    return t

def build(data, out_path, per_lesson=False):
    lessons = data.get("lessons", [])
    course = data.get("course", "")
    if per_lesson:
        base, ext = os.path.splitext(out_path)
        for les in lessons:
            doc = Document()
            _set_default_style(doc)
            _add_title(doc, les, course)
            build_lesson_table(doc, les, course)
            p = os.path.join(os.path.dirname(out_path),
                             f"{base}_第{les.get('index',1)}课时{ext or '.docx'}")
            doc.save(p)
        return len(lessons)
    else:
        doc = Document()
        _set_default_style(doc)
        for i, les in enumerate(lessons):
            if i > 0:
                doc.add_page_break()
            _add_title(doc, les, course)
            build_lesson_table(doc, les, course)
        doc.save(out_path)
        return 1


def _set_default_style(doc):
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = SIZE
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def _add_title(doc, les, course):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{course}  教案（第 {les.get('index',1)} 课时 / 共 {les.get('total',1)} 课时）")
    run.font.name = FONT
    run.font.size = Pt(12)
    run.font.bold = True
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    p.paragraph_format.space_after = Pt(6)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True, help="教案数据 JSON 文件")
    ap.add_argument("--out", required=True, help="输出 docx 路径")
    ap.add_argument("--per-lesson", action="store_true", help="每课时一个独立 docx")
    args = ap.parse_args()
    with open(args.data, encoding="utf-8") as f:
        data = json.load(f)
    n = build(data, args.out, args.per_lesson)
    sys.stderr.write(f"OK: 生成 {n} 个 docx -> {args.out}\n")


if __name__ == "__main__":
    main()
