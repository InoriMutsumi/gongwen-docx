#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""公文版式 Word 生成器（GB/T 9704-2012 版式部分，无红头）。

用法示例：
  # 只查字体
  python gongwen_docx.py --check-fonts

  # 生成文档
  python gongwen_docx.py --out 输出.docx \
      --title "关于××××的通知" --to "各处室：" --body 正文.txt \
      --sign "××××办公室" --date "2025年10月25日" \
      --attach "1. ××××表" --note "（此件公开发布）"

正文文件（--body）的层级标记：
  # 一级标题      → 3 号黑体        「一、」
  ## 二级标题      → 3 号楷体_GB2312 「（一）」
  ### 三级标题     → 3 号仿宋        「1.」
  其余非空行       → 正文段（3 号仿宋，首行缩进 2 字）
  空行用于分隔，不产生空段。

退出码：0 正常；2 缺字体未获允许；3 缺依赖；1 参数或读取错误。
"""

from __future__ import annotations

import argparse
import os
import sys

# ---------------------------------------------------------------- 版式常量

MM_TOP = 37          # 页边距 上
MM_BOTTOM = 35       # 页边距 下
MM_LEFT = 28         # 页边距 左
MM_RIGHT = 26        # 页边距 右
LINE_PITCH_PT = 29   # 固定行距（225mm ÷ 22 行 ≈ 29 磅）
LINE_PITCH_TWIPS = 580  # 同上，换成 twips 写进 docGrid
INDENT_CHARS = 2     # 首行缩进
SIZE_TITLE = 22      # 2 号
SIZE_BODY = 16       # 3 号
SIZE_PAGE = 14       # 4 号

FONT_TITLE = "方正小标宋简体"
FONT_BODY = "仿宋_GB2312"
FONT_H1 = "黑体"
FONT_H2 = "楷体_GB2312"
FONT_H3 = "仿宋_GB2312"
FONT_PAGE = "宋体"

REQUIRED_FONTS = [
    (FONT_TITLE, "标题"),
    (FONT_BODY, "正文"),
    (FONT_H1, "一级标题"),
    (FONT_H2, "二级标题"),
    (FONT_PAGE, "页码"),
]

# 注册表读不到时，用文件名兜底认字体（值统一大写，不含扩展名）
FONT_FILE_HINTS = {
    FONT_TITLE: {"FZXBSJW", "FZXIAOBIAOSONG-B05S", "FZXBSK", "FZXBSJW--GB1-0"},
    FONT_BODY: {"FZFSJW", "FZFSK", "SIMFANG", "FZFSJW--GB1-0"},
    FONT_H2: {"FZKTJW", "FZKATI", "SIMKAI", "FZKTJW--GB1-0", "KAIU"},
    FONT_H1: {"SIMHEI", "FZHTJW"},
    FONT_PAGE: {"SIMSUN", "SIMSUNB"},
}


def _strip_font_suffix(name: str) -> str:
    for suffix in (" (TrueType)", " (OpenType)", " (TrueType)"):
        if name.endswith(suffix):
            name = name[: -len(suffix)]
    return name.strip()


def installed_fonts() -> tuple[set[str], set[str]]:
    """返回 (已安装字体名集合, 已安装字体文件名主干集合)。"""
    names: set[str] = set()
    files: set[str] = set()

    if sys.platform == "win32":
        try:
            import winreg
        except ImportError:
            winreg = None  # type: ignore
        if winreg is not None:
            sub = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
            for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                try:
                    with winreg.OpenKey(root, sub) as key:
                        count = winreg.QueryInfoKey(key)[1]
                        for i in range(count):
                            name, value, _ = winreg.EnumValue(key, i)
                            names.add(_strip_font_suffix(str(name)))
                            base = os.path.splitext(os.path.basename(str(value)))[0]
                            if base:
                                files.add(base.upper())
                except OSError:
                    continue

    # 字体目录兜底（用户自装字体有时没进注册表）
    win_dir = os.environ.get("WINDIR", r"C:\Windows")
    dirs = [
        os.path.join(win_dir, "Fonts"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
    ]
    for d in dirs:
        try:
            for fn in os.listdir(d):
                base = os.path.splitext(fn)[0]
                if base:
                    files.add(base.upper())
        except OSError:
            continue

    return names, files


def find_missing_fonts() -> list[tuple[str, str]]:
    names, files = installed_fonts()
    missing: list[tuple[str, str]] = []
    for font, role in REQUIRED_FONTS:
        if font in names or font.lower() in {n.lower() for n in names}:
            continue
        hints = FONT_FILE_HINTS.get(font, set())
        if hints & files:
            continue
        missing.append((font, role))
    return missing


# ---------------------------------------------------------------- docx 依赖

def _import_docx():
    try:
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Mm, Pt
    except ImportError:
        print("缺少依赖 python-docx。请先执行：pip install python-docx", file=sys.stderr)
        raise SystemExit(3)
    return (Document, WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, OxmlElement, qn, Mm, Pt)


# ---------------------------------------------------------------- 排版工具

def make_helpers(Document, WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, OxmlElement, qn, Mm, Pt):
    def set_run_font(run, name, size, bold=False):
        run.font.name = name
        run.font.size = size
        run.bold = bold
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rFonts.set(qn(attr), name)

    def _set_ind_attr(par, attr, value):
        pPr = par._p.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)
        ind.set(qn(attr), str(value))

    def add_par(doc, text="", *, font=FONT_BODY, size=None, align=None,
                indent_chars=INDENT_CHARS, left_chars=0.0, right_chars=0.0,
                bold=False):
        if size is None:
            size = Pt(SIZE_BODY)
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(LINE_PITCH_PT)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        if align is not None:
            p.alignment = align
        # 先用 pt 值把 w:ind 建出来（保证 pPr 里顺序合法），再补字符值
        pf.first_line_indent = Pt(indent_chars * SIZE_BODY)
        if left_chars:
            pf.left_indent = Pt(left_chars * SIZE_BODY)
        if right_chars:
            pf.right_indent = Pt(right_chars * SIZE_BODY)
        if indent_chars:
            _set_ind_attr(p, "w:firstLineChars", int(indent_chars * 100))
        if left_chars:
            _set_ind_attr(p, "w:leftChars", int(left_chars * 100))
        if right_chars:
            _set_ind_attr(p, "w:rightChars", int(right_chars * 100))
        run = p.add_run(text)
        set_run_font(run, font, size, bold=bold)
        return p

    def blank_line(doc, count=1):
        for _ in range(count):
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.line_spacing = Pt(LINE_PITCH_PT)
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)

    def setup_section(doc):
        s = doc.sections[0]
        s.page_width = Mm(210)
        s.page_height = Mm(297)
        s.top_margin = Mm(MM_TOP)
        s.bottom_margin = Mm(MM_BOTTOM)
        s.left_margin = Mm(MM_LEFT)
        s.right_margin = Mm(MM_RIGHT)
        s.footer_distance = Mm(28)
        sectPr = s._sectPr
        docGrid = sectPr.find(qn("w:docGrid"))
        if docGrid is None:
            docGrid = OxmlElement("w:docGrid")
            sectPr.append(docGrid)
        docGrid.set(qn("w:type"), "lines")
        docGrid.set(qn("w:linePitch"), str(LINE_PITCH_TWIPS))
        return s

    def _append_field(par, instr, placeholder="1"):
        def field_run(build):
            r = par.add_run()
            build(r._r)
            return r

        def begin(el):
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), "begin")
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = " %s " % instr
            el.append(fc)
            el.append(it)

        def separate(el):
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), "separate")
            el.append(fc)
            t = OxmlElement("w:t")
            t.text = placeholder
            el.append(t)

        def end(el):
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), "end")
            el.append(fc)

        return [field_run(begin), field_run(separate), field_run(end)]

    def add_page_numbers(doc, section, hide_first_page=False):
        try:
            doc.settings.odd_and_even_pages_header_footer = True
        except Exception:
            pass

        def fill(par, align, edge_pt):
            par.alignment = align
            pf = par.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            if align == WD_ALIGN_PARAGRAPH.RIGHT:
                pf.right_indent = Pt(edge_pt)
            else:
                pf.left_indent = Pt(edge_pt)
            runs = [par.add_run("— ")]
            runs.extend(_append_field(par, "PAGE"))
            runs.append(par.add_run(" —"))
            for r in runs:
                set_run_font(r, FONT_PAGE, Pt(SIZE_PAGE))

        odd = section.footer.paragraphs[0]
        fill(odd, WD_ALIGN_PARAGRAPH.RIGHT, SIZE_BODY)  # 单页：居右空一字

        try:
            section.even_page_footer.is_linked_to_previous = False
            even = section.even_page_footer.paragraphs[0]
            fill(even, WD_ALIGN_PARAGRAPH.LEFT, SIZE_BODY)  # 双页：居左空一字
        except Exception:
            pass

        if hide_first_page:
            try:
                section.different_first_page_header_footer = True
            except Exception:
                pass

    return set_run_font, add_par, blank_line, setup_section, add_page_numbers


def text_width_chars(s: str) -> float:
    """估算字符宽度：全角算 1 字，半角算 0.5 字。"""
    # 全角判定：CJK 区与全角标点，外加 × 号这类在公文里按全角计宽的符号
    wide_symbols = {0xD7, 0x2014, 0x2018, 0x2019, 0x201C, 0x201D, 0x2026, 0x00B7}
    w = 0.0
    for ch in s:
        if ord(ch) > 0x2E7F or ord(ch) in wide_symbols:
            w += 1.0
        else:
            w += 0.5
    return w


def parse_body(text: str):
    blocks = []
    for raw in text.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.startswith("### "):
            blocks.append(("h3", line[4:].strip()))
        elif line.startswith("## "):
            blocks.append(("h2", line[3:].strip()))
        elif line.startswith("# "):
            blocks.append(("h1", line[2:].strip()))
        else:
            blocks.append(("p", line.strip()))
    return blocks


# ---------------------------------------------------------------- 主流程

def build_document(args):
    (Document, WD_ALIGN_PARAGRAPH, WD_LINE_SPACING,
     OxmlElement, qn, Mm, Pt) = _import_docx()
    (set_run_font, add_par, blank_line,
     setup_section, add_page_numbers) = make_helpers(
        Document, WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, OxmlElement, qn, Mm, Pt)

    doc = Document()
    section = setup_section(doc)

    # 标题：上方留 2 行空白，2 号方正小标宋简体，居中
    blank_line(doc, 2)
    add_par(doc, args.title, font=FONT_TITLE, size=Pt(SIZE_TITLE),
            align=WD_ALIGN_PARAGRAPH.CENTER, indent_chars=0)

    # 主送机关：标题下空 1 行，顶格
    if args.to:
        blank_line(doc, 1)
        add_par(doc, args.to, font=FONT_BODY, indent_chars=0)

    # 正文：主送机关下空 1 行
    body_text = ""
    if args.body:
        try:
            with open(args.body, "r", encoding="utf-8-sig") as fh:
                body_text = fh.read()
        except OSError as exc:
            print("读不到正文文件 %s：%s" % (args.body, exc), file=sys.stderr)
            raise SystemExit(1)
    blocks = parse_body(body_text)
    if blocks:
        blank_line(doc, 1)
    for kind, text in blocks:
        if kind == "h1":
            add_par(doc, text, font=FONT_H1)
        elif kind == "h2":
            add_par(doc, text, font=FONT_H2)
        elif kind == "h3":
            add_par(doc, text, font=FONT_H3)
        else:
            add_par(doc, text, font=FONT_BODY)

    # 附件说明：正文下空 1 行，左空 2 字
    if args.attach:
        blank_line(doc, 1)
        first = True
        for item in args.attach:
            label = "附件：%s" % item if first else item
            add_par(doc, label, font=FONT_BODY,
                    indent_chars=0, left_chars=INDENT_CHARS)
            first = False

    # 发文机关署名 + 成文日期：附件说明下空 1 行；日期右空 4 字，署名居中于日期
    if args.sign or args.date:
        blank_line(doc, 1)
        date_w = text_width_chars(args.date) if args.date else 0.0
        if args.sign:
            if args.date:
                # 以成文日期为准居中：署名比日期宽时向右伸出，窄时向内收
                right = 4.0 + (date_w - text_width_chars(args.sign)) / 2.0
                right = max(0.0, right)
            else:
                right = 4.0
            add_par(doc, args.sign, font=FONT_BODY, align=WD_ALIGN_PARAGRAPH.RIGHT,
                    indent_chars=0, right_chars=right)
        if args.date:
            add_par(doc, args.date, font=FONT_BODY, align=WD_ALIGN_PARAGRAPH.RIGHT,
                    indent_chars=0, right_chars=4.0)

    # 附注：成文日期下 1 行，左空 2 字，圆括号
    if args.note:
        note = args.note
        if not (note.startswith("（") and note.endswith("）")):
            note = "（%s）" % note.strip("（）()")
        add_par(doc, note, font=FONT_BODY, indent_chars=0, left_chars=INDENT_CHARS)

    add_page_numbers(doc, section, hide_first_page=args.hide_first_page_number)

    out_dir = os.path.dirname(os.path.abspath(args.out))
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    doc.save(args.out)
    return doc, blocks


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="按 GB/T 9704-2012 版式生成公文 Word（无红头）")
    parser.add_argument("--check-fonts", action="store_true", help="只检查字体")
    parser.add_argument("--force", action="store_true", help="缺字体也继续生成")
    parser.add_argument("--out", help="输出 .docx 路径")
    parser.add_argument("--title", help="标题")
    parser.add_argument("--to", help="主送机关，如「各处室：」")
    parser.add_argument("--body", help="正文纯文本文件（支持 # / ## / ### 层级）")
    parser.add_argument("--sign", help="发文机关署名")
    parser.add_argument("--date", help="成文日期，如 2025年10月25日")
    parser.add_argument("--attach", action="append", help="附件说明，可重复")
    parser.add_argument("--note", help="附注（圆括号内容）")
    parser.add_argument("--hide-first-page-number", action="store_true",
                        help="首页不编页码（发文惯例）")
    args = parser.parse_args(argv)

    missing = find_missing_fonts()

    if args.check_fonts:
        if not missing:
            print("字体齐全：方正小标宋简体 / 仿宋_GB2312 / 楷体_GB2312 / 黑体 / 宋体")
            return 0
        print("缺少以下公文字体（共 %d 项）：" % len(missing))
        for font, role in missing:
            print("  - %-18s（用于%s）" % (font, role))
        print("请先安装上述字体；临时可用 --force 继续，但生成的文档字体与规范不符。")
        return 2

    if not args.out or not args.title:
        parser.error("生成文档需要同时提供 --out 和 --title")

    if missing:
        print("缺少以下公文字体（共 %d 项）：" % len(missing), file=sys.stderr)
        for font, role in missing:
            print("  - %-18s（用于%s）" % (font, role), file=sys.stderr)
        if not args.force:
            print("未获允许，已停止。装上字体后重跑，或明确同意后加 --force。", file=sys.stderr)
            return 2
        print("已按 --force 继续生成，文档字体将不完全符合规范。", file=sys.stderr)

    doc, blocks = build_document(args)
    print("已生成：%s" % os.path.abspath(args.out))
    print("版式：A4，页边距 上%d/下%d/左%d/右%d mm，固定行距 %d 磅，每页 22 行"
          % (MM_TOP, MM_BOTTOM, MM_LEFT, MM_RIGHT, LINE_PITCH_PT))
    print("正文段数：%d" % len(blocks))
    print("字体：标题 %s %dpt｜正文 %s %dpt｜页码 %s %dpt"
          % (FONT_TITLE, SIZE_TITLE, FONT_BODY, SIZE_BODY, FONT_PAGE, SIZE_PAGE))
    if missing:
        print("注意：下列字体缺失，已按要求继续：%s"
              % "、".join(f for f, _ in missing))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
