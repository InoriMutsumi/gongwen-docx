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
  #### 四级标题    → 3 号仿宋        「（1）」
  其余非空行       → 正文段（3 号仿宋，首行缩进 2 字）

正文行首真的要以 # 开头时，写成 \\# —— 反斜杠会被吃掉，井号原样保留。

分段规则：空行分段。连续的非空普通行合并成同一个正文段（段内中文直接相连，
英文/数字相接处补一个半角空格）；空行本身不产生空段。整份文件一个空行都没有
时，退回「一行一段」，免得整篇正文被合并成一大段。

层级标记只换字体，不自动加编号——「一、」「（一）」「1.」要自己写进文本。

字体检测只支持 Windows（读注册表 + 系统字体目录）。非 Windows，或一条字体信息
都读不出来时，脚本不猜：按「无法确认」处理，退出码 5，加 --force 可放行继续。

字体缺失时文件名仍按规范写入，脚本不做替换；--force / --allow-missing-fonts
只负责放行，实际显示由打开文档的机器决定。

退出码：0 正常；1 参数、读取或写入错误；2 缺字体未获允许；3 缺依赖；
        4 文档已生成但页码设置失败；5 字体检测不可用，未获允许。
"""

from __future__ import annotations

import argparse
import datetime
import os
import re
import sys
import tempfile


def _setup_console() -> None:
    """把控制台输出切成 UTF-8。

    Windows 上默认编码可能是 cp1252 / cp936，打印中文直接抛 UnicodeEncodeError，
    会让已经生成好的文档卡在最后一步。errors="replace" 保证不崩。

    只在 main() 里调用：这个脚本也会被测试或别的程序 import，import 时不该顺手
    改掉宿主的 stdout/stderr。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

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


def _os_error_text(exc: OSError) -> str:
    """只取错误描述，不带路径——日志可能原样贴到 issue 里。"""
    return exc.strerror or exc.__class__.__name__


def _error_text(exc: BaseException) -> str:
    """清理阶段的错误描述。非 OSError 只留类名：异常消息里可能带本机路径。"""
    if isinstance(exc, OSError):
        return exc.strerror or exc.__class__.__name__
    return exc.__class__.__name__


def _strip_font_suffix(name: str) -> str:
    for suffix in (" (TrueType)", " (OpenType)", " (TrueType)"):
        if name.endswith(suffix):
            name = name[: -len(suffix)]
    return name.strip()


def installed_fonts() -> tuple[set[str], set[str], list[str]]:
    """返回 (已安装字体名集合, 已安装字体文件名主干集合, 检测失败原因列表)。"""
    names: set[str] = set()
    files: set[str] = set()
    errors: list[str] = []

    if sys.platform != "win32":
        errors.append("当前平台 %s 不是 Windows，读不到字体注册表和 Windows 字体目录"
                      % sys.platform)
        return names, files, errors

    try:
        import winreg
    except ImportError:
        winreg = None  # type: ignore
    if winreg is not None:
        sub = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        for root, root_name in ((winreg.HKEY_LOCAL_MACHINE, "HKLM（本机）"),
                                (winreg.HKEY_CURRENT_USER, "HKCU（当前用户）")):
            try:
                with winreg.OpenKey(root, sub) as key:
                    count = winreg.QueryInfoKey(key)[1]
                    for i in range(count):
                        name, value, _ = winreg.EnumValue(key, i)
                        names.add(_strip_font_suffix(str(name)))
                        base = os.path.splitext(os.path.basename(str(value)))[0]
                        if base:
                            files.add(base.upper())
            except OSError as exc:
                errors.append("读注册表 %s 失败：%s" % (root_name, _os_error_text(exc)))
    else:
        errors.append("当前 Python 没有 winreg 模块，读不到字体注册表")

    # 字体目录兜底（用户自装字体有时没进注册表）
    win_dir = os.environ.get("WINDIR", r"C:\Windows")
    dirs = [
        ("系统字体目录", os.path.join(win_dir, "Fonts")),
        ("用户字体目录", os.path.join(os.environ.get("LOCALAPPDATA", ""),
                                "Microsoft", "Windows", "Fonts")),
    ]
    for label, d in dirs:
        try:
            for fn in os.listdir(d):
                base = os.path.splitext(fn)[0]
                if base:
                    files.add(base.upper())
        except OSError as exc:
            errors.append("读%s失败：%s" % (label, _os_error_text(exc)))

    return names, files, errors


def find_missing_fonts() -> tuple[list[tuple[str, str]], str | None]:
    """返回 (缺失字体, 检测不可用的原因)。

    原因是 None 表示检测生效、结论可信；非 None 表示一条字体信息都没读到，
    此时「缺失」只是全部必需字体，不代表真的没装。
    """
    names, files, errors = installed_fonts()
    if not names and not files:
        reason = "；".join(errors) if errors else "没有读到任何已安装字体"
        return [(font, role) for font, role in REQUIRED_FONTS], reason
    if errors:
        # 部分来源读失败：结论可能少算，但还没到「一条都没读到」的程度，
        # 所以不把退出码抬到 5，只用一句提醒把不确定性说清楚。
        print("提醒：字体检测不完整（%s），缺失结论可能不准。" % "；".join(errors),
              file=sys.stderr)
    missing: list[tuple[str, str]] = []
    for font, role in REQUIRED_FONTS:
        if font in names or font.lower() in {n.lower() for n in names}:
            continue
        hints = FONT_FILE_HINTS.get(font, set())
        if hints & files:
            continue
        missing.append((font, role))
    return missing, None


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
                hanging_chars=0.0, bold=False):
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
        # hanging_chars：续行比首行多缩进的字数（附件说明用）
        pf.first_line_indent = Pt((indent_chars - hanging_chars) * SIZE_BODY)
        if left_chars or hanging_chars:
            pf.left_indent = Pt((left_chars + hanging_chars) * SIZE_BODY)
        if right_chars:
            pf.right_indent = Pt(right_chars * SIZE_BODY)
        if indent_chars and not hanging_chars:
            _set_ind_attr(p, "w:firstLineChars", int(indent_chars * 100))
        if hanging_chars:
            _set_ind_attr(p, "w:hangingChars", int(hanging_chars * 100))
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
        """写入奇偶页页码；返回失败项列表，空列表表示全部成功。

        「单/双页居右、居左空一字」里的「一字」按正文字号（3 号 = 16 pt）计宽，
        与首行缩进同一把尺子；页码本身是 4 号字，两者口径不同，此处取正文口径。
        """
        problems = []
        try:
            doc.settings.odd_and_even_pages_header_footer = True
        except Exception as exc:
            problems.append("奇偶页页脚开关设置失败：%s" % exc)

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

        try:
            odd = section.footer.paragraphs[0]
            fill(odd, WD_ALIGN_PARAGRAPH.RIGHT, SIZE_BODY)  # 单页：居右空一字
        except Exception as exc:
            problems.append("单页页码写入失败：%s" % exc)

        try:
            section.even_page_footer.is_linked_to_previous = False
            even = section.even_page_footer.paragraphs[0]
            fill(even, WD_ALIGN_PARAGRAPH.LEFT, SIZE_BODY)  # 双页：居左空一字
        except Exception as exc:
            problems.append("双页页码写入失败：%s" % exc)

        if hide_first_page:
            try:
                section.different_first_page_header_footer = True
            except Exception as exc:
                problems.append("首页不编页码设置失败：%s" % exc)
        return problems

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


def _needs_space(prev: str, nxt: str) -> bool:
    """段内换行合并时：英文或数字相接补一个空格，中文直接相连。"""
    if not prev or not nxt:
        return False
    return prev.isascii() and nxt.isascii() and prev.isalnum() and nxt.isalnum()


def parse_body(text: str):
    """空行分段：连续非空普通行合并成一段，# 标记行单独成段。

    整份文件里一个空行都没有时，退回「一行一段」的老写法（否则会把整篇正文
    合并成一大段）。一旦出现空行，就一律按空行分段。
    """
    blocks: list[tuple[str, str]] = []
    buf: list[str] = []
    lines = text.splitlines()
    has_blank = any(not ln.strip() for ln in lines)

    def flush():
        if buf:
            blocks.append(("p", "".join(buf)))
            buf.clear()

    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            flush()
            continue
        if line.startswith("\\#"):
            # 转义：正文行首真要写 # 时用 \#，这里吃掉反斜杠
            line = line[1:]
        elif line.startswith("#### "):
            flush()
            blocks.append(("h4", line[5:].strip()))
            continue
        elif line.startswith("### "):
            flush()
            blocks.append(("h3", line[4:].strip()))
            continue
        elif line.startswith("## "):
            flush()
            blocks.append(("h2", line[3:].strip()))
            continue
        elif line.startswith("# "):
            flush()
            blocks.append(("h1", line[2:].strip()))
            continue
        piece = line.strip()
        if has_blank:
            if buf and _needs_space(buf[-1][-1:], piece[:1]):
                buf.append(" ")
            buf.append(piece)
        else:
            buf.append(piece)
            flush()
    flush()
    return blocks


# ---------------------------------------------------------------- 输入校验

def _valid_date(s: str) -> bool:
    m = re.match(r"^(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日$", s.strip())
    if not m:
        return False
    y, mo, d = (int(x) for x in m.groups())
    try:
        datetime.date(y, mo, d)
    except ValueError:
        return False
    return True


def strip_attach_prefix(s: str) -> str:
    s = s.strip()
    for prefix in ("附件：", "附件:"):
        if s.startswith(prefix):
            return s[len(prefix):].strip()
    return s


def validate_args(args) -> tuple[list[str], list[str]]:
    """返回 (阻断性错误, 提醒)。错误退出码 1，提醒只打印。"""
    errors: list[str] = []
    notices: list[str] = []

    if not args.out.lower().endswith(".docx"):
        errors.append("--out 必须以 .docx 结尾：%s" % _display_path(args.out))
    if not (args.title or "").strip():
        errors.append("--title 不能是空白")
    if args.date and not _valid_date(args.date):
        errors.append("--date 应为「YYYY 年 M 月 D 日」且日期真实存在：%s" % args.date)
    if args.to is not None and not args.to.strip():
        errors.append("--to 不能是空白（不要主送机关就别传这个参数）")
    elif args.to and not args.to.rstrip().endswith("："):
        errors.append("--to 应以全角冒号「：」结尾（主送机关后是冒号）：%s" % args.to)
    if args.body and not os.path.isfile(args.body):
        errors.append("正文文件不存在：%s" % _display_path(args.body))
    for item in args.attach or []:
        if item.strip().startswith(("附件：", "附件:")):
            notices.append("--attach 自带的「附件：」已去掉，避免重复：%s" % item)
    return errors, notices


# ---------------------------------------------------------------- 主流程

def _display_path(path: str) -> str:
    """日志里用的短路径：能相对当前目录就相对，否则只留文件名。"""
    try:
        rel = os.path.relpath(os.path.abspath(path))
    except (ValueError, OSError):
        return os.path.basename(path)
    if rel.startswith(".."):
        return os.path.basename(path)
    return rel


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
        except UnicodeDecodeError:
            print("正文文件不是 UTF-8 编码：%s（用记事本「另存为」选 UTF-8 再试）"
                  % _display_path(args.body), file=sys.stderr)
            raise SystemExit(1)
        except OSError as exc:
            print("读不到正文文件 %s：%s" % (_display_path(args.body), _os_error_text(exc)),
                  file=sys.stderr)
            raise SystemExit(1)
        if not body_text.strip():
            print("正文文件是空的：%s" % _display_path(args.body), file=sys.stderr)
            raise SystemExit(1)
    blocks = parse_body(body_text)
    if blocks:
        blank_line(doc, 1)
    for kind, text in blocks:
        if kind == "h1":
            add_par(doc, text, font=FONT_H1)
        elif kind == "h2":
            add_par(doc, text, font=FONT_H2)
        elif kind in ("h3", "h4"):
            # 三、四级标题字体同正文，区别只在于独立成段
            add_par(doc, text, font=FONT_H3)
        else:
            add_par(doc, text, font=FONT_BODY)

    # 附件说明：正文下空 1 行，左空 2 字；续行与附件名称首字对齐
    # 先滤掉空白项再决定要不要空行：整组都是空白时不该在正文后留一个空段
    attach_items = [text for text in
                    (strip_attach_prefix(item) for item in (args.attach or []))
                    if text]
    if attach_items:
        blank_line(doc, 1)
        for idx, text in enumerate(attach_items):
            if idx == 0:
                # 「附件：」占 3 字，悬挂 3 字让第 2 行起对齐到「1.」
                add_par(doc, "附件：%s" % text, font=FONT_BODY,
                        indent_chars=0, left_chars=INDENT_CHARS, hanging_chars=3.0)
            else:
                # 保持与第 1 条的「1.」对齐；悬挂 1.5 字让折行退到名称首字附近
                add_par(doc, text, font=FONT_BODY,
                        indent_chars=0, left_chars=INDENT_CHARS + 3.0,
                        hanging_chars=1.5)

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

    page_problems = add_page_numbers(doc, section, hide_first_page=args.hide_first_page_number)
    return doc, blocks, page_problems


class _ArgumentParser(argparse.ArgumentParser):
    """参数错误统一退出码 1，避免与「缺字体 = 2」混淆。"""

    def error(self, message):
        self.print_usage(sys.stderr)
        self.exit(1, "%s: 参数错误：%s\n" % (self.prog, message))


def main(argv=None):
    _setup_console()
    parser = _ArgumentParser(
        description="按 GB/T 9704-2012 版式生成公文 Word（无红头）")
    parser.add_argument("--check-fonts", action="store_true", help="只检查字体")
    parser.add_argument("--force", "--allow-missing-fonts", dest="force",
                        action="store_true",
                        help="缺字体也继续生成（只放行，不替换字体名）")
    parser.add_argument("--out", help="输出 .docx 路径")
    parser.add_argument("--title", help="标题")
    parser.add_argument("--to", help="主送机关，如「各处室：」")
    parser.add_argument("--body",
                        help="正文纯文本文件（支持 # / ## / ### / #### 层级，"
                             "\\# 转义行首井号）")
    parser.add_argument("--sign", help="发文机关署名")
    parser.add_argument("--date", help="成文日期，如 2025年10月25日")
    parser.add_argument("--attach", action="append", help="附件说明，可重复")
    parser.add_argument("--note", help="附注（圆括号内容）")
    parser.add_argument("--hide-first-page-number", action="store_true",
                        help="首页不编页码（发文惯例）")
    args = parser.parse_args(argv)

    if args.check_fonts:
        missing, detect_error = find_missing_fonts()
        if detect_error:
            print("字体检测不可用：%s" % detect_error, file=sys.stderr)
            print("无法确认字体是否齐全。确认已装好后，加 --force（别名 "
                  "--allow-missing-fonts）可跳过检查。", file=sys.stderr)
            return 5
        if not missing:
            print("字体齐全：方正小标宋简体 / 仿宋_GB2312 / 楷体_GB2312 / 黑体 / 宋体")
            return 0
        print("缺少以下公文字体（共 %d 项）：" % len(missing))
        for font, role in missing:
            print("  - %-18s（用于%s）" % (font, role))
        print("请先安装上述字体；临时可用 --force（别名 --allow-missing-fonts）继续，"
              "但脚本不会替换字体名，显示效果由打开文档的机器决定。")
        return 2

    if not args.out:
        parser.error("生成文档需要 --out")

    errors, notices = validate_args(args)
    for note in notices:
        print("提醒：%s" % note)
    if errors:
        for err in errors:
            print("参数错误：%s" % err, file=sys.stderr)
        return 1

    missing, detect_error = find_missing_fonts()
    if detect_error:
        print("字体检测不可用：%s" % detect_error, file=sys.stderr)
        if not args.force:
            print("无法确认字体是否齐全，已停止。确认已装好后加 --force 继续。",
                  file=sys.stderr)
            return 5
        print("已按 --force 继续：字体检查结果不可信，请自行确认字体已装。",
              file=sys.stderr)
    elif missing:
        print("缺少以下公文字体（共 %d 项）：" % len(missing), file=sys.stderr)
        for font, role in missing:
            print("  - %-18s（用于%s）" % (font, role), file=sys.stderr)
        if not args.force:
            print("未获允许，已停止。装上字体后重跑，或明确同意后加 --force。", file=sys.stderr)
            return 2
        print("已按 --force 继续生成：文档内仍写规范字体名，缺失字体由 Word/WPS "
              "打开时替代显示，实际效果不可控。", file=sys.stderr)

    doc, blocks, page_problems = build_document(args)

    out_dir = os.path.dirname(os.path.abspath(args.out))
    tmp_path = ""

    def _cleanup_tmp() -> None:
        if not tmp_path:
            return
        try:
            os.unlink(tmp_path)
        except Exception as cleanup_exc:
            # 清理本身出错只提醒，绝不替换掉真正的异常（下面 except 块里会原样抛出）
            print("提醒：临时文件没清掉：%s（%s）"
                  % (_display_path(tmp_path), _error_text(cleanup_exc)),
                  file=sys.stderr)

    try:
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        # 唯一临时名：同一目录里并发生成两份文档时不会互相覆盖
        fd, tmp_path = tempfile.mkstemp(
            prefix=os.path.basename(args.out) + ".", suffix=".tmp",
            dir=out_dir or None)
        os.close(fd)
        doc.save(tmp_path)
        os.replace(tmp_path, args.out)
    except OSError as exc:
        print("写入失败：%s（目标 %s）" % (_os_error_text(exc), _display_path(args.out)),
              file=sys.stderr)
        _cleanup_tmp()
        return 1
    except BaseException:
        # 不是 OSError 的意外（python-docx 内部异常等）也要清掉临时文件，
        # 再把异常原样抛出去——吞掉的话真正的 bug 就再也没人知道了。
        _cleanup_tmp()
        raise

    print("已生成：%s" % _display_path(args.out))
    print("版式：A4，页边距 上%d/下%d/左%d/右%d mm，固定行距 %d 磅，每页 22 行"
          % (MM_TOP, MM_BOTTOM, MM_LEFT, MM_RIGHT, LINE_PITCH_PT))
    print("正文段数：%d" % len(blocks))
    print("字体：标题 %s %dpt｜正文 %s %dpt｜页码 %s %dpt"
          % (FONT_TITLE, SIZE_TITLE, FONT_BODY, SIZE_BODY, FONT_PAGE, SIZE_PAGE))
    if detect_error:
        print("注意：字体未做检查（检测不可用），文档内仍按规范字体名写入。")
    elif missing:
        print("注意：下列字体缺失，文档内仍按规范字体名写入（打开时由系统替代）：%s"
              % "、".join(f for f, _ in missing))
    if page_problems:
        print("警告：页码设置未全部成功，请检查页脚：", file=sys.stderr)
        for problem in page_problems:
            print("  - %s" % problem, file=sys.stderr)
        return 4
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
