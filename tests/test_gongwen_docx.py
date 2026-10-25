# -*- coding: utf-8 -*-
"""gongwen_docx 的回归测试。

不依赖真实字体：凡是要生成文档的用例，都先把字体检查打桩成「齐全」，
这样 Linux / macOS 的 CI 也能跑。真正跑字体检测的用例只测分支逻辑。
"""
import contextlib
import io
import os
import shutil
import sys
import tempfile
import types
import unittest
from unittest import mock

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(HERE), "scripts")
sys.path.insert(0, SCRIPTS)

import gongwen_docx as gw  # noqa: E402
from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Mm  # noqa: E402


def _args(**kw):
    base = dict(out="a.docx", title="t", date=None, to=None, body=None, attach=None)
    base.update(kw)
    return types.SimpleNamespace(**base)


class ParseBodyTest(unittest.TestCase):
    def test_blank_line_splits(self):
        self.assertEqual(gw.parse_body("第一行\n第二行\n\n第二段\n"),
                         [("p", "第一行第二行"), ("p", "第二段")])

    def test_no_blank_line_falls_back_to_one_line_per_par(self):
        blocks = gw.parse_body("甲\n乙\n丙\n")
        self.assertEqual([b[1] for b in blocks], ["甲", "乙", "丙"])

    def test_heading_levels(self):
        blocks = gw.parse_body("# 一、\n## （一）\n### 1.\n#### （1）\n")
        self.assertEqual([b[0] for b in blocks], ["h1", "h2", "h3", "h4"])
        self.assertEqual([b[1] for b in blocks], ["一、", "（一）", "1.", "（1）"])

    def test_hash_escape(self):
        self.assertEqual(gw.parse_body("\\# 正文里的井号\n"), [("p", "# 正文里的井号")])

    def test_ascii_join_adds_one_space(self):
        self.assertEqual(gw.parse_body("用 DeepSeek\nHarness 测试\n\n尾段\n"),
                         [("p", "用 DeepSeek Harness 测试"), ("p", "尾段")])

    def test_chinese_join_has_no_space(self):
        self.assertEqual(gw.parse_body("第一行\n第二行\n\n尾段\n"),
                         [("p", "第一行第二行"), ("p", "尾段")])


class TextUtilTest(unittest.TestCase):
    def test_width_chars(self):
        self.assertEqual(gw.text_width_chars("××市"), 3.0)
        self.assertEqual(gw.text_width_chars("AB"), 1.0)

    def test_valid_date(self):
        self.assertTrue(gw._valid_date("2025 年 10 月 25 日"))
        self.assertTrue(gw._valid_date("2025年1月5日"))
        self.assertFalse(gw._valid_date("2025-10-25"))
        self.assertFalse(gw._valid_date("2025 年 2 月 30 日"))

    def test_strip_attach_prefix(self):
        self.assertEqual(gw.strip_attach_prefix("附件：1. 某某表"), "1. 某某表")
        self.assertEqual(gw.strip_attach_prefix("2. 别的"), "2. 别的")


class ValidateArgsTest(unittest.TestCase):
    def test_ok(self):
        errors, notices = gw.validate_args(_args(title="正常标题"))
        self.assertEqual((errors, notices), ([], []))

    def test_out_must_be_docx(self):
        errors, _ = gw.validate_args(_args(out="a.txt"))
        self.assertTrue(any(".docx" in e for e in errors))

    def test_blank_title(self):
        errors, _ = gw.validate_args(_args(title="   "))
        self.assertTrue(any("--title" in e for e in errors))

    def test_to_needs_fullwidth_colon(self):
        errors, _ = gw.validate_args(_args(to="各处室"))
        self.assertTrue(any("冒号" in e for e in errors))
        errors, _ = gw.validate_args(_args(to="各处室："))
        self.assertEqual(errors, [])

    def test_bad_date(self):
        errors, _ = gw.validate_args(_args(date="2025-10-25"))
        self.assertTrue(any("--date" in e for e in errors))

    def test_attach_prefix_is_notice_and_stripped_later(self):
        errors, notices = gw.validate_args(_args(attach=["附件：1. 某某表"]))
        self.assertEqual(errors, [])
        self.assertTrue(notices)


class ExitCodeTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)

    def _main(self, argv, missing=([], None)):
        with mock.patch.object(gw, "find_missing_fonts", return_value=missing):
            return gw.main(argv)

    def test_check_fonts_ok(self):
        self.assertEqual(self._main(["--check-fonts"]), 0)

    def test_check_fonts_missing_returns_2(self):
        self.assertEqual(self._main(["--check-fonts"], ([("黑体", "一级标题")], None)), 2)

    def test_check_fonts_detect_error_returns_5(self):
        self.assertEqual(self._main(["--check-fonts"], ([("黑体", "一级标题")], "模拟检测失败")), 5)

    def test_generate_without_force_when_detect_error_returns_5(self):
        code = self._main(["--out", os.path.join(self.tmp, "x.docx"), "--title", "t"],
                          ([("黑体", "一级标题")], "模拟检测失败"))
        self.assertEqual(code, 5)

    def test_blank_title_returns_1(self):
        self.assertEqual(self._main(["--out", os.path.join(self.tmp, "x.docx"), "--title", "   "]), 1)

    def test_to_without_colon_returns_1(self):
        self.assertEqual(
            self._main(["--out", os.path.join(self.tmp, "x.docx"), "--title", "t", "--to", "各处室"]), 1)

    def test_empty_body_file_exits_1(self):
        body = os.path.join(self.tmp, "empty.txt")
        open(body, "w", encoding="utf-8").close()
        with self.assertRaises(SystemExit) as ctx:
            self._main(["--out", os.path.join(self.tmp, "x.docx"), "--title", "t", "--body", body])
        self.assertEqual(ctx.exception.code, 1)

    def test_page_problem_returns_4(self):
        def fake_build(args):
            doc = mock.Mock()
            doc.save = lambda path: open(path, "wb").write(b"x")
            return doc, [], ["页码炸了"]

        with mock.patch.object(gw, "find_missing_fonts", return_value=([], None)), \
                mock.patch.object(gw, "build_document", side_effect=fake_build):
            code = gw.main(["--out", os.path.join(self.tmp, "x.docx"), "--title", "t"])
        self.assertEqual(code, 4)


class GenerateTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)

    def _generate(self, body_text=None, **extra):
        argv = ["--out", os.path.join(self.tmp, "out.docx"), "--title", "关于开展××工作的通知",
                "--to", "各处室：", "--sign", "××市人民政府办公室",
                "--date", "2025 年 10 月 25 日",
                "--attach", "1. ××××表", "--attach", "2. ××××清单",
                "--note", "此件公开发布"]
        if body_text is not None:
            path = os.path.join(self.tmp, "body.txt")
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(body_text)
            argv += ["--body", path]
        for key, value in extra.items():
            argv += [key, value]
        with mock.patch.object(gw, "find_missing_fonts", return_value=([], None)), \
                contextlib.redirect_stdout(io.StringIO()):
            code = gw.main(argv)
        self.assertEqual(code, 0)
        return Document(os.path.join(self.tmp, "out.docx"))

    def test_page_setup(self):
        doc = self._generate()
        s = doc.sections[0]
        self.assertAlmostEqual(s.top_margin.mm, 37, places=1)
        self.assertAlmostEqual(s.bottom_margin.mm, 35, places=1)
        self.assertAlmostEqual(s.left_margin.mm, 28, places=1)
        self.assertAlmostEqual(s.right_margin.mm, 26, places=1)
        self.assertAlmostEqual(s.footer_distance.mm, 28, places=1)

    def test_title_font_and_centering(self):
        doc = self._generate()
        title = [p for p in doc.paragraphs if p.text.strip()][0]
        self.assertEqual(title.text, "关于开展××工作的通知")
        run = title.runs[0]
        self.assertEqual(run.font.size.pt, 22)
        rfonts = run._element.rPr.rFonts
        self.assertEqual(rfonts.get(qn("w:eastAsia")), "方正小标宋简体")

    def test_body_font_and_indent(self):
        doc = self._generate("正文第一段。\n\n正文第二段。\n")
        body = [p for p in doc.paragraphs if p.text.startswith("正文第一段")][0]
        run = body.runs[0]
        self.assertEqual(run.font.size.pt, 16)
        self.assertEqual(run._element.rPr.rFonts.get(qn("w:eastAsia")), "仿宋_GB2312")
        ind = body._p.pPr.find(qn("w:ind"))
        self.assertEqual(ind.get(qn("w:firstLineChars")), "200")
        self.assertAlmostEqual(body.paragraph_format.line_spacing.pt, 29, places=1)

    def test_heading_fonts(self):
        doc = self._generate("# 一、总体要求\n\n## （一）主要任务\n\n### 1. 第一项\n\n#### （1）细节\n")
        fonts = {p.text: p.runs[0]._element.rPr.rFonts.get(qn("w:eastAsia"))
                 for p in doc.paragraphs if p.text.strip()}
        self.assertEqual(fonts["一、总体要求"], "黑体")
        self.assertEqual(fonts["（一）主要任务"], "楷体_GB2312")
        self.assertEqual(fonts["1. 第一项"], "仿宋_GB2312")
        self.assertEqual(fonts["（1）细节"], "仿宋_GB2312")

    def test_attachment_hanging_indent(self):
        doc = self._generate()
        first = [p for p in doc.paragraphs if p.text == "附件：1. ××××表"][0]
        second = [p for p in doc.paragraphs if p.text == "2. ××××清单"][0]
        ind1 = first._p.pPr.find(qn("w:ind"))
        self.assertEqual(ind1.get(qn("w:hangingChars")), "300")
        self.assertAlmostEqual(first.paragraph_format.left_indent.pt, 80, places=1)
        self.assertAlmostEqual(first.paragraph_format.first_line_indent.pt, -48, places=1)
        self.assertAlmostEqual(second.paragraph_format.left_indent.pt, 80, places=1)

    def test_page_number_field(self):
        doc = self._generate()
        xml = doc.sections[0].footer.paragraphs[0]._p.xml
        self.assertIn("PAGE", xml)
        self.assertIn("— ", doc.sections[0].footer.paragraphs[0].text)

    def test_note_wrapped_in_parentheses(self):
        doc = self._generate()
        note = [p for p in doc.paragraphs if "此件公开发布" in p.text][0]
        self.assertEqual(note.text, "（此件公开发布）")

    def test_no_absolute_path_in_log(self):
        buf = io.StringIO()
        with mock.patch.object(gw, "find_missing_fonts", return_value=([], None)), \
                contextlib.redirect_stdout(buf):
            gw.main(["--out", os.path.join(self.tmp, "out.docx"), "--title", "t"])
        self.assertNotIn(self.tmp, buf.getvalue())
        self.assertIn("out.docx", buf.getvalue())

    def test_atomic_write_leaves_no_tmp(self):
        self._generate()
        self.assertEqual([f for f in os.listdir(self.tmp) if f.endswith(".tmp")], [])

    def test_output_dir_is_created(self):
        deep = os.path.join(self.tmp, "a", "b", "out.docx")
        with mock.patch.object(gw, "find_missing_fonts", return_value=([], None)), \
                contextlib.redirect_stdout(io.StringIO()):
            code = gw.main(["--out", deep, "--title", "t"])
        self.assertEqual(code, 0)
        self.assertTrue(os.path.isfile(deep))


if __name__ == "__main__":
    unittest.main()
