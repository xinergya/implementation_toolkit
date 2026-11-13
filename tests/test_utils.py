"""核心纯函数与工具类的单元测试。

运行方式：python -m unittest discover -s tests
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.pathutil import strip_quotes, sanitize_filename, format_size, is_valid_word_file
from utils.constants import com_error_kind
from tools.file_matcher_engine import AutoFileProcessorUI


class TestPathUtil(unittest.TestCase):
    def test_strip_quotes(self):
        self.assertEqual(strip_quotes('  "C:/a b.docx"  '), 'C:/a b.docx')
        self.assertEqual(strip_quotes("'x.docx'"), 'x.docx')
        self.assertEqual(strip_quotes(''), '')
        self.assertEqual(strip_quotes(None), '')

    def test_sanitize_filename(self):
        self.assertEqual(sanitize_filename('a/b:c*d?e"f<g>h|i'), 'a_b_c_d_e_f_g_h_i')
        self.assertEqual(sanitize_filename('张三_丰'), '张三_丰')
        self.assertEqual(sanitize_filename('x' * 120, max_len=50), 'x' * 50)

    def test_format_size(self):
        self.assertEqual(format_size(512), '512 B')
        self.assertEqual(format_size(2048), '2.0 KB')
        self.assertEqual(format_size(3 * 1024 * 1024), '3.00 MB')

    def test_is_valid_word_file(self):
        self.assertTrue(is_valid_word_file('a.docx'))
        self.assertTrue(is_valid_word_file('a.doc'))
        self.assertFalse(is_valid_word_file('~$a.docx'))
        self.assertFalse(is_valid_word_file('a.pdf'))


class TestComErrorKind(unittest.TestCase):
    def test_stuck(self):
        self.assertEqual(com_error_kind('错误 80080005 服务器运行失败'), 'stuck')

    def test_crash(self):
        self.assertEqual(com_error_kind('800706BA RPC 服务器不可用'), 'crash')
        self.assertEqual(com_error_kind('RPC 错误'), 'crash')

    def test_permission(self):
        self.assertEqual(com_error_kind('Permission denied'), 'permission')

    def test_unknown(self):
        self.assertIsNone(com_error_kind('unknown'))
        self.assertIsNone(com_error_kind(''))
        self.assertIsNone(com_error_kind(None))


class TestBoundaryMatch(unittest.TestCase):
    def setUp(self):
        self.dummy = object.__new__(AutoFileProcessorUI)

    def test_boundary(self):
        cases = [
            ('张三', '张三.jpg', True),
            ('张三', '张三_身份证.jpg', True),
            ('张三', '张三-身份证.jpg', True),
            ('张三', '张三丰.jpg', False),
            ('张三', '1张三.jpg', False),
            ('张三', 'a张三.jpg', False),
            ('张三', '张三', True),
            ('', 'x.jpg', False),
        ]
        for token, text, expected in cases:
            self.assertEqual(self.dummy._boundary_match(token, text), expected, (token, text))


if __name__ == '__main__':
    unittest.main()
