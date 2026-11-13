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
from tools.batch_img_insert_engine import (
    load_image_types, extract_template_variables, BatchImageInsertEngine,
    parse_keywords, match_keywords,
)


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


class TestBatchImageInsert(unittest.TestCase):
    def test_load_image_types(self):
        types = load_image_types()
        self.assertIsInstance(types, list)
        self.assertGreaterEqual(len(types), 1)
        for t in types:
            self.assertIn("name", t)
            self.assertIn("dir", t)

    def test_extract_template_variables(self):
        import tempfile
        from docx import Document
        tmp = tempfile.mkdtemp()
        tpl = os.path.join(tmp, 't.docx')
        d = Document()
        d.add_paragraph('姓名：{{姓名}}')
        d.add_paragraph('{{身份证}}')
        d.add_paragraph('{{学历证书}}')
        d.add_paragraph('{{社保}}')
        d.save(tpl)
        variables = extract_template_variables(tpl)
        self.assertTrue({'姓名', '身份证', '学历证书', '社保'}.issubset(variables))

    def test_engine_filters_empty_names(self):
        engine = BatchImageInsertEngine(
            'x.xlsx', 't.docx',
            [{'name': '身份证', 'dir': ''}, {'name': '', 'dir': ''}, {'name': '社保', 'dir': '/a'}],
            'out',
        )
        self.assertEqual([t['name'] for t in engine.image_types], ['身份证', '社保'])

    def test_parse_keywords(self):
        self.assertEqual(parse_keywords('毕业,学位，学士'), ['毕业', '学位', '学士'])
        self.assertEqual(parse_keywords(''), [])
        self.assertEqual(parse_keywords('身份证,,身份证'), ['身份证'])

    def test_match_keywords(self):
        self.assertTrue(match_keywords('/x/孙方磊 身份证.JPG', ['身份证']))
        self.assertTrue(match_keywords('/x/孙方磊 毕业证书.jpg', ['毕业', '学位']))
        self.assertFalse(match_keywords('/x/孙方磊 身份证.JPG', ['毕业', '学位']))

    def test_find_person_images_and_match(self):
        import tempfile
        from PIL import Image
        tmp = tempfile.mkdtemp()
        archive = os.path.join(tmp, '归档')
        os.makedirs(archive)
        person_dir = os.path.join(archive, '1 118513 孙方磊')
        os.makedirs(person_dir)
        for fn in ('118513 孙方磊 身份证.JPG', '118513 孙方磊 毕业证书.jpg',
                   '118513 孙方磊 学士学位证书.jpg', '118513 孙方磊-PMP证书_1.png'):
            Image.new('RGB', (20, 20)).save(os.path.join(person_dir, fn))
        shebao = os.path.join(tmp, '社保')
        os.makedirs(shebao)
        Image.new('RGB', (20, 20)).save(os.path.join(shebao, '孙方磊_1.jpg'))

        engine = BatchImageInsertEngine('x', 't', [], 'out')
        # 分人子目录模式
        self.assertEqual(len(engine.find_person_images('孙方磊', '118513', archive)), 4)
        # 平铺模式
        self.assertEqual(len(engine.find_person_images('孙方磊', '', shebao)), 1)

        types = [
            {'name': '身份证', 'dir': archive, 'keywords': '身份证'},
            {'name': '学历证书', 'dir': archive, 'keywords': '毕业,学位,学士'},
            {'name': '资质证书', 'dir': archive, 'keywords': ''},
            {'name': '社保证明', 'dir': shebao, 'keywords': ''},
        ]
        engine2 = BatchImageInsertEngine('x', 't', types, 'out')
        result = engine2._match_types_for_person('孙方磊', '118513')
        self.assertEqual(len(result['身份证']), 1)
        self.assertEqual(len(result['学历证书']), 2)
        self.assertEqual(len(result['资质证书']), 1)
        self.assertEqual(len(result['社保证明']), 1)


if __name__ == '__main__':
    unittest.main()
