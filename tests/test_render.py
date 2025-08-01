import unittest
from xml.etree import ElementTree as ETree

from xlsx_streaming import render

from .utils import gen_xlsx_sheet


class TestOpenXML(unittest.TestCase):

    def test_rm_namespace(self):
        element = ETree.fromstring(gen_xlsx_sheet())
        render.rm_namespace(element)
        element_iter = element.iter()
        self.assertEqual(next(element_iter).tag, 'worksheet')
        self.assertEqual(next(element_iter).tag, 'sheetPr')
        self.assertEqual(next(element_iter).tag, 'outlinePr')

    def test_get_elements_from_template(self):
        header, views, template = render.get_elements_from_template(gen_xlsx_sheet())  # pylint: disable=unbalanced-tuple-unpacking
        self.assertTrue(header is None)
        self.assertTrue(views is None)
        self.assertEqual(template.tag, 'row')
        self.assertEqual(template.get('r'), '1')
        element_iter = template.iter()
        next(element_iter)
        child = next(element_iter)
        self.assertEqual(child.tag, 'c')
        self.assertEqual(child.get('r'), 'A1')

    def test_views_from_template(self):
        header, views, _ = render.get_elements_from_template(
            gen_xlsx_sheet(with_header=True, with_views=True)
        )
        self.assertTrue(header is not None)
        self.assertTrue(views is not None)
        self.assertEqual(views.tag, 'sheetViews')
        self.assertEqual(len(views), 1)
        self.assertEqual(len(views[0]), 1)
        self.assertEqual(views[0][0].tag, 'pane')
        self.assertEqual(views[0][0].get('xSplit'), None)
        self.assertEqual(views[0][0].get('ySplit'), '1')
        self.assertEqual(views[0][0].get('state'), 'frozen')

    def test_get_column(self):
        self.assertEqual(render.get_column(ETree.Element('c', r='A1')), 'A')
        self.assertEqual(render.get_column(ETree.Element('c', r='ABC123')), 'ABC')

    def test_update_cell(self):
        cell = ETree.Element('c', t='n', r='A1')
        sub_elem = ETree.SubElement(cell, 'v')
        sub_elem.text = '123'
        render.update_cell(cell, line=30, value=125)
        self.assertEqual(cell.get('r'), 'A30')
        self.assertEqual(cell.get('t'), 'n')
        self.assertEqual(sub_elem.text, '125')
        render.update_cell(cell, line=30, value=None)
        self.assertEqual(sub_elem.text, '')

        cell = ETree.Element('c', r='A1')
        sub_elem = ETree.SubElement(cell, 'v')
        sub_elem.text = '123'
        render.update_cell(cell, line=30, value=125)
        self.assertEqual(cell.get('r'), 'A30')
        self.assertEqual(cell.get('t'), None)
        self.assertEqual(sub_elem.text, '125')
        render.update_cell(cell, line=30, value=None)
        self.assertEqual(sub_elem.text, '')

        cell = ETree.Element('c', t='b', r='B1')
        sub_elem = ETree.SubElement(cell, 'v')
        sub_elem.text = '1'
        render.update_cell(cell, line=12, value=False)
        self.assertEqual(cell.get('r'), 'B12')
        self.assertEqual(cell.get('t'), 'b')
        self.assertEqual(sub_elem.text, '0')
        render.update_cell(cell, line=12, value=None)
        self.assertEqual(sub_elem.text, '')

        cell = ETree.Element('c', t='s', r='C1')
        sub_elem = ETree.SubElement(ETree.SubElement(cell, 'is'), 't')
        sub_elem.text = 'Ok'
        render.update_cell(cell, line=2, value='Updated')
        self.assertEqual(cell.get('r'), 'C2')
        self.assertEqual(cell.get('t'), 'inlineStr')
        element_iter = cell.iter()
        next(element_iter)
        self.assertEqual(next(element_iter).tag, 'is')
        child = next(element_iter)
        self.assertEqual(child.tag, 't')
        self.assertEqual(child.text, 'Updated')
        render.update_cell(cell, line=2, value=None)
        self.assertEqual(child.text, '')

    def gen_row(self):
        row = ETree.Element('row', r='1')
        first_cell = ETree.SubElement(row, 'c', t='n', r='A1')
        ETree.SubElement(first_cell, 'v').text = '12'
        second_cell = ETree.SubElement(row, 'c', t='s', r='B1')
        ETree.SubElement(second_cell, 'v').text = 'second'
        third_cell = ETree.SubElement(row, 'c', r='C1')
        ETree.SubElement(third_cell, 'v').text = '21'
        return row

    def test_render_row(self):
        template_row = self.gen_row()
        row = render.render_row([42, 'Noé!>', 24], template_row, 12)
        ETree.fromstring(row)
        expected = (
            '<row r="12">'
            '<c t="n" r="A12"><v>42</v></c>'
            '<c t="inlineStr" r="B12"><is><t>Noé!&gt;</t></is></c>'
            '<c r="C12"><v>24</v></c>'
            '</row>'.encode()
        )
        self.assertEqual(row, expected)


    def test_render_row_wrong_template(self):
        template_row = self.gen_row()
        row = render.render_row([42, 'Noé!>', 24, 'NoTemplateElement'], template_row, 1)
        ETree.fromstring(row)
        self.assertEqual(
            row,
            '<row r="1">'
            '<c r="A1" t="inlineStr"><is><t>42</t></is></c>'
            '<c r="B1" t="inlineStr"><is><t>Noé!&gt;</t></is></c>'
            '<c r="C1" t="inlineStr"><is><t>24</t></is></c>'
            '<c r="D1" t="inlineStr"><is><t>NoTemplateElement</t></is></c>'
            '</row>'.encode()
        )

    def test_render_row_null_template(self):
        template_row = None
        row = render.render_row([42, 'Noé!>', 24], template_row, 2)
        ETree.fromstring(row)
        self.assertEqual(
            row,
            '<row r="2">'
            '<c r="A2" t="inlineStr"><is><t>42</t></is></c>'
            '<c r="B2" t="inlineStr"><is><t>Noé!&gt;</t></is></c>'
            '<c r="C2" t="inlineStr"><is><t>24</t></is></c>'
            '</row>'.encode()
        )

    def test_render_row_invalid_xml_char(self):
        row = render.render_row(["foo\x02bar", "_x0002_"], None, 2)
        ETree.fromstring(row)
        self.assertEqual(
            row,
            '<row r="2">'
            '<c r="A2" t="inlineStr"><is><t>foo_x0002_bar</t></is></c>'
            '<c r="B2" t="inlineStr"><is><t>_x005F_x0002_</t></is></c>'
            "</row>".encode(),
        )

    def test_render_rows(self):
        template_row = self.gen_row()
        rows, lines = render.render_rows([[42, 'Noé!>', 24], [18, '<éON', 21]], template_row, 1)
        expected = (
            '<row r="1">'
            '<c t="n" r="A1"><v>42</v></c>'
            '<c t="inlineStr" r="B1"><is><t>Noé!&gt;</t></is></c>'
            '<c r="C1"><v>24</v></c>'
            '</row>\n'
            '<row r="2">'
            '<c t="n" r="A2"><v>18</v></c>'
            '<c t="inlineStr" r="B2"><is><t>&lt;éON</t></is></c>'
            '<c r="C2"><v>21</v></c>'
            '</row>'.encode()
        )
        self.assertEqual(rows, expected)
        self.assertEqual(lines, 2)

    def _verify_sheet(self, data):
        document = b''
        xlsx_doc = render.render_worksheet(data, gen_xlsx_sheet())
        stream = next(xlsx_doc)
        document += stream
        self.assertTrue(stream.startswith(b'<worksheet'))
        stream = next(xlsx_doc)
        document += stream
        self.assertTrue(stream.startswith(b'<sheetData'))
        stream = next(xlsx_doc)
        document += stream
        self.assertTrue(stream.startswith(b'<row'))
        stream = next(xlsx_doc)
        document += stream
        self.assertTrue(stream.startswith(b' </sheetData'))

        tree_iter = ETree.fromstring(document).iter()
        self.assertEqual(next(tree_iter).tag, '{%s}worksheet' % render.OPENXML_NS)
        self.assertEqual(next(tree_iter).tag, '{%s}sheetData' % render.OPENXML_NS)
        self.assertEqual(next(tree_iter).tag, '{%s}row' % render.OPENXML_NS)
        self.assertEqual(next(tree_iter).tag, '{%s}c' % render.OPENXML_NS)

    def test_render_worksheet(self):
        data = [[[42, 'Noé!>', 24], [18, '<éON', 21]]]
        self._verify_sheet(data)

        # render worksheet with an iterator
        data = iter([iter([[42, 'Noé!>', 24], [18, '<éON', 21]])])
        self._verify_sheet(data)
