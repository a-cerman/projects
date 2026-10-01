import tempfile
import unittest
import zipfile
import os
import subprocess
import sys
from pathlib import Path
from xml.sax.saxutils import escape
from app import parse_docx, validate

class ParserTests(unittest.TestCase):
    def test_utf8_child_input(self):
        text = 'Äpfel, Öl, Überprüfung, Grüße und Straße.'
        env = os.environ.copy()
        env.update(PYTHONIOENCODING='utf-8', PYTHONUTF8='1')
        result = subprocess.run(
            [sys.executable, '-X', 'utf8', '-c', 'import sys; sys.stdout.write(sys.stdin.read())'],
            input=text, text=True, encoding='utf-8', capture_output=True, env=env, check=True)
        self.assertEqual(result.stdout, text)

    def document(self, lines):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        path = Path(temporary.name) / 'test.docx'
        xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        xml += ''.join('<w:p><w:r><w:t>' + escape(line) + '</w:t></w:r></w:p>' for line in lines)
        xml += '</w:body></w:document>'
        with zipfile.ZipFile(path, 'w') as z:
            z.writestr('word/document.xml', xml)
        return path

    def test_german_and_heading_variations(self):
        actual = parse_docx(self.document(['FOLIE - 1.', 'Grüße & Ergebnisse', 'Zweite Zeile', 'folie–2', 'Vielen Dank.']))
        self.assertEqual(actual, {1: 'Grüße & Ergebnisse\nZweite Zeile', 2: 'Vielen Dank.'})

    def test_rejected_documents(self):
        for lines in (['FOLIE - 1.', 'Text', 'FOLIE - 1.', 'Text'], ['FOLIE - 1.'], ['Titel', 'FOLIE - 1.', 'Text'], ['FOLIE - 0.', 'Text']):
            with self.subTest(lines=lines), self.assertRaises(ValueError):
                parse_docx(self.document(lines))

    def test_exact_slide_matching(self):
        validate({1: 'a', 2: 'b'}, 2)
        for mapping in ({1: 'a'}, {1: 'a', 3: 'b'}):
            with self.assertRaises(ValueError):
                validate(mapping, 2)

if __name__ == '__main__':
    unittest.main()
