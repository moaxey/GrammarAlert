import unittest
from xml.etree import ElementTree as ET
from checker import get_desktop, get_active_model
import subprocess
import os
import time
from GrammarAlert import run_check, decode_errors, \
    check_sentence, segment

class ATD_XML_Test(unittest.TestCase):

    def test_run_check(self):
        sentence = '“This is a sentence”'
        response = run_check(sentence)
        self.assertEqual(
            response,
            """<results>
</results>"""
        )
        errors = decode_errors(response)
        erroot = errors.getroot()
        self.assertEqual(
            len(erroot), 0
        )
        cerrors = check_sentence(sentence)
        cerroot = cerrors.getroot()
        self.assertEqual(
            ET.tostring(erroot), ET.tostring(cerroot)
        )

    def test_run_check_segment(self):
        sentence = 'It was determined by the committee that the report was inconclusive'
        response = run_check(sentence)
        self.assertEqual(
            response,
            """<results>
  <error>
    <string>was determined by the</string>
    <description>Passive voice</description>
    <precontext>It</precontext>

    <type>grammar</type>
    <url>http://service.afterthedeadline.com/info.slp?text=was+determined+by+the&amp;tags=VBD%2FVBN%2FIN%2FDT&amp;engine=3</url>

  </error>
</results>"""
        )
        errors = decode_errors(response)
        erroot = errors.getroot()
        self.assertEqual(
            len(erroot), 1
        )
        for e in erroot:
            self.assertEqual(e.tag, 'error')
            if e.tag == 'error':
                for t in e:
                    self.assertTrue(
                        t.tag in (
                            'string',
	                    'description',
	                    'precontext',
	                    'type',
	                    'url'
                        )
                    )
        segments = segment(sentence, erroot)
        self.assertEqual(
            segments[0],
            (None, 'It ', None)
        )
        self.assertEqual(
            segments[1],
            ('grammar', 'was determined by the', 'http://127.0.0.1:1049/info.slp?text=was+determined+by+the&tags=VBD%2FVBN%2FIN%2FDT&engine=3')
        )
        self.assertEqual(
            segments[2],
            (None,
             ' committee that the report was inconclusive',
             None
            )
        )

class LO_Cursor_Test(unittest.TestCase):

    @classmethod
    def setUpClass(self):
        if 'SOFFICE_BINARY' not in os.environ.keys():
            raise RuntimeError('Please set environment variable SOFFICE_BINARY to the path to your soffice executable to run these tests.')
        self.soffice = subprocess.Popen(
            '{} --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo --nodefault --headless "test_doc.odt"'.format(os.environ['SOFFICE_BINARY']),
            shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        self.desktop = None
        start = time.time()
        while self.desktop is None and time.time() - start < 30:
            print(time.time() - start)
            try:
                self.desktop = get_desktop()
            except Exception as e:
                print(e)
                time.sleep(5)
        if self.desktop is None:
            raise RuntimeError('Unable to connect to soffice')
        self.model = get_active_model(self.desktop)

    def test_segments_match_cursor(self):
        print('sm', self.model)

    @classmethod
    def tearDownClass(self):
        if not self.soffice.poll():
            self.soffice.terminate()

if __name__=='__main__':
    unittest.main()
