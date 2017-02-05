import unittest
from xml.etree import ElementTree as ET
from checker import get_desktop, get_active_model
import subprocess
import os
import time
import uno
from com.sun.star.beans import PropertyValue
from GrammarAlert import run_check, decode_errors, \
    check_sentence, markup_errors, select_part


@unittest.skip('not testing xml')
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
        self.desktop = get_desktop()
        self.model = get_active_model(self.desktop)

    def test_markup(self):
        tenum = self.model.Text.createEnumeration()
        # the third sentence of test_doc.odt has errors
        testtext = tenum.nextElement()
        testtext = tenum.nextElement()
        testtext = tenum.nextElement()
        sentence = testtext.getString()
        errors = check_sentence(sentence)
        erroot = errors.getroot()
        markups = markup_errors(sentence, erroot)
        for m in markups:
            mtext = next(m[2].iter(tag='string')).text
            self.assertEqual(sentence[m[0]:m[1]], mtext)

    @classmethod
    def XtearDownClass(self):
        self.model.dispose()
        if self.desktop is not None:
            self.desktop.dispose()
        if self.soffice is not None:
            self.soffice.terminate()
            self.soffice.communicate()

if __name__=='__main__':
    unittest.main()
