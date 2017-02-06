import unittest
from xml.etree import ElementTree as ET
from checker import get_desktop, get_active_model
import subprocess
import os
import time
import uno
from com.sun.star.beans import PropertyValue
from GrammarAlert import run_check, decode_errors, markup_document, \
    check_sentence, markup_errors, select_markup, highlight_error

#@unittest.skip('not testing xml')
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
        self.assertEqual(
            len(errors), 0
        )
        cerrors = check_sentence(sentence)
        self.assertEqual(
            errors, cerrors
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
        self.assertEqual(
            len(errors), 1
        )
        for e in errors:
            self.assertEqual(type(e).__name__, 'dict')
            for t in e.keys():
                self.assertTrue(
                    t in (
                        'string',
	                'description',
	                'precontext',
	                'type',
	                'url',
                        'options',
                    )
                )
        markups = markup_errors(sentence, errors)
        self.assertEqual(
            markups[0][:2],
            (3, 24)
        )
        expected_values = {
            'type': 'grammar',
            'string': 'was determined by the',
            'options': [],
            'description': 'Passive voice',
            'url': 'http://127.0.0.1:1049/info.slp?text=was+determined+by+the&tags=VBD%2FVBN%2FIN%2FDT&engine=3'
        }
        self.assertEqual(
            sorted(markups[0][2].keys()),
            sorted(expected_values.keys())
        )
        for key in expected_values.keys():
            self.assertEqual(
                markups[0][2][key],
                expected_values[key]
            )

#@unittest.skip('not testing cursor')
class LO_Cursor_Test(unittest.TestCase):

    @classmethod
    def setUpClass(self):
        self.desktop = get_desktop()
        self.model = get_active_model(self.desktop)

    def test_markup(self):
        markup_document(self.desktop, self.model)


if __name__=='__main__':
    unittest.main()
