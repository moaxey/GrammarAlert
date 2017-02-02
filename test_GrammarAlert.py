import unittest
from xml.etree import ElementTree as ET

from GrammarAlert import run_check, decode_errors, \
    check_sentence

class GTest(unittest.TestCase):
    
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
        

    def test_run_check_longer(self):
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
        cerrors = check_sentence(sentence) 
        cerroot = cerrors.getroot()
        self.assertEqual(
            ET.tostring(erroot), ET.tostring(cerroot)
        )
        

if __name__=='__main__':
    unittest.main()
