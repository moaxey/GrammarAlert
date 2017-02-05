"""
Grammarchecker runs each sentence through an atd server,
highlighting errors

Tests can be run using the python which comes with a recent
Libreoffice

GrammarAlert() must be copied and run as a macro

ATD server must run on 127.0.0.1 and it is not secure
to use an externally hosted service.

"""
from urllib import request, parse
from urllib.error import URLError
from xml.etree import ElementTree as ET
import sys
import re

ATD_SERVER = 'http://127.0.0.1:1049'
QUERY = 'checkDocument?data={}'
# cursor movement
EXTEND = True
MOVE = False

def select_part(cur, part):
    selected = cur.getString()
    while selected != part[1]:
        cur.gotoNextWord(EXTEND)
        selected = cur.getString()
    return cur

def run_check(sentence):
    sentence = parse.quote_plus(sentence)
    url="{}/{}".format(
        ATD_SERVER,
        QUERY.format(sentence)
    )
    req = request.Request(
        url = url
    )
    try:
        f = request.urlopen(req)
    except URLError as e:
        raise URLError(
            'Unable to connect to local atd-server ({}). Is it running'.format(
                ATD_SERVER
            )
        )
    return f.read().decode('utf-8')

def decode_errors(response_data):
    results = ET.ElementTree(
        ET.fromstring(
            response_data
        )
    )
    return results

def check_sentence(sentence):
    response = run_check(sentence)
    return decode_errors(response)

def markup_errors(sentence, erroot):
    """ markup an ordered list of the errors: start, end, error """
    markups = []
    for error in erroot:
        errstring = next(error.iter(tag='string')).text
        match = re.search(errstring, sentence)
        if match is None:
            raise RuntimeError('String "{}" not found in "{}"'.format(
                errstring, sentence
            ))
        mend = 0
        while match.span() in [(n[0], n[1]) for n in markups]:
            mend += match.end()
            match = re.search(errstring, sentence[mend:])
        markups.append((match.start() + mend, match.end() + mend, error))
        markups = sorted(markups, key=lambda x: x[0])
    return markups

def GrammarAlert():
    """ Highlight grammar errors """
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    if not hasattr(model, "Text"):
        return None
    for t in model.Text.createEnumeration():
        sentence = t.getString()
        errors = check_sentence(sentence)
        erroot = errors.getroot()
        if len(erroot) > 0:
            markups = markup_errors(sentence, erroot)
            cur = t.createTextCursor()
            """
            # can add hyperlink to cursor
            # can set PropertyValue
            cur.gotoNextWord(True)
            cur.getString()
            cur.HyperLinkURL = 'http://localhost:1049/'
            cur.setPropertyValue("CharBackColor", int("ffffbf", 16))
            """
            for part in parts:
                cur = select_part(cur, part)
                
                
            t.setPropertyValue(
                "CharBackColor", int("ffffbf", 16)
            )
        else:
            t.setPropertyToDefault("CharBackColor")
    tRange = model.Text.End
    tRange.String = "GrammarAlert ran"
    return None

"""
model.getCurrentController()
cursor = model.getText().createTextCursor()
cur.gotoNextWord(True)
cur.getString()

tRange = text.End
tRange.String = “Hello World (in Python)”
tRange.setPropertyValue("CharBackColor", int("ffffbf", 16))

"""
