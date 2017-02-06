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

COLOURS = {
    'warning': int("ffffbf", 16),
    'spelling': int("ffbcbc", 16),
    'grammar': int("b7f39f", 16),
    'style': int("c0c7dc", 16),
    'suggestion': int("b1dbcb", 16),
}

def select_markup(text, markup, sentence_start):
    cursor = text.Text.createTextCursor()
    cursor.goRight(markup[0] + sentence_start, MOVE)
    cursor.goRight(markup[1] - markup[0], EXTEND)
    return cursor

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
    errors = []
    results = ET.ElementTree(
        ET.fromstring(
            response_data
        )
    )
    results = results.getroot()
    for error in results:
        errdict = {}
        errdict['string'] = next(error.iter(tag='string')).text
        errdict['type'] = next(error.iter(tag='type')).text
        errdict['description'] = next(error.iter(tag='description')).text
        try:
            murl = next(error.iter(tag='url'))
            errdict['url'] = murl.text.replace(
                'http://service.afterthedeadline.com/',
                'http://127.0.0.1:1049/'
            )
        except StopIteration:
            errdict['ul'] = None
        try:
            msugg = next(error.iter(tag='suggestions'))
            errdict['options'] = [n.text for n in msugg.iter(tag='option')]
        except StopIteration:
            errdict['options'] = []
        errors.append(errdict)
    return errors

def check_sentence(sentence):
    response = run_check(sentence)
    return decode_errors(response)

def markup_errors(sentence, errdict):
    """ markup an ordered list of the errors: start, end, error """
    markups = []
    for error in errdict:
        match = re.search(error['string'], sentence)
        if match is None:
            raise RuntimeError('String "{}" not found in "{}"'.format(
                error['string'], sentence
            ))
        mend = 0
        while match.span() in [(n[0], n[1]) for n in markups]:
            mend += match.end()
            match = re.search(error['string'], sentence[mend:])
        markups.append((match.start() + mend, match.end() + mend, error))
        markups = sorted(markups, key=lambda x: x[0])
    return markups

def highlight_error(cursor, error):
    if error['type'] == 'spelling':
        return None
    cursor.setPropertyValue("CharBackColor", COLOURS[error['type']])
    if 'url' in error.keys() and error['url'] is not None:
        cursor.HyperLinkURL = error['url']
    insert_at = cursor.End
    if 'options' in error.keys() and len(error['options']) > 0:
        insert_at = insert_at.End
        insert_at.setString("<<{}>>".format(', '.join(error['options'])))

def simplify_characters_for_url(string):
    string = string.replace('’', "'")
    string = re.sub(r'[”“]', '"', string)
    return string

def markup_document(desktop, model):
    if not hasattr(model, "Text"):
        return None
    alltext = model.Text.getString()
    sentences = alltext.split('.')
    for sentence in sentences[::-1]:
        sstart = alltext.find(sentence)
        ssentence = simplify_characters_for_url(sentence)
        errors = check_sentence(ssentence)
        if len(errors) > 0:
            markups = markup_errors(ssentence, errors)
            for markup in markups[::-1]:
                cursor = select_markup(model.Text, markup, sstart)
                highlight_error(cursor, markup[2])

def GrammarAlert():
    """ Highlight grammar errors """
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    markup_document(desktop, model)
