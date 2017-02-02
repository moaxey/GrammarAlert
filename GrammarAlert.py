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
from xml.etree import ElementTree as ET
import sys

ATD_SERVER = 'http://127.0.0.1:1049'
QUERY = 'checkDocument?data={}'

def run_check(sentence):
    sentence = parse.quote_plus(sentence)
    url="{}/{}".format(
        ATD_SERVER,
        QUERY.format(sentence)
    )
    req = request.Request(
        url = url
    )
    f = request.urlopen(req)
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

def GrammarAlert():
    """ Highlight grammar errors """
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    if not hasattr(model, "Text"):
        return None
    for t in model.Text.createEnumeration():
        errors = check_sentence(t.getString())
        erroot = errors.getroot()
        if len(erroot) > 0:
            t.setPropertyValue(
                "CharBackColor", int("ffffbf", 16)
            )
        else:
            t.setPropertyToDefault("CharBackColor")    
    tRange = model.Text.End
    tRange.String = "GrammarAlert ran"
    return None
