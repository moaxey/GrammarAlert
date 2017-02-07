GrammarAlert
============

A simple macro for LibreOffice to highlight errors with suggestions and links to
explanations provided by the open source [After the deadline
(ATD)](https://open.afterthedeadline.com/) language checking service. This macro
is hard-coded to use that server on localhost.

There is an unmaintained java-based LibreOffice extension which I'm sure is
better than this software, but does not work with LibreOffice 5+ called
[atd-openoffice](https://github.com/Automattic/atd-openoffice.git).

Tested in LibreOffice 5.1 and 5.3dev.

Installation
------------

Copy the GrammarAlert.py script to your LibreOffice python scripts folder. The
site varies per platform. For example, OSX and in recent nightly build:

>   /Applications/LibreOfficedev.app/Contents/Resources/Scripts/python/

Usage
-----

The macro will appear in:

>   Tools > Macros > Organise Macros > Python

GrammarAlert will appear under 'LibreOffice Macros'. Run the check function to
check the current document.

Development
-----------

-   Use the LibreOffice python to run the tests.

-   Run ATD service on localhost.

-   Run LibreOffice in server mode, listening on port 2002.

### Ideas for improvements

-   Highlighting with non-destructive underlines like spellcheck.

-   Providing suggestions in a contextual menu instead of mangling the text.

### Known limitations

-   ATD service only supports English US. ATD spelling errors not highlighted to
    give compatibility with other English variants.

-   ATD service ignores curly quotes and reports them as errors because they are
    missing. There is a simple filter for them included to replace them with
    their ASCII equivalents. Other extended characters are likely to cause
    similar problems.

