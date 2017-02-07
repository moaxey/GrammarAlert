GrammarAlert
------------

A simple macro for Libreoffice to highlight errors with suggestions and links
to explanations provided by the open source [After the deadline (ATD)](https://open.afterthedeadline.com/ "Install the After the Deadline server") language
checkin service. This macro is hard-coded to use an instance of that server on localhost.

There is an unmaintained java-based LibreOffice extension which I'm sure is
better than this software, but does not work with Libroffice 5+.
[atd-openoffice](https://github.com/Automattic/atd-openoffice.git "atd-openoffice on github"))

Macro has been tested in LibreOffice 5.1 and 5.3dev.


Installation
============

Copy the GrammarAlert.py script to your Libreoffice python scripts folder.
The location varies per platform. For example, OSX and in recent nightly build:

> /Applications/LibreOfficedev.app/Contents/Resources/Scripts/python/


Usage
=====

The macro will appear in:

> Tools > Macros > Organise Macros > Python

GrammarAlert will appear under 'LibreOffice Macros'. Run the check function
to check the current document.


Development
===========

Run the test script (based on unittest) using the python which is bundled
with your copy of libreoffice.

- ATD service must be running on localhost
- Libreoffice must be running in server mode, listening on port 2002

### Ideas for improvements

- Highlighting with non-destructive underlinse like spellcheck.
- Providing suggestions in a contextual menu instead of mangling the text

### Known limitations

- ATD service only supports English US so its spelling errors are suppressed here to provide compatability with other English variants. 
- ATD service ignores curley quotes and reports them as errors because they
are missing. There is a simple filter for them included to replace them with
their ASCII equivalents. Other extended characters are likely to cause similar
problems.

