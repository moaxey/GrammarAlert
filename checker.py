import uno

"""
>>> from checker import *
>>> desktop = get_desktop()
>>> model = get_active_model(desktop)

"""

def get_desktop():
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
	"com.sun.star.bridge.UnoUrlResolver", localContext )
    ctx = resolver.resolve(
        "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx
    )
    return desktop

def get_active_model(desktop):
    desktop.setActiveFrame(desktop.getFrames()[0])
    return desktop.getCurrentComponent()

