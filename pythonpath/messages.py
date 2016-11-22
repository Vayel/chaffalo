import uno

from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.awt.MessageBoxType import ERRORBOX, MESSAGEBOX


def message(parent, text, title, type=MESSAGEBOX, buttons=BUTTONS_OK):
    ctx = uno.getComponentContext()
    sm = ctx.ServiceManager
    sv = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx) 
    box = sv.createMessageBox(parent, type, buttons, title, text)
    
    return box.execute()


def error(parent, text, title):
    return message(parent, text, title, type=ERRORBOX, buttons=BUTTONS_OK)
