import uno
import unohelper

from com.sun.star.task import XJobExecutor


class Main(unohelper.Base, XJobExecutor):
    def __init__ (self, ctx):
        self.ctx = ctx

    def trigger(self, cmd):
        desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        model = desktop.getCurrentComponent()
        text = model.Text
        cursor = text.createTextCursor()
        text.insertString(cursor, "Hello World! \n{}".format(cmd), 0)


# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(Main, 'org.chaffalo.Main', ('com.sun.star.task.Job',),)
