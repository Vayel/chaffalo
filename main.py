import os

import uno
import unohelper

from com.sun.star.task import XJobExecutor

import config


class Main(unohelper.Base, XJobExecutor):
    def __init__ (self, ctx):
        self.ctx = ctx
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop",
            self.ctx
        )
        self.db_ctx = self.smgr.createInstanceWithContext(
            "com.sun.star.sdb.DatabaseContext",
            self.ctx
        )
        self.dispatcher = self.smgr.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper",
            self.ctx
        )

    @staticmethod
    def log(document, txt):
        text = document.Text
        cursor = text.createTextCursor()
        text.insertString(cursor, txt + "\n", 0)

    @classmethod
    def get_document_path(cls, document):
        return uno.fileUrlToSystemPath(document.URL)

    @classmethod
    def get_document_name(cls, document):
        return os.path.splitext(os.path.basename(cls.get_document_path(document)))[0]

    def update_indexes(self, document):
        self.dispatcher.executeDispatch(document.getCurrentController(),
                                        ".uno:UpdateAll", "", 0, ())

    def print_document(self, document):
        """
        if path.endswith("PC.odt"):
            self.insert_project_manager_data(document)
            self.insert_methodology(document)
        """
        
        stg = document.createInstance("com.sun.star.text.DocumentSettings")
        db_name = stg.CurrentDatabaseDataSource
        table_name = stg.CurrentDatabaseCommand
        output_folder = os.path.join(
            os.path.dirname(self.get_document_path(document)),
            config.PDF_DIRNAME
        )

        self.update_indexes(document)

        # Fill field from db and export as PDF
        mail_merge = self.smgr.createInstanceWithContext(
            "com.sun.star.text.MailMerge",
            self.ctx
        )
        mail_merge.DocumentURL = document.URL
        mail_merge.DataSourceName = db_name
        mail_merge.CommandType = uno.getConstantByName("com.sun.star.sdb.CommandType.TABLE")
        mail_merge.Command = table_name
        mail_merge.OutputType = uno.getConstantByName("com.sun.star.text.MailMergeType.FILE")
        mail_merge.OutputURL = uno.systemPathToFileUrl(output_folder)
        mail_merge.SaveFilter = "writer_pdf_Export"
        mail_merge.SaveFilterData = tuple([])
        mail_merge.execute(())

        document.store()

        # Rename PDF file
        pdf_path = os.path.join(output_folder, self.get_document_name(document) + "0.pdf")
        os.rename(pdf_path, pdf_path.replace("0.pdf", ".pdf"))

    def trigger(self, cmd):
        model = self.desktop.getCurrentComponent()

        if cmd == "print":
            self.print_document(model)
        else:
            raise ValueError("Unknown command '{}'".format(cmd))


# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(Main, 'org.javelot.Main', ('com.sun.star.task.Job',),)
