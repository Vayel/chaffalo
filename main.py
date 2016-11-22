import os

import uno
import unohelper

from com.sun.star.task import XJobExecutor
from com.sun.star.beans import UnknownPropertyException

import config
import messages
import tools


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

    def update_indexes(self, document):
        self.dispatcher.executeDispatch(document.getCurrentController(),
                                        ".uno:UpdateAll", "", 0, ())

    def open_document(self, path, hidden=True):
        struct = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        struct.Name = "Hidden"
        struct.Value = hidden

        return self.desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(path),
            "_blank",
            0,
            tuple([struct])
        )
    
    def update_sections_link(self, document, dst_path):
        """For each linked section of src, replace the filename by dst_path"""

        sections = document.getTextSections()
        for i in range(sections.getCount()):
            section = sections.getByIndex(i)

            if not section.FileLink.FileURL:
                continue

            section.FileLink.FileURL = uno.systemPathToFileUrl(dst_path)

        document.store()

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
            os.path.dirname(tools.get_document_path(document)),
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
        pdf_path = os.path.join(output_folder, tools.get_document_name(document) + "0.pdf")
        os.rename(pdf_path, pdf_path.replace("0.pdf", ".pdf"))

    def register_db(self, ods_path, db_path, db_name):
        if os.path.isfile(db_path):
            return

        if self.db_ctx.hasByName(db_name):
            raise ValueError("The database {} already exists.".format(db_name))

        db_folder_path = os.path.dirname(ods_path)
        db_url = uno.systemPathToFileUrl(db_path)
        db_folder_url = uno.systemPathToFileUrl(db_folder_path)
         
        data_src = self.smgr.createInstanceWithContext(
            "com.sun.star.sdb.DataSource",
            self.ctx
        )
        data_src.URL = "sdbc:calc:" + ods_path  
        data_src.DatabaseDocument.storeAsURL(db_url, ()) 
     
        self.db_ctx.registerObject(db_name, data_src)
        data_src.DatabaseDocument.close(True) 

    def link_db(self, document, db_name, table_name):
        # Update db fields url
        fields_enum = document.getTextFields().createEnumeration()

        while fields_enum.hasMoreElements():
            master_field = fields_enum.nextElement().getTextFieldMaster()

            try:
                master_field.setPropertyValue("DataBaseName", db_name)
                master_field.setPropertyValue("DataTableName", table_name)
            except UnknownPropertyException:
                pass

        # Attach the db to the document
        stg = document.createInstance("com.sun.star.text.DocumentSettings")
        stg.setPropertyValue("CurrentDatabaseDataSource", db_name)
        stg.setPropertyValue("CurrentDatabaseCommand", table_name)
        stg.setPropertyValue(
            "CurrentDatabaseCommandType",
            uno.getConstantByName("com.sun.star.sdb.CommandType.TABLE")
        )

        document.store()

    def configure_project(self, document):
        folder = tools.get_document_folder(document)
        ods = tools.get_ods_fname(folder)
        db_name_base = db_name = ods[:-4]
        n = 0

        while True:
            try:
                self.register_db(
                    os.path.join(folder, ods),
                    os.path.join(folder, db_name + ".odb"),
                    db_name
                )
            except ValueError:
                n += 1
                db_name = db_name_base + "-" + str(n)
            else:
                break

        for tmpl in tools.get_odt_fnames(folder):
            odt = self.open_document(os.path.join(folder, tmpl))

            self.link_db(odt, db_name, tmpl[:-4])
            self.update_sections_link(odt, os.path.join(folder, config.CONTENT_FNAME))

            odt.dispose()

    def trigger(self, cmd):
        model = self.desktop.getCurrentComponent()

        try:
            if cmd == "print":
                self.print_document(model)
            if cmd == "configure":
                self.configure_project(model)
            else:
                raise ValueError("Unknown command '{}'".format(cmd))
        except Exception as e:
            messages.error(
                model.CurrentController.Frame.ContainerWindow,
                str(e),
                "System error"
            )


# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(Main, 'org.javelot.Main', ('com.sun.star.task.Job',),)
