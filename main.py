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
    
    def update_sections_link(self, document, dst_dir):
        """For each linked section of src, replace the filename by dst_path"""

        sections = document.getTextSections()
        for i in range(sections.getCount()):
            section = sections.getByIndex(i)

            if not section.FileLink.FileURL:
                continue

            url = uno.systemPathToFileUrl(os.path.join(
                dst_dir,
                os.path.basename(section.FileLink.FileURL)
            ))
            section.FileLink.FileURL = url

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

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

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
        new_path = pdf_path.replace("0.pdf", ".pdf")
        os.rename(pdf_path, new_path)

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

        # Look for an unused db name
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
            already_opened = tools.get_document_fname(document) == tmpl
            odt = document if already_opened else self.open_document(os.path.join(folder, tmpl))

            self.link_db(odt, db_name, tmpl[:-4])
            self.update_sections_link(odt, folder)

            if not already_opened:
                odt.dispose()
        
        messages.message(
            document.CurrentController.Frame.ContainerWindow,
            str("Les modèles ont été configurés avec succès."),
            "Configuration terminée"
        )

    def update_gantt(self, document):
        folder = tools.get_document_folder(document)
        calc = self.open_document(os.path.join(folder, tools.get_ods_fname(folder)))
    
        LAST_ROW_NUMBER_CELL = "K2"
        TIME_UNIT_CELL = "Q2"
        DURATION_CELL = "O2"
        CHART_NAME = "Gantt"
        TIME_STEP = 1

        sheet = calc.getSheets().getByName("PC")
        time_axis_max = int(sheet.getCellRangeByName(
            DURATION_CELL
        ).getValue())
        time_axis_unit = sheet.getCellRangeByName(
            TIME_UNIT_CELL
        ).getString()
        time_axis_title = "Durée ({})".format(time_axis_unit)

        sheet = calc.getSheets().getByName("Phases")
        last_row_number = int(sheet.getCellRangeByName(
            LAST_ROW_NUMBER_CELL
        ).getValue())
        durations_range = sheet.getCellRangeByName("E2:F{}".format(last_row_number))
        names_range = sheet.getCellRangeByName("A2:A{}".format(last_row_number))

        data = durations_range.getDataArray()
        descriptions = tuple([t[0] for t in names_range.getDataArray()])

        chart = document.getEmbeddedObjects().getByName(CHART_NAME).getEmbeddedObject()
        chart.getDiagram().getYAxis().setPropertyValue("Max", time_axis_max)
        chart.getDiagram().getYAxis().setPropertyValue("StepMain", TIME_STEP)
        chart.getData().setData(data)
        chart.getData().setRowDescriptions(descriptions)
        chart.getDiagram().getYAxis().AxisTitle.String = time_axis_title

        calc.dispose()

    def update_budget_table(self, document):
        folder = tools.get_document_folder(document)
        calc = self.open_document(os.path.join(folder, tools.get_ods_fname(folder)))

        SHEET_NAME = "Phases"
        LAST_ROW_NUMBER_CELL = "J2"
        LAST_COLUMN_NAME = "D"
        SECTION_NAME = "BudgetEtude"
        TABLE_NAME = "TableauBudgetEtude"
        COL_SEP_POS_DELTAS = [900, -500, 0]

        # Get data
        sheet = calc.getSheets().getByName(SHEET_NAME)
        last_row_number = int(sheet.getCellRangeByName(
            LAST_ROW_NUMBER_CELL
        ).getValue())
        data = sheet.getCellRangeByName("A1:{}{}".format(
            LAST_COLUMN_NAME,
            last_row_number
        ))
        
        # Create table
        if document.getTextTables().hasByName(TABLE_NAME):
            table = document.getTextTables().getByName(TABLE_NAME)
            table.setName(TABLE_NAME + "_")
            table.dispose()

        section = document.getTextSections().getByName(SECTION_NAME)
        anchor = section.getAnchor()

        table = document.createInstance("com.sun.star.text.TextTable")
        table.setName(TABLE_NAME)
        rows = data.getRows()
        columns = data.getColumns()
        row_count = rows.getCount()
        col_count = columns.getCount()
        table.initialize(row_count, col_count)

        anchor.getText().insertTextContent(anchor, table, False)
        
        # Format text
        for i in range(row_count):
            for j in range(col_count):
                val = data.getCellByPosition(j, i).getText().getString()
                cell = table.getCellByPosition(j, i)
                cell.setString(val)

                if i >= row_count - 2 and j == 0:
                    tools.superscript_cell_note(cell)

                if i == 0 or i == row_count - 1 and j == 0:
                    tools.bold_cell(cell)

                if val == "":
                    tools.hide_border(cell)

                tools.align_cell_text(cell, row_count, i, j)

        # Resize columns
        pos = [
            sep.value.Position + delta
            for sep, delta in zip(table.TableColumnSeparators, COL_SEP_POS_DELTAS)
        ]
        seps = [
            uno.createUnoStruct("com.sun.star.text.TableColumnSeparator", p, True)
            for p in pos
        ]
        table.TableColumnSeparators = tuple(seps)

        calc.dispose()
    
    def update_methodology_table(self, document):
        folder = tools.get_document_folder(document)
        calc = self.open_document(os.path.join(folder, tools.get_ods_fname(folder)))

        tools.update_phase_table(
            document,
            calc,
            "TableauMéthodologie",
            "VieuxTableauMéthodologie",
            2
        )

        calc.dispose()

    def update_feature_table(self, document):
        folder = tools.get_document_folder(document)
        calc = self.open_document(os.path.join(folder, tools.get_ods_fname(folder)))

        tools.update_phase_table(
            document,
            calc,
            "TableauFonctionnalités",
            "VieuxTableauFonctionnalités",
            2,
            ["Table Contents", "Fonctionnalités"]
        )

        calc.dispose()
   
    def trigger(self, cmd):
        model = self.desktop.getCurrentComponent()

        try:
            if cmd == "print":
                self.print_document(model)
            elif cmd == "configure":
                self.configure_project(model)
            elif cmd == "updateGantt":
                self.update_gantt(model)
            elif cmd == "updateBudgetTable":
                self.update_budget_table(model)
            elif cmd == "updateFeatureTable":
                self.update_feature_table(model)
            elif cmd == "updateMethodologyTable":
                self.update_methodology_table(model)
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
