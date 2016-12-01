import os
import uno

import config


def get_ods_fname(folder):
    return [fname for fname in os.listdir(folder) if fname.endswith(".ods")][0]


def get_odt_fnames(folder):
    return [fname for fname in os.listdir(folder) if fname.endswith(".odt")]


def get_document_path(document):
    return uno.fileUrlToSystemPath(document.URL)


def get_document_folder(document):
    return os.path.dirname(get_document_path(document))


def get_document_fname(document):
    return os.path.basename(get_document_path(document))


def get_document_name(document):
    return os.path.splitext(get_document_fname(document))[0]


def bold_cell(cell):
    cursor = cell.getStart()
    cursor.gotoEnd(True)
    cursor.setPropertyValue("CharWeight", uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD"))


def hide_border(cell):
    prop_names = ["LeftBorder", "TopBorder", "RightBorder"]
    border = uno.createUnoStruct("com.sun.star.table.BorderLine")

    for prop_name in prop_names:
        cell.setPropertyValue(prop_name, border)


def superscript_cell_note(cell):
    cursor = cell.getEnd()
    cursor.goLeft(1, True)
    cursor.setPropertyValue("CharEscapement", 50)
    cursor.setPropertyValue("CharHeight", 8)
    

def align_cell_text(cell, row_count, i, j):
    col_aligns = [
        uno.Enum("com.sun.star.style.ParagraphAdjust", "LEFT"),
        uno.Enum("com.sun.star.style.ParagraphAdjust", "CENTER"),
        uno.Enum("com.sun.star.style.ParagraphAdjust", "RIGHT"),
        uno.Enum("com.sun.star.style.ParagraphAdjust", "RIGHT"),
    ]
    
    if cell.getString() == "-":
        align = uno.Enum("com.sun.star.style.ParagraphAdjust", "CENTER")
    elif i >= row_count - 6 and j == 0:
        align = uno.Enum("com.sun.star.style.ParagraphAdjust", "RIGHT")
    else:
        align = col_aligns[j]

    cursor = cell.getStart()
    cursor.gotoEnd(True)
    cursor.setPropertyValue("ParaAdjust", align)


def get_phase_names(calc_doc):
    LAST_ROW_NUMBER_CELL = "K2"

    sheet = calc_doc.getSheets().getByName("Phases")
    last_row_number = int(sheet.getCellRangeByName(
        LAST_ROW_NUMBER_CELL
    ).getValue())
    names = sheet.getCellRangeByName("A2:A{}".format(
        last_row_number
    )).getDataArray()

    return [name[0] for name in names]
    

def set_table_row_count(table, count):
    delta = count - table.getRows().getCount()

    if delta < 0:
        delta = abs(delta)
        table.getRows().removeByIndex(table.getRows().getCount() - delta, delta)
    elif delta > 0:
        table.getRows().insertByIndex(table.getRows().getCount(), delta)


def get_cell_style(cell):
    cursor = cell.getStart()
    cursor.gotoEnd(True)

    return cursor.ParaStyleName


def set_cell_style(cell, style):
    cursor = cell.getStart()
    cursor.gotoEnd(True)
    cursor.ParaStyleName = style


def copy_cell_string(dst, src):
    dst.setString(src.getString())
    set_cell_style(dst, get_cell_style(src))


def replace_table_by_another(dst, src):
    assert(dst.getColumns().getCount() == src.getColumns().getCount())

    set_table_row_count(dst, src.getRows().getCount())

    for i in range(src.getRows().getCount()):
        for j in range(src.getColumns().getCount()):
            dst_cell = dst.getCellByPosition(j, i)
            src_cell = src.getCellByPosition(j, i)
            copy_cell_string(dst_cell, src_cell)


def update_phase_table(writer_doc, calc_doc, table_name, old_table_name, col_count, styles=None):
    def insert_table(name):
        table = writer_doc.createInstance("com.sun.star.text.TextTable")
        table.setName(name)
        table.initialize(1, col_count)

        bookmark = writer_doc.getBookmarks().getByName(name)
        anchor = bookmark.getAnchor()
        anchor.getText().insertTextContent(anchor, table, False)

        return table
    
    if styles is None:
        styles = ["Table Contents"] * col_count

    if not writer_doc.getTextTables().hasByName(table_name):
        table = insert_table(table_name)
    else:
        table = writer_doc.getTextTables().getByName(table_name)

    if not writer_doc.getTextTables().hasByName(old_table_name):
        old_table = insert_table(old_table_name)
    else:
        old_table = writer_doc.getTextTables().getByName(old_table_name)
    
    # Save old phases
    replace_table_by_another(old_table, table)

    # Add new phases
    names = get_phase_names(calc_doc)[:-1]
    set_table_row_count(table, len(names))

    for i in range(table.getRows().getCount()):
        cell = table.getCellByPosition(0, i)
        cell.setString(names[i])
        cell.setPropertyValue("IsProtected", True)
        set_cell_style(cell, styles[0])

        for j in range(1, col_count):
            cell = table.getCellByPosition(j, i)
            set_cell_style(cell, styles[j])

        for old_i in range(old_table.getRows().getCount()):
            if old_table.getCellByPosition(0, old_i).getString() != table.getCellByPosition(0, i).getString():
                continue

            old_cell = old_table.getCellByPosition(col_count - 1, old_i)
            cell = table.getCellByPosition(col_count - 1, i)
            copy_cell_string(cell, old_cell)
            old_table.getRows().removeByIndex(old_i, 1)
            break


def find_insert_position(document, elements, section_name, element_name):
    if elements.hasByName(element_name):
        el = elements.getByName(element_name)
        return el, el.getAnchor()

    if document.getTextSections().hasByName(section_name):
        return None, document.getTextSections().getByName(section_name).getAnchor()

    return None, TODO
