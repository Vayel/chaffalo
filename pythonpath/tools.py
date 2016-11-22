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


def get_document_name(document):
    return os.path.splitext(os.path.basename(get_document_path(document)))[0]


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
