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


def calc_macro(func):
    def wrapper(event):
        writer_doc = XSCRIPTCONTEXT.getDocument()
        writer_path = uno.fileUrlToSystemPath(writer_doc.getLocation())
        folder = os.path.dirname(writer_path)
        calc_path = os.path.join(folder, get_ods_fname(folder))

        struct = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        struct.Name = "Hidden"
        struct.Value = True

        desktop = XSCRIPTCONTEXT.getDesktop()
        calc_doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(calc_path),
            "_blank",
            0,
            tuple([struct])
        )

        func(writer_doc, calc_doc)
        
        calc_doc.dispose()

    return wrapper
