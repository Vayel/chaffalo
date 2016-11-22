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
