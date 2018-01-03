# Javelot

Javelot is a [LibreOffice extension](https://wiki.openoffice.org/wiki/Extensions_development)
helping [Nsigma](http://nsigma.fr/) project managers to write administrative documents.

## Installation

* Clone this repository
* `make`
* Add the extension to LO as explained [here](https://wiki.documentfoundation.org/Documentation/HowTo/install_extension)

## How it works

See the `docs` folder of this repository for details about the code. The extension is written in Python and
uses the [UNO API](https://forum.openoffice.org/fr/forum/viewtopic.php?f=37&t=53131).

The odt documents are in a common folder and share variables (dates, names, etc.)
through an ods document. The extension performs two kinds of operations.

### Document management

The extension automates some operations to manage the document itself. For instance:

* Linking the database
* Printing the document to the pdf format

### Document edition

The extension reads some data from the database (for instance, the budget table)
and inserts it in the opened document. 

Data is named by [named cells](https://help.libreoffice.org/Calc/Naming_Cells).
Thus, you do not need to dive into the code to change some data range. See the
`docs` folder for details about that.

The object (table, chart, text...) is inserted that way:

* If it already exists, the old one is replaced
* If it does not exist but there is a section to contain it, it is inserted in the section
* If it does not exist and there is no section for it, it is inserted at the cursor position

The names of the sections where data must be inserted are defined in the `Config` table
of the database.
