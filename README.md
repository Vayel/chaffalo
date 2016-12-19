# Javelot

Javelot is a [LibreOffice extension](https://wiki.openoffice.org/wiki/Extensions_development)
helping [Nsigma](http://nsigma.fr/) project managers to write administrative documents.

## Installation

* Clone this repository
* `make`
* Add the extension to LO as explained [here](https://wiki.documentfoundation.org/Documentation/HowTo/install_extension)

## How it works

See the `docs` folder of this repository for details about the code. The extension is written in Python and
uses the [UNO API](https://wiki.openoffice.org/wiki/Documentation/DevGuide/OpenOffice.org_Developers_Guide).

The odt documents are in a common folder and share variables (dates, names, etc.)
through an ods document. The extension links all the odt documents to the database. 
Then, it reads some data from the second and inserts them in a special way in the
text documents.

The data are named by [named cells](https://help.libreoffice.org/Calc/Naming_Cells).
Thus, you do not need to dive into the code to change some data range. See the
`docs` folder for détails about that.
