PYFILES=$(wildcard pythonpath/*.py)

all: javelot.oxt

javelot.oxt: addons.xcu main.py $(PYFILES) META-INF/manifest.xml
	zip -r javelot addons.xcu LICENSE main.py META-INF pythonpath
	mv javelot.zip javelot.oxt
