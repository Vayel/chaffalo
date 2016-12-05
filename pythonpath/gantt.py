import uno

import config
import tools


CONFIG_SHEET_NAME = "Gantt"
DATA_SHEET_NAME_CELL = "data_sheet"
CHART_NAME_CELL = "chart_name"
SECTION_NAME_CELL = "section_name"
TIME_UNIT_CELL = "time_unit"
TIME_STEP_CELL = "time_step"
LABEL_LEGEND_CELL = "label_legend"

LABEL_RANGE = "labels"
VALUE_RANGE = "values"
LENGTH_RANGE = "lengths"


def get_anchor(document, config_sheet):
    # Look for a section
    section_name = config_sheet.getCellRangeByName(SECTION_NAME_CELL).getString()
    if document.getTextSections().hasByName(section_name):
        return document.getTextSections().getByName(section_name).getAnchor()

    # Return the current cursor position
    return document.getCurrentController().getViewCursor()


def get_data(sheet):
    values = (x[0] for x in sheet.getCellRangeByName(VALUE_RANGE).getDataArray())
    lengths = (x[0] for x in sheet.getCellRangeByName(LENGTH_RANGE).getDataArray())

    return tuple(zip(values, lengths))


def get_labels(sheet):
    return tuple(x[0] for x in sheet.getCellRangeByName(LABEL_RANGE).getDataArray())


def get_time_step(sheet):
    return float(sheet.getCellRangeByName(TIME_STEP_CELL).getValue())


def get_label_legend(sheet):
    return sheet.getCellRangeByName(LABEL_LEGEND_CELL).getString()


def get_time_axis_unit(sheet):
    return sheet.getCellRangeByName(TIME_UNIT_CELL).getString()


def get_chart_name(sheet):
    return sheet.getCellRangeByName(CHART_NAME_CELL).getString()


def get_chart(document, sheet):
    name = get_chart_name(sheet)
    if document.getEmbeddedObjects().hasByName(name):
        return document.getEmbeddedObjects().getByName(name).getEmbeddedObject()

    # Create a new chart
    # https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Embedded_Objects
    embedded_obj = document.createInstance("com.sun.star.text.TextEmbeddedObject")
    embedded_obj.setName(name)
    embedded_obj.CLSID = "12DCAE26-281F-416F-a234-c3086127382e"

    anchor = get_anchor(document, sheet)
    anchor.getText().insertTextContent(anchor, embedded_obj, False)
    embedded_obj.AnchorType = uno.Enum("com.sun.star.text.TextContentAnchorType", "AS_CHARACTER")
    
    chart = embedded_obj.getEmbeddedObject()
    chart.setDiagram(chart.createInstance("com.sun.star.chart.BarDiagram"))

    diagram = chart.getDiagram()
    diagram.getXAxis().AxisTitle.String = get_label_legend(sheet)
    diagram.getXAxis().setPropertyValue("ReverseDirection", True)
    diagram.getYAxis().setPropertyValue("CrossoverPosition", uno.Enum("com.sun.star.chart.ChartAxisPosition", "END"))
    diagram.Wall.FillColor = tools.RGB(255, 255, 255)

    diagram.setPropertyValue("Stacked", True)
    diagram.setPropertyValue("Percent", False)
    diagram.setPropertyValue("Vertical", True)

    return chart


def update(document, calc):
    config_sheet = calc.getSheets().getByName(CONFIG_SHEET_NAME)
    data_sheet_name = config_sheet.getCellRangeByName(DATA_SHEET_NAME_CELL).getString()
    data_sheet = calc.getSheets().getByName(data_sheet_name)

    data = get_data(data_sheet)
    labels = get_labels(data_sheet)

    time_axis_max = data[-1][1]
    time_axis_title = "Durée ({})".format(get_time_axis_unit(config_sheet))
    
    chart = get_chart(document, config_sheet)

    # Update the chart
    chart.getDiagram().getYAxis().setPropertyValue("Max", time_axis_max)
    chart.getDiagram().getYAxis().setPropertyValue("StepMain", get_time_step(config_sheet))
    chart.getData().setData(data)
    chart.getData().setRowDescriptions(labels)
    chart.getDiagram().getYAxis().AxisTitle.String = time_axis_title

    # https://forum.openoffice.org/en/forum/viewtopic.php?t=36001
    chart.getFirstDiagram().getCoordinateSystems()[0].getChartTypes()[0].getDataSeries()[0].Transparency = 100
    chart.getFirstDiagram().getCoordinateSystems()[0].getChartTypes()[0].getDataSeries()[1].Color = tools.RGB(0, 153, 0)
