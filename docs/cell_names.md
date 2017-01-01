# Cell names

What cell names can be used to configure the extension?

## Config sheet

Cell name | Description | Cell content example
----------|-------------|----------------------
gantt_data_sheet | The sheet containing the data for the Gantt diagram | Phases
gantt_chart_name | The name of the Gantt chart in odt documents | Gantt
gantt_section_name | The name of the section for the Gantt diagram in odt documents | Gantt
gantt_time_step | The step of the time axis of the Gantt | 1
gantt_label_legend | The legend of the label axis | Phases

## Gantt data sheet

Cell name | Description | Cell content example
----------|-------------|-----------------------
gantt_time_unit | The unit displayed in the time axis legend | weeks

Range name | Description | Cells content
-----------|-------------|--------------
gantt_labels | The range of the labels | strings
gantt_beginnings | The range of the origin dates | positive-or-zero integers
gantt_lengths | The range of the durations | positive integers

The ranges `gantt_labels`, `gantt_beginnings` and `gantt_lengths` must have the
same length.
