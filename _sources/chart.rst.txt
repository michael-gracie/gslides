Charts & Series
=========================================

Below, the documentation will cover some key considerations before creating a ``Series`` and a ``Chart``.

Ensuring your data is pivoted
------------------------------------------

It is necessary to structure your data properly into Google sheets such that it can be plotted. Google sheets can only plot data that has been pivoted.

A table such as this:

.. csv-table::
   :header: "Shape", "Color", "Size"

   "Cube","Red","1"
   "Cube","Green","2"
   "Ball","Red","3"
   "Ball","Green","4"

Needs to be manipulated like so:

.. csv-table::
  :header: "Shape", "Red", "Green"

  "Cube","1","2"
  "Ball","3","4"

This can be done through `df.pivot() <https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.pivot.html>`_.

Utilizing the ``series_columns`` parameter
------------------------------------------

Each series in a chart can be formatted in a unique way. For this reason, users are able to pass multiple ``Series`` objects into a ``Chart`` object such that they have full control over the display for a given chart. In fact, it is possible to pass different ``Series`` types into the same chart.

- If ``series_columns`` is ``None`` or ``[]``, the chart will plot all columns in the ``Frame`` object except the ``x_column``.
- If different ``Series`` types are passed into the same chart, the chart is considered a ``COMBO`` chart. Only combinations of ``Area``, ``Columns`` and ``Line`` series are allowed.
- If the same ``series_column`` is set in multiple ``Series`` objects, the latter most configuration of that column will be used.

This provides great flexibility. A common pattern is to set all columns to a base configuration by passing ``None``. Then, creating another ``Series`` object to override that configuration for a single column to set an alternative color or style, highlighting that particular series.

Series parameters
------------------------------------------

Various parameters can be set based on the type of series.

.. csv-table::
   :header: "Parameter", "Line", "Area", "Column", "Scatter", "Histogram"

   "series_columns","Y","Y","Y","Y","Y"
   "line_style","Y","Y",,,
   "line_width","Y","Y",,,
   "point_enabled","Y","Y",,,
   "point_shape","Y","Y",,"Y",
   "point_size","Y","Y",,"Y",
   "data_label_enabled","Y","Y","Y","Y",
   "data_label_placement","Y","Y","Y","Y",
   "bucket_size",,,,,"Y"
   "outlier_percentage",,,,,"Y"
   "color","Y","Y","Y","Y","Y"

To understand the values for each parameter accepted run the following:

.. code-block:: python

  from gslides import CHART_PARAMS
  for key, val in CHART_PARAMS.items():
      print(f"{key}: {val['params']}")

Stacking
------------------------------------------

Stacking can only be enabled at the ``Chart`` level when at least one of the ``Series`` is either of type ``Area`` or ``Column``.

See the ``CHART_PARAMS`` dictionary for valid stacking types.

Axis formatting
------------------------------------------

To change the formatting on a given axis you may set the ``x_axis_format`` or ``y_axis_format`` parameter when initializing a ``Chart`` object. These columns either accept a basic formatting configuration that can be found `here <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#NumberFormatType>`_, or a customized configuration that is explained `here <https://developers.google.com/sheets/api/guides/formats>`_.

Histogram
------------------------------------------

Histograms are unique in that Google will perform the bucketing of values before plotting. In this way ``x_axis_column`` for Histogram charts can simply be set to a dummy string.
