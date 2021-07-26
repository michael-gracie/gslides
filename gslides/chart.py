# -*- coding: utf-8 -*-
"""
Charts & series class
"""

import logging
import pprint
from typing import Any, Dict, List, Optional, Sequence, Tuple, Type, TypeVar, cast

from . import creds, package_font, package_palette
from .colors import Palette, translate_color
from .frame import Frame
from .utils import (
    hex_to_rgb,
    json_val_extract,
    validate_params_float,
    validate_params_int,
    validate_params_list,
    validate_series_columns,
)

logger = logging.getLogger(__name__)

TSeries = TypeVar("TSeries", bound="Series")


class Series:
    r"""Parent class for all series configurations.

    :param type: Type of series
    :type type: str
    :param \*\*kwargs: Dictionary of keyword arguments
    :type \*\*kwargs: dict
    """

    def __init__(self, type: str, **kwargs: Dict[str, Any]) -> None:
        """Constructor method"""
        self.type = type
        self._check_point_args(kwargs)
        self._check_data_label_args(kwargs)
        validate_params_list(kwargs)
        validate_params_int(kwargs)
        validate_params_float(kwargs)
        validate_series_columns(kwargs)
        self.params_dict = kwargs

    @classmethod
    def line(
        cls: Type[TSeries],
        series_columns: Optional[List[str]] = None,
        line_style: Optional[str] = None,
        line_width: Optional[int] = None,
        point_enabled: bool = False,
        point_shape: Optional[str] = None,
        point_size: Optional[int] = None,
        data_label_enabled: bool = False,
        data_label_placement: Optional[str] = None,
        color: Optional[str] = None,
    ) -> TSeries:
        """A line plot

        :param series_columns: The columns to plot. None or an empty list will
            plot all columns
        :type series_columns: list, optional
        :param line_style: The style of line to plot, see
            gslides.config.CHART_PARAMS['line_style'] for accepted parameters
        :type line_style: str, optional
        :param line_width: The width of line to plot
        :type line_width: int, optional
        :param point_enabled: Boolean for whether the plot should include points
        :type point_enabled: bool, optional
        :param point_shape: The shape of point to plot,
            see gslides.config.CHART_PARAMS['point_shape'] for accepted parameters
        :type point_shape: str, optional
        :param point_size: The size of point to plot
        :type point_size: int, optional
        :param data_label_enabled: Boolean for whether the plot should include data labels
        :type data_label_enabled: bool, optional
        :param data_label_placement: The placement of the data label to plot,
            see gslides.config.CHART_PARAMS['data_label_placement'] for accepted parameters
        :type data_label_placement: str, optional
        :param color: A color to override the existing palette. Parameters can either
            be a hex-code or a named colored. see gslides.config.color_mapping.keys()
            for accepted named colors
        :type color: str, optional
        :return: A :class:`Series` object
        :rtype: :class:`Series`
        """

        kwargs = locals().copy()
        del kwargs["cls"]
        return cls("Line", **kwargs)

    @classmethod
    def area(
        cls: Type[TSeries],
        series_columns: Optional[List[str]] = None,
        line_style: Optional[str] = None,
        line_width: Optional[int] = None,
        point_enabled: bool = False,
        point_shape: Optional[str] = None,
        point_size: Optional[int] = None,
        data_label_enabled: bool = False,
        data_label_placement: Optional[str] = None,
        color: Optional[str] = None,
    ) -> TSeries:
        """A area plot

        :param series_columns: The columns to plot. None or an empty list will
            plot all columns
        :type series_columns: list, optional
        :param line_style: The style of line to plot,
            see gslides.config.CHART_PARAMS['line_style'] for accepted parameters
        :type line_style: str, optional
        :param line_width: The width of line to plot
        :type line_width: int, optional
        :param point_enabled: Boolean for whether the plot should include points
        :type point_enabled: bool, optional
        :param point_shape: The shape of point to plot,
            see gslides.config.CHART_PARAMS['point_shape'] for accepted parameters
        :type point_shape: str, optional
        :param point_size: The size of point to plot
        :type point_size: int, optional
        :param data_label_enabled: Boolean for whether the plot should include data labels
        :type data_label_enabled: bool, optional
        :param data_label_placement: The placement of the data label to plot,
            see gslides.config.CHART_PARAMS['data_label_placement'] for accepted parameters
        :type data_label_placement: str, optional
        :param color: A color to override the existing palette. Parameters can either
            be a hex-code or a named colored. see gslides.config.color_mapping.keys()
            for accepted named colors
        :type color: str, optional
        :return: A :class:`Series` object
        :rtype: :class:`Series`
        """
        kwargs = locals().copy()
        del kwargs["cls"]
        return cls("Area", **kwargs)

    @classmethod
    def scatter(
        cls: Type[TSeries],
        series_columns: Optional[List[str]] = None,
        point_shape: Optional[str] = None,
        point_size: Optional[int] = None,
        data_label_enabled: bool = False,
        data_label_placement: Optional[str] = None,
        color: Optional[str] = None,
    ) -> TSeries:
        """A scatter plot

        :param series_columns: The columns to plot. None or an empty list will
            plot all columns
        :type series_columns: list, optional
        :param point_shape: The shape of point to plot,
            see gslides.config.CHART_PARAMS['point_shape'] for accepted parameters
        :type point_shape: str, optional
        :param point_size: The size of point to plot
        :type point_size: int, optional
        :param data_label_enabled: Boolean for whether the plot should include data labels
        :type data_label_enabled: bool, optional
        :param data_label_placement: The placement of the data label to plot,
            see gslides.config.CHART_PARAMS['data_label_placement'] for accepted parameters
        :type data_label_placement: str, optional
        :param color: A color to override the existing palette. Parameters can either
            be a hex-code or a named colored. see gslides.config.color_mapping.keys()
            for accepted named colors
        :type color: str, optional
        :return: A :class:`Series` object
        :rtype: :class:`Series`
        """
        kwargs = locals().copy()
        del kwargs["cls"]
        return cls("Scatter", **kwargs)

    @classmethod
    def column(
        cls: Type[TSeries],
        series_columns: Optional[List[str]] = None,
        data_label_enabled: bool = False,
        data_label_placement: Optional[str] = None,
        color: Optional[str] = None,
    ) -> TSeries:
        """A column plot

        :param series_columns: The columns to plot. None or an empty list will
            plot all columns
        :type series_columns: list, optional
        :param data_label_enabled: Boolean for whether the plot should include
            data labels
        :type data_label_enabled: bool, optional
        :param data_label_placement: The placement of the data label to plot,
            see gslides.config.CHART_PARAMS['data_label_placement'] for
            accepted parameters
        :type data_label_placement: str, optional
        :param color: A color to override the existing palette. Parameters can either
            be a hex-code or a named colored. see gslides.config.color_mapping.keys()
            for accepted named colors
        :type color: str, optional
        :return: A :class:`Series` object
        :rtype: :class:`Series`
        """
        kwargs = locals().copy()
        del kwargs["cls"]
        return cls("Column", **kwargs)

    @classmethod
    def histogram(
        cls: Type[TSeries],
        series_columns: Optional[List[str]] = None,
        bucket_size: Optional[int] = None,
        outlier_percentage: Optional[float] = None,
        color: Optional[str] = None,
    ) -> TSeries:
        """A histogram plot

        :param series_columns: The columns to plot. None or an empty list will
            plot all columns
        :type series_columns: list, optional
        :param bucket_size: The size of the bucket
        :type bucket_size: int, optional
        :param outlier_percentage: The percentile at which oberservations should
            be excluded
        :type outlier_percentage: float, optional
        :param color: A color to override the existing palette. Parameters can either
            be a hex-code or a named colored. see gslides.config.color_mapping.keys()
            for accepted named colors
        :type color: str, optional
        :return: A :class:`Series` object
        :rtype: :class:`Series`
        """
        kwargs = locals().copy()
        del kwargs["cls"]
        return cls("Histogram", **kwargs)

    def __repr__(self) -> str:
        """Prints class information.

        :return: String with helpful class infromation
        :rtype: str

        """
        output = f"Series Type: {self.type}"
        for key, val in self.params_dict.items():
            if val is None or val is False:
                pass
            else:
                output += f"\n - {key}: {val}"
        return output

    def _check_point_args(self, kwargs: dict) -> None:
        """Checks the args `point_shape`, `point_enabled` and `point_size` to ensure
        that `point_enabled` is set to `True` when utilizing the `point_size` and
        `point_shape` parameter.

        :param kwargs: Dictionary of keyword arguments
        :type kwargs: dict
        :raises ValueError:
        """
        if "point_enabled" in kwargs.keys():
            if kwargs["point_shape"] and kwargs["point_enabled"] is False:
                raise ValueError(
                    "point_enabled must be True if point_shape is specified"
                )
            elif kwargs["point_size"] and kwargs["point_enabled"] is False:
                raise ValueError(
                    "point_enabled must be True if point_size is specified"
                )
            else:
                return
        else:
            return

    def _check_data_label_args(self, kwargs: dict) -> None:
        """Checks the args `data_label_placement`, `data_label_enabled` to ensure
        that `point_enabled` is set to `True` when utilizing the `data_label_placement`
        parameter.

        :param kwargs: Dictionary of keyword arguments
        :type kwargs: dict
        :raises ValueError:
        """
        if "data_label_enabled" in kwargs.keys():
            if kwargs["data_label_placement"] and kwargs["data_label_enabled"] is False:
                raise ValueError(
                    "data_label_enabled must be True if point_shape is specified"
                )
            else:
                return
        else:
            return

    def render_basic_chart_json(
        self,
        palette: Optional[Palette],
        sheet_id: str,
        start_row_index: int,
        end_row_index: int,
        start_column_index: int,
        end_column_index: int,
        type: Optional[str] = None,
    ) -> dict:
        """Renders the json for the creation of a basic chart. See here
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/charts#BasicChartSpec
        for information about basic charts.

        :param palette: :class:`Palette` object to control colors
        :type palette: :class:`Palette`
        :param sheet_id: ID for the Google sheet.
        :type sheet_id: str
        :param start_row_index: The starting index of the row
        :type start_row_index: int
        :param end_row_index: The ending index of the row
        :type end_row_index: int
        :param start_column_index: The starting index of the column
        :type start_column_index: int
        :param end_column_index: The ending index of the column
        :type end_column_index: int
        :param type: Type of series
        :type type: str, optional
        :return: json for the API call
        :rtype: dict

        """

        json = {
            "target_axis": "LEFT_AXIS",
            "series": {
                "sourceRange": {
                    "sources": [
                        {
                            "sheetId": sheet_id,
                            "startRowIndex": start_row_index,
                            "endRowIndex": end_row_index,
                            "startColumnIndex": start_column_index,
                            "endColumnIndex": end_column_index,
                        }
                    ]
                }
            },
        }
        if "line_width" in self.params_dict.keys():
            json["lineStyle"] = {
                "type": self.params_dict["line_style"],
                "width": self.params_dict["line_width"],
            }
        if (
            "point_enabled" in self.params_dict.keys()
            and self.params_dict["point_enabled"]
        ):
            json["pointStyle"] = {
                "shape": self.params_dict["point_shape"],
                "size": self.params_dict["point_size"] or 5,
            }
        if "data_label_enabled" in self.params_dict.keys():
            if self.params_dict["data_label_enabled"]:
                json["dataLabel"] = {
                    "placement": self.params_dict["data_label_placement"],
                    "type": "DATA",
                    "textFormat": {
                        "fontFamily": package_font.font,
                        "fontSize": 12,
                    },
                }
        if type:
            json["type"] = type
        if self.params_dict["color"]:
            col = cast(str, self.params_dict["color"])
            r, g, b = hex_to_rgb(translate_color(col))
            json["color"] = {"red": r, "green": g, "blue": b}
            json["colorStyle"] = {"rgbColor": {"red": r, "green": g, "blue": b}}
        elif palette:
            r, g, b = next(palette)
            json["color"] = {"red": r, "green": g, "blue": b}
            json["colorStyle"] = {"rgbColor": {"red": r, "green": g, "blue": b}}
        return json

    def render_histogram_chart_json(
        self,
        palette: Optional[Palette],
        sheet_id: str,
        start_row_index: int,
        end_row_index: int,
        start_column_index: int,
        end_column_index: int,
    ) -> dict:
        """Renders the json for the creation of a basic chart. See here
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/charts#histogramchartspec
        for information about basic charts.

        :param palette: :class:`Palette` object to control colors
        :type palette: :class:`Palette`
        :param sheet_id: ID for the Google sheet.
        :type sheet_id: str
        :param start_row_index: The starting index of the row
        :type start_row_index: int
        :param end_row_index: The ending index of the row
        :type end_row_index: int
        :param start_column_index: The starting index of the column
        :type start_column_index: int
        :param end_column_index: The ending index of the column
        :type end_column_index: int
        :return: json for the API call
        :rtype: dict

        """
        json: Dict[str, Any] = {
            "data": {
                "sourceRange": {
                    "sources": [
                        {
                            "sheetId": sheet_id,
                            "startRowIndex": start_row_index,
                            "endRowIndex": end_row_index,
                            "startColumnIndex": start_column_index,
                            "endColumnIndex": end_column_index,
                        }
                    ]
                }
            }
        }
        if self.params_dict["color"]:
            col = cast(str, self.params_dict["color"])
            r, g, b = hex_to_rgb(translate_color(col))
            json["barColor"] = {"red": r, "green": g, "blue": b}
            json["barColorStyle"] = {"rgbColor": {"red": r, "green": g, "blue": b}}
        elif palette:
            r, g, b = next(palette)
            json["barColor"] = {"red": r, "green": g, "blue": b}
            json["barColorStyle"] = {"rgbColor": {"red": r, "green": g, "blue": b}}
        return json


class Chart:
    """An object that configures the creation of a chart in Google sheets.

    :param data: The data in Google sheets that will be plotted, a frame object
    :type data: :class:`gslides.Frame`
    :param x_axis_column: The name column that corresponds to the x-values.
        No parameter needed for a histogram.
    :type x_axis_column: str
    :param series: The :class:`gslides.addchart.series` objects to plot
    :type series: list[:class:`gslides.addchart.series`]
    :param stacking: The type of stacking to plot,
        see gslides.config.CHART_PARAMS['stacking'] for accepted parameters
    :type stacking: str, optional
    :param title: The title for the plot
    :type title: str, optional
    :param x_axis_label: The x_axis_label for the plot
    :type x_axis_label: str, optional
    :param y_axis_label: The y_axis_label for the plot
    :type y_axis_label: str, optional
    :param x_min: The minimum value for the x axis
    :type x_min: float, optional
    :param x_max: The maximum value for the x axis
    :type x_max: float, optional
    :param y_min: The minimum value for the y axis
    :type y_min: float, optional
    :param y_max: The maximum value for the y axis
    :type y_max: float, optional
    :param x_axis_format: The format of the x axis labels. Either the values
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#NumberFormatType
        or patterns here
        https://developers.google.com/sheets/api/guides/formats are accepted
    :type x_axis_format: str, optional
    :param y_axis_format: The format of the y axis labels. Either the values
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#NumberFormatType
        or patterns here
        https://developers.google.com/sheets/api/guides/formats are accepted
    :type y_axis_format: str, optional
    :param palette: The palette to use to plot,
        see gslides.colors.base_palettes for accepted parameters
    :type palette: str, optional
    :param legend_position: The position of the legend to plot,
        see gslides.config.CHART_PARAMS['stacking'] for accepted parameters
    :type legend_position: str, optional
    :param size: A tuple for the width and length of the plot in pixels.
        The Google suggested size is 600 by 371 pixels.
    :type size: tuple, optional
    """

    def __init__(
        self,
        data: Frame,
        x_axis_column: str,
        series: Sequence[Series],
        stacking: Optional[bool] = None,
        title: Optional[str] = None,
        x_axis_label: Optional[str] = None,
        y_axis_label: Optional[str] = None,
        x_min: Optional[float] = None,
        x_max: Optional[float] = None,
        y_min: Optional[float] = None,
        y_max: Optional[float] = None,
        x_axis_format: Optional[str] = None,
        y_axis_format: Optional[str] = None,
        palette: Optional[str] = None,
        legend_position: Optional[str] = None,
    ) -> None:
        """Constructor method"""
        self.type = self._determine_chart_type(series)
        self._check_stacking(series, stacking)
        self.data = data
        self.series = series
        self.stacking = stacking
        self.x_axis_column = x_axis_column
        self.title = title
        self.x_axis_label = x_axis_label
        self.y_axis_label = y_axis_label
        self.x_min = x_min
        self.x_max = x_max
        self.y_min = y_min
        self.y_max = y_max
        self.x_axis_format = x_axis_format
        self.y_axis_format = y_axis_format
        self.palette = palette
        self.legend_position = legend_position
        self.header_count = 1
        self.executed = False
        self.ch_id: Optional[str] = None
        self.bucket_size: Optional[int] = None
        self.outlier_percentage: Optional[float] = None
        validate_params_list(self.__dict__)

    def __repr__(self) -> str:
        """Prints class information.

        :return: String with helpful class infromation
        :rtype: str

        """
        output = f"Chart\n" f" - title = {self.title}"
        return output

    def _determine_chart_type(self, series: Sequence[Series]) -> str:
        """Determines the type of chart based on the class of series passed

        :param series: The :class:`gslides.addchart.series` objects to plot
        :type series: list[:class:`gslides.addchart.series`]
        :raises ValueError: Only Line, Area and Column series can be used in combination
        :return: The type of plot to create
        :rtype: str
        """
        chart_types = set([serie.type for serie in series])
        if len(chart_types) == 1:
            return chart_types.pop().upper()
        else:
            allowable_types = {"Line", "Area", "Column"}
            if len(chart_types.intersection(allowable_types)) == len(chart_types):
                return "COMBO"
            else:
                raise ValueError(
                    "Only Line, Area and Column series can be used in combination"
                )

    def _check_stacking(
        self, series: Sequence[Series], stacking: Optional[bool]
    ) -> None:
        """Checks the `stacking` argument to ensure that a stackable series is
        included. Stackable series include `Area` and `Column`.

        :param series: The :class:`gslides.addchart.series` objects to plot
        :type series: list[:class:`gslides.addchart.series`]
        :param series: The series objects to plot
        :type series: Sequence[Series]
        :raises ValueError: Stacking can only be enabled for Area and Column charts
        """
        chart_types = set([serie.type for serie in series])
        stacking_types = {"Area", "Column"}
        if stacking:
            if chart_types.intersection(stacking_types):
                return
            else:
                raise ValueError(
                    "Stacking can only be enabled for Area and Column charts"
                )
        else:
            return

    def _resolve_series(self) -> dict:
        """Resolves the series to determine which column get which configuration.
        Order in the list of series matters where if a column is set by 2 series
        classes the last series class will determine the configuration for the
        column

        :return: The column to plot and their corresponding configuration
        :rtype: dict
        """
        series_mapping = dict()
        for serie in self.series:
            if not serie.params_dict["series_columns"]:
                for column in self.data.df.columns.to_list():
                    if column != self.x_axis_column:
                        series_mapping[column] = serie
            else:
                for column in serie.params_dict["series_columns"]:
                    if column in self.data.df.columns.to_list():
                        series_mapping[column] = serie
            if (
                "outlier_percentage" in serie.params_dict.keys()
                and serie.params_dict["outlier_percentage"]
            ):
                self.outlier_percentage = cast(
                    Optional[float], serie.params_dict["outlier_percentage"]
                )
            if (
                "bucket_size" in serie.params_dict.keys()
                and serie.params_dict["bucket_size"]
            ):
                self.bucket_size = cast(Optional[int], serie.params_dict["bucket_size"])
        return series_mapping

    def render_basic_chart_json(self, size: Tuple[int, int]) -> dict:
        """Renders the json for the creation of a basic chart. See here
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/charts#histogramchartspec
        for information about basic charts.

        :param size: Tuple of width and height in PX
        :type size: tuple
        :return: json for the API call
        :rtype: dict
        """
        json: Dict[str, Any] = {
            "chart": {
                "spec": {
                    "title": self.title,
                    "titleTextPosition": {"horizontalAlignment": "CENTER"},
                    "titleTextFormat": {
                        "fontFamily": package_font.font,
                        "fontSize": 16,
                        "bold": True,
                        "foregroundColor": {"red": 0, "green": 0, "blue": 0},
                        "foregroundColorStyle": {
                            "rgbColor": {"red": 0, "green": 0, "blue": 0}
                        },
                    },
                    "basicChart": {
                        "chartType": self.type,
                        "axis": [
                            {
                                "position": "BOTTOM_AXIS",
                                "title": self.x_axis_label,
                                "format": {
                                    "fontFamily": package_font.font,
                                    "fontSize": 14,
                                    "bold": True,
                                    "foregroundColor": {
                                        "red": 0,
                                        "green": 0,
                                        "blue": 0,
                                    },
                                    "foregroundColorStyle": {
                                        "rgbColor": {"red": 0, "green": 0, "blue": 0}
                                    },
                                },
                                "viewWindowOptions": {},
                            },
                            {
                                "position": "LEFT_AXIS",
                                "title": self.y_axis_label,
                                "format": {
                                    "fontFamily": package_font.font,
                                    "fontSize": 14,
                                    "bold": True,
                                    "foregroundColor": {
                                        "red": 0,
                                        "green": 0,
                                        "blue": 0,
                                    },
                                    "foregroundColorStyle": {
                                        "rgbColor": {"red": 0, "green": 0, "blue": 0}
                                    },
                                },
                                "viewWindowOptions": {},
                            },
                        ],
                        "domains": [{"domain": {"sourceRange": {"sources": []}}}],
                        "series": [],
                        "legendPosition": self.legend_position,
                        "headerCount": self.header_count,
                    },
                    "hiddenDimensionStrategy": "SKIP_HIDDEN_ROWS_AND_COLUMNS",
                    "fontName": package_font.font,
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": self.data.sheet_id,
                            "rowIndex": self.data.start_row_index,
                            "columnIndex": self.data.start_column_index,
                        },
                        "offsetXPixels": 0,
                        "offsetYPixels": 0,
                        "widthPixels": size[0],
                        "heightPixels": size[1],
                    }
                },
            }
        }
        domain_col_num = (
            self.data.start_column_index
            + self.data.df.columns.to_list().index(self.x_axis_column)
        )
        domain_json = {
            "sheetId": self.data.sheet_id,
            "startRowIndex": self.data.start_row_index - 1,
            "endRowIndex": self.data.end_row_index,
            "startColumnIndex": domain_col_num - 1,
            "endColumnIndex": domain_col_num,
        }
        json["chart"]["spec"]["basicChart"]["domains"][0]["domain"]["sourceRange"][
            "sources"
        ].append(domain_json)
        series_mapping = self._resolve_series()
        if self.palette:
            p: Optional[Palette] = Palette(self.palette)
        elif package_palette.palette:
            p = Palette(package_palette.palette)
        else:
            p = None
        for key, val in series_mapping.items():
            serie_col_num = (
                self.data.start_column_index + self.data.df.columns.to_list().index(key)
            )
            if self.type == "COMBO":
                series_json = val.render_basic_chart_json(
                    p,
                    self.data.sheet_id,
                    self.data.start_row_index - 1,
                    self.data.end_row_index,
                    serie_col_num - 1,
                    serie_col_num,
                    type=val.type.upper(),
                )
            else:
                series_json = val.render_basic_chart_json(
                    p,
                    self.data.sheet_id,
                    self.data.start_row_index - 1,
                    self.data.end_row_index,
                    serie_col_num - 1,
                    serie_col_num,
                )
            json["chart"]["spec"]["basicChart"]["series"].append(series_json)
        if self.stacking:
            json["chart"]["spec"]["basicChart"]["stackedType"] = self.stacking
        if self.x_min is not None:
            json["chart"]["spec"]["basicChart"]["axis"][0]["viewWindowOptions"][
                "viewWindowMin"
            ] = self.x_min
        if self.x_max is not None:
            json["chart"]["spec"]["basicChart"]["axis"][0]["viewWindowOptions"][
                "viewWindowMax"
            ] = self.x_max
        if self.y_min is not None:
            json["chart"]["spec"]["basicChart"]["axis"][1]["viewWindowOptions"][
                "viewWindowMin"
            ] = self.y_min
        if self.y_max is not None:
            json["chart"]["spec"]["basicChart"]["axis"][1]["viewWindowOptions"][
                "viewWindowMax"
            ] = self.y_max
        return json

    def render_histogram_chart_json(self, size: Tuple[int, int]) -> dict:
        """Renders the json for the creation of a basic chart. See here
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/charts#histogramchartspec
        for information about basic charts.

        :param size: Tuple of width and height in PX
        :type size: tuple
        :return: json for the API call
        :rtype: dict
        """
        series_mapping = self._resolve_series()
        json: Dict[str, Any] = {
            "chart": {
                "spec": {
                    "title": self.title,
                    "titleTextPosition": {"horizontalAlignment": "CENTER"},
                    "titleTextFormat": {
                        "fontFamily": package_font.font,
                        "fontSize": 16,
                        "bold": True,
                        "foregroundColor": {"red": 0, "green": 0, "blue": 0},
                        "foregroundColorStyle": {
                            "rgbColor": {"red": 0, "green": 0, "blue": 0}
                        },
                    },
                    "histogramChart": {
                        "series": [],
                        "legendPosition": self.legend_position,
                        "bucketSize": self.bucket_size,
                        "outlierPercentile": self.outlier_percentage,
                    },
                    "hiddenDimensionStrategy": "SKIP_HIDDEN_ROWS_AND_COLUMNS",
                    "fontName": package_font.font,
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": self.data.sheet_id,
                            "rowIndex": self.data.start_row_index - 1,
                            "columnIndex": self.data.start_column_index - 1,
                        },
                        "offsetXPixels": 0,
                        "offsetYPixels": 0,
                        "widthPixels": size[0],
                        "heightPixels": size[1],
                    }
                },
            }
        }
        if self.palette:
            p: Optional[Palette] = Palette(self.palette)
        else:
            p = None
        for key, val in series_mapping.items():
            serie_col_num = (
                self.data.start_column_index + self.data.df.columns.to_list().index(key)
            )
            series_json = val.render_histogram_chart_json(
                p,
                self.data.sheet_id,
                self.data.start_row_index - 1,
                self.data.end_row_index,
                serie_col_num - 1,
                serie_col_num,
            )
            json["chart"]["spec"]["histogramChart"]["series"].append(series_json)
        return json

    def create(self, size: Tuple[int, int] = (600, 371)) -> dict:
        """Creates the chart in Googe sheets

        :param size: Tuple of width and height in PX
        :type size: tuple
        :return: The json returned by the call
        :rtype: dict

        """
        size = (int(size[0]), int(size[1]))
        format_columns = {}
        if self.x_axis_format:
            format_columns[self.x_axis_column] = self.x_axis_format
        if self.y_axis_format:
            series_mapping = self._resolve_series()
            for key in series_mapping.keys():
                format_columns[key] = self.y_axis_format
        if format_columns:
            self.data.format_frame(format_columns)
        service: Any = creds.sheet_service
        if self.type == "HISTOGRAM":
            json = self.render_histogram_chart_json(size)
        else:
            json = self.render_basic_chart_json(size)
        body = {"requests": [{"addChart": json}]}
        logger.info("Executing chart creation")
        logger.info(f"Request: {pprint.pformat(body)}")
        output: dict = (
            service.spreadsheets()
            .batchUpdate(
                spreadsheetId=self.data.spreadsheet_id,
                body=body,
            )
            .execute()
        )
        logger.info("Chart created successfully")
        self.ch_id = json_val_extract(output, "chartId")[0]
        self.executed = True
        return output

    @property
    def chart_id(self) -> Optional[str]:
        """Returns the chart_id of the created chart.

        :raises RuntimeError: Must run the execute method before passing the chart id
        :return: The chart_id of the created chart.
        :rtype: str
        """

        if self.executed:
            return self.ch_id
        else:
            raise RuntimeError(
                "Must run the execute method before passing the chart id"
            )
