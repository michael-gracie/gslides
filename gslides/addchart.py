# -*- coding: utf-8 -*-
from .colors import Palette, translate_color
from .utils import (
    hex_to_rgb,
    json_val_extract,
    validate_params_float,
    validate_params_int,
    validate_params_list,
)


"""Module to render the addChart Google API call"""
"""
TODO:
- formatting config file
- finish pytest
- copy slide
- mypy
- check for continuous series in date axis
"""


class Chart:
    def __init__(
        self,
        data,
        series,
        stacking=None,
        x_column=None,
        title=None,
        x_axis_label=None,
        y_axis_label=None,
        x_min=None,
        x_max=None,
        y_min=None,
        y_max=None,
        palette=None,
        legend_position=None,
        size=(600, 371),
    ):
        self.type = self._determine_chart_type(series)
        self._check_stacking(series, stacking)
        self.data = data
        self.series = series
        self.stacking = stacking
        self.x_column = x_column
        self.title = title
        self.x_axis_label = x_axis_label
        self.y_axis_label = y_axis_label
        self.x_min = x_min
        self.x_max = x_max
        self.y_min = y_min
        self.y_max = y_max
        self.palette = palette
        self.legend_position = legend_position
        self.size = (int(size[0]), int(size[1]))
        self.header_count = 1
        self.executed = False
        self.ch_id = None
        self.bucket_size = None
        self.outlier_percentage = None
        validate_params_list(self.__dict__)

    def _determine_chart_type(self, series):
        chart_types = set([serie.__class__.__name__ for serie in series])
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

    def _check_stacking(self, series, stacking):
        chart_types = set([serie.__class__.__name__ for serie in series])
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

    def _resolve_series(self):
        series_mapping = dict()
        for serie in self.series:
            if not serie.y_columns:
                for column in self.data.df.columns.to_list():
                    if column != self.x_column:
                        series_mapping[column] = serie
            else:
                for column in serie.y_columns:
                    if column in self.data.df.columns.to_list():
                        series_mapping[column] = serie
            if (
                "outlier_percentage" in serie.__dict__.keys()
                and serie.outlier_percentage
            ):
                self.outlier_percentage = serie.outlier_percentage
            if "bucket_size" in serie.__dict__.keys() and serie.bucket_size:
                self.bucket_size = serie.bucket_size
        return series_mapping

    def render_basic_chart_json(self):
        json = {
            "chart": {
                "spec": {
                    "title": self.title,
                    "titleTextPosition": {"horizontalAlignment": "CENTER"},
                    "titleTextFormat": {
                        "fontFamily": "Roboto",
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
                                    "fontFamily": "Roboto",
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
                                    "fontFamily": "Roboto",
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
                    "fontName": "Roboto",
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
                        "widthPixels": self.size[0],
                        "heightPixels": self.size[1],
                    }
                },
            }
        }
        domain_col_num = (
            self.data.start_column_index
            + self.data.df.columns.to_list().index(self.x_column)
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
            p = Palette(self.palette)
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
                    type=val.__class__.__name__.upper(),
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
            json["chart"]["spec"]["basicChart"]["stackedType"] = "STACKED"
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

    def render_histogram_chart_json(self):
        series_mapping = self._resolve_series()  # figure out
        json = {
            "chart": {
                "spec": {
                    "title": self.title,
                    "titleTextPosition": {"horizontalAlignment": "CENTER"},
                    "titleTextFormat": {
                        "fontFamily": "Roboto",
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
                    "fontName": "Roboto",
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
                        "widthPixels": self.size[0],
                        "heightPixels": self.size[1],
                    }
                },
            }
        }
        if self.palette:
            p = Palette(self.palette)
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

    def execute(self, service):
        if self.type == "HISTOGRAM":
            json = self.render_histogram_chart_json()
        else:
            json = self.render_basic_chart_json()
        output = (
            service.spreadsheets()
            .batchUpdate(
                spreadsheetId=self.data.spreadsheet_id,
                body={"requests": [{"addChart": json}]},
            )
            .execute()
        )
        self.ch_id = json_val_extract(output, "chartId")
        self.executed = True
        return output

    @property
    def chart_id(self):
        if self.executed:
            return self.ch_id
        else:
            raise RuntimeError(
                "Must run the execute method before passing the chart id"
            )


class Series:
    def __init__(self, **kwargs):
        self._check_point_args(kwargs)
        self._check_data_label_args(kwargs)
        self.__dict__.update(kwargs)
        validate_params_list(kwargs)
        validate_params_int(kwargs)
        validate_params_float(kwargs)

    def __repr__(self):
        output = f"Series Type: {self.__class__.__name__}"
        for key, val in self.__dict__.items():
            if val is None or val is False:
                pass
            else:
                output += f"\n - {key}: {val}"
        return output

    def _check_point_args(self, kwargs):
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

    def _check_data_label_args(self, kwargs):
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
        palette,
        sheetId,
        start_row_index,
        end_row_index,
        start_column_index,
        end_column_index,
        type=None,
    ):

        json = {
            "target_axis": "LEFT_AXIS",
            "series": {
                "sourceRange": {
                    "sources": [
                        {
                            "sheetId": sheetId,
                            "startRowIndex": start_row_index,
                            "endRowIndex": end_row_index,
                            "startColumnIndex": start_column_index,
                            "endColumnIndex": end_column_index,
                        }
                    ]
                }
            },
        }
        if "line_width" in self.__dict__.keys():
            json["lineStyle"] = {"type": self.line_style, "width": self.line_width}
        if "point_enabled" in self.__dict__.keys() and self.point_enabled:
            json["pointStyle"] = {"shape": self.point_shape, "size": self.point_size}
        if "data_label_enabled" in self.__dict__.keys():
            if self.data_label_enabled:
                json["dataLabel"] = {
                    "placement": self.data_label_placement,
                    "type": "DATA",
                    "textFormat": {
                        "fontFamily": "Roboto",  # TODO figure out what to do
                        "fontSize": 12,
                    },
                }
        if type:
            json["type"] = type
        if self.color:
            r, g, b = hex_to_rgb(translate_color(self.color))
            json["color"] = {"red": r, "green": g, "blue": b}
            json["colorStyle"] = {"rgbColor": {"red": r, "green": g, "blue": b}}
        elif palette:
            r, g, b = next(palette)
            json["color"] = {"red": r, "green": g, "blue": b}
            json["colorStyle"] = {"rgbColor": {"red": r, "green": g, "blue": b}}
        return json

    def render_histogram_chart_json(
        self,
        palette,
        sheetId,
        startRowIndex,
        endRowIndex,
        startColumnIndex,
        endColumnIndex,
    ):
        json = {
            "data": {
                "sourceRange": {
                    "sources": [
                        {
                            "sheetId": sheetId,
                            "startRowIndex": startRowIndex,
                            "endRowIndex": endRowIndex,
                            "startColumnIndex": startColumnIndex,
                            "endColumnIndex": endColumnIndex,
                        }
                    ]
                }
            }
        }
        if self.color:
            r, g, b = hex_to_rgb(translate_color(self.color))
            json["barColor"] = {"red": r, "green": g, "blue": b}
            json["barColorStyle"] = {"rgbColor": {"red": r, "green": g, "blue": b}}
        elif palette:
            r, g, b = next(palette)
            json["barColor"] = {"red": r, "green": g, "blue": b}
            json["barColorStyle"] = {"rgbColor": {"red": r, "green": g, "blue": b}}
        return json


class Line(Series):
    def __init__(
        self,
        y_columns=None,
        line_style=None,
        line_width=None,
        point_enabled=False,
        point_shape=None,
        point_size=None,
        data_label_enabled=False,
        data_label_placement=None,
        color=None,
    ):
        kwargs = locals().copy()
        del kwargs["self"], kwargs["__class__"]
        super().__init__(**kwargs)


class Column(Series):
    def __init__(
        self,
        y_columns=None,
        data_label_enabled=False,
        data_label_placement=None,
        color=None,
    ):
        kwargs = locals().copy()
        del kwargs["self"], kwargs["__class__"]
        super().__init__(**kwargs)


class Area(Series):
    def __init__(
        self,
        y_columns=None,
        line_style=None,
        line_width=None,
        point_enabled=False,
        point_shape=None,
        point_size=None,
        data_label_enabled=False,
        data_label_placement=None,
        color=None,
    ):
        kwargs = locals().copy()
        del kwargs["self"], kwargs["__class__"]
        super().__init__(**kwargs)


class Scatter(Series):
    def __init__(
        self,
        y_columns=None,
        point_shape=None,
        point_size=None,
        data_label_enabled=False,
        data_label_placement=None,
        color=None,
    ):
        kwargs = locals().copy()
        del kwargs["self"], kwargs["__class__"]
        kwargs["point_enabled"] = True
        super().__init__(**kwargs)


class Histogram(Series):
    def __init__(
        self,
        y_columns=None,
        bucket_size=None,
        outlier_percentage=None,
        color=None,
    ):
        kwargs = locals().copy()
        del kwargs["self"], kwargs["__class__"]
        super().__init__(**kwargs)
