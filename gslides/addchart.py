# -*- coding: utf-8 -*-
from .utils import json_val_extract


"""Module to render the addChart Google API calll"""
"""
TODO:
- create pallette functionality
- validate args structure!!!!!

- UPDATE start_row_index!!!!! -- check if we are doing it write - what happens for nulls?
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
        self.header_count = 1
        self.executed = False
        self.ch_id = None

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
            if serie.y_columns is None:
                for column in self.data.df.columns.to_list():
                    if column != self.x_column:
                        series_mapping[column] = serie
            else:
                for column in serie.y_columns:
                    if column in self.data.df.columns.to_list():
                        series_mapping[column] = serie
        return series_mapping

    def render_basic_chart_json(self):
        json = {
            "chart": {
                "spec": {
                    "title": self.title,
                    "basicChart": {
                        "chartType": self.type,
                        "axis": [
                            {
                                "position": "BOTTOM_AXIS",
                                "title": self.x_axis_label,
                                "format": {"fontFamily": "Roboto"},
                                "viewWindowOptions": {},
                            },
                            {
                                "position": "LEFT_AXIS",
                                "title": self.y_axis_label,
                                "format": {"fontFamily": "Roboto"},
                                "viewWindowOptions": {},
                            },
                        ],
                        "domains": [{"domain": {"sourceRange": {"sources": []}}}],
                        "series": [],
                        "legendPosition": self.legend_position,
                        "headerCount": self.header_count,
                    },
                    "hiddenDimensionStrategy": "SKIP_HIDDEN_ROWS_AND_COLUMNS",
                    "titleTextFormat": {"fontFamily": "Roboto"},
                    "fontName": "Roboto",
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": self.data.sheet_id,
                            "rowIndex": self.data.start_row_index,
                        },
                        "offsetXPixels": 100,
                        "offsetYPixels": 9,
                        "widthPixels": 600,
                        "heightPixels": 371,
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
        for key, val in series_mapping.items():
            serie_col_num = (
                self.data.start_column_index + self.data.df.columns.to_list().index(key)
            )
            if self.type == "COMBO":
                series_json = val.render_basic_chart_json(
                    None,
                    self.data.sheet_id,
                    self.data.start_row_index - 1,
                    self.data.end_row_index,
                    serie_col_num - 1,
                    serie_col_num,
                    type=val.__class__.__name__.upper(),
                )
            else:
                series_json = val.render_basic_chart_json(
                    None,
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
        pass

    def execute(self, service):
        output = (
            service.spreadsheets()
            .batchUpdate(
                spreadsheetId=self.data.spreadsheet_id,
                body={"requests": [{"addChart": self.render_basic_chart_json()}]},
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
            "series": {
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
        data_label_enabled=False,
        data_label_placement=None,
        color=None,
    ):
        kwargs = locals().copy()
        del kwargs["self"], kwargs["__class__"]
        super().__init__(**kwargs)
