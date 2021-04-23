# -*- coding: utf-8 -*-
from typing import Any, Dict, List, Optional, Sequence, Tuple, cast

from googleapiclient.discovery import Resource

from .colors import Palette, translate_color
from .sheetsframe import SheetsFrame
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


class Series:
    def __init__(self, **kwargs: dict) -> None:
        self._check_point_args(kwargs)
        self._check_data_label_args(kwargs)
        validate_params_list(kwargs)
        validate_params_int(kwargs)
        validate_params_float(kwargs)
        self.params_dict = kwargs

    def __repr__(self) -> str:
        output = f"Series Type: {self.__class__.__name__}"
        for key, val in self.params_dict.items():
            if val is None or val is False:
                pass
            else:
                output += f"\n - {key}: {val}"
        return output

    def _check_point_args(self, kwargs: dict) -> None:
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
        sheetId: str,
        start_row_index: int,
        end_row_index: int,
        start_column_index: int,
        end_column_index: int,
        type: Optional[str] = None,
    ) -> dict:

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
                "size": self.params_dict["point_size"],
            }
        if "data_label_enabled" in self.params_dict.keys():
            if self.params_dict["data_label_enabled"]:
                json["dataLabel"] = {
                    "placement": self.params_dict["data_label_placement"],
                    "type": "DATA",
                    "textFormat": {
                        "fontFamily": "Roboto",  # TODO figure out what to do
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
        sheetId: str,
        start_row_index: int,
        end_row_index: int,
        start_column_index: int,
        end_column_index: int,
    ) -> dict:
        json: Dict[str, Any] = {
            "data": {
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


class Line(Series):
    def __init__(
        self,
        y_columns: Optional[List[str]] = None,
        line_style: Optional[str] = None,
        line_width: Optional[int] = None,
        point_enabled: bool = False,
        point_shape: Optional[str] = None,
        point_size: Optional[int] = None,
        data_label_enabled: bool = False,
        data_label_placement: Optional[str] = None,
        color: Optional[str] = None,
    ) -> None:
        kwargs = locals().copy()
        del kwargs["self"], kwargs["__class__"]
        super().__init__(**kwargs)


class Column(Series):
    def __init__(
        self,
        y_columns: Optional[List[str]] = None,
        data_label_enabled: bool = False,
        data_label_placement: Optional[str] = None,
        color: Optional[str] = None,
    ) -> None:
        kwargs = locals().copy()
        del kwargs["self"], kwargs["__class__"]
        super().__init__(**kwargs)


class Area(Series):
    def __init__(
        self,
        y_columns: Optional[List[str]] = None,
        line_style: Optional[str] = None,
        line_width: Optional[int] = None,
        point_enabled: bool = False,
        point_shape: Optional[str] = None,
        point_size: Optional[int] = None,
        data_label_enabled: bool = False,
        data_label_placement: Optional[str] = None,
        color: Optional[str] = None,
    ) -> None:
        kwargs = locals().copy()
        del kwargs["self"], kwargs["__class__"]
        super().__init__(**kwargs)


class Scatter(Series):
    def __init__(
        self,
        y_columns: Optional[List[str]] = None,
        point_shape: Optional[str] = None,
        point_size: Optional[int] = None,
        data_label_enabled: bool = False,
        data_label_placement: Optional[str] = None,
        color: Optional[str] = None,
    ) -> None:
        kwargs = locals().copy()
        del kwargs["self"], kwargs["__class__"]
        kwargs["point_enabled"] = True
        super().__init__(**kwargs)


class Histogram(Series):
    def __init__(
        self,
        y_columns: Optional[List[str]] = None,
        bucket_size: Optional[int] = None,
        outlier_percentage: Optional[float] = None,
        color: Optional[str] = None,
    ) -> None:
        kwargs = locals().copy()
        del kwargs["self"], kwargs["__class__"]
        super().__init__(**kwargs)


class Chart:
    def __init__(
        self,
        data: SheetsFrame,
        x_column: str,
        series: Sequence[Series],
        stacking: Optional[bool] = None,
        title: Optional[str] = None,
        x_axis_label: Optional[str] = None,
        y_axis_label: Optional[str] = None,
        x_min: Optional[float] = None,
        x_max: Optional[float] = None,
        y_min: Optional[float] = None,
        y_max: Optional[float] = None,
        palette: Optional[str] = None,
        legend_position: Optional[str] = None,
        size: Tuple[int, int] = (600, 371),
    ) -> None:
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
        self.ch_id: Optional[str] = None
        self.bucket_size: Optional[int] = None
        self.outlier_percentage: Optional[float] = None
        validate_params_list(self.__dict__)

    def _determine_chart_type(self, series: Sequence[Series]) -> str:
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

    def _check_stacking(
        self, series: Sequence[Series], stacking: Optional[bool]
    ) -> None:
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

    def _resolve_series(self) -> dict:
        series_mapping = dict()
        for serie in self.series:
            if not serie.params_dict["y_columns"]:
                for column in self.data.df.columns.to_list():
                    if column != self.x_column:
                        series_mapping[column] = serie
            else:
                for column in serie.params_dict["y_columns"]:
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

    def render_basic_chart_json(self) -> dict:
        json: Dict[str, Any] = {
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
            p: Optional[Palette] = Palette(self.palette)
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

    def render_histogram_chart_json(self) -> dict:
        series_mapping = self._resolve_series()
        json: Dict[str, Any] = {
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

    def execute(self, service: Resource) -> dict:
        if self.type == "HISTOGRAM":
            json = self.render_histogram_chart_json()
        else:
            json = self.render_basic_chart_json()
        output: dict = (
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
    def chart_id(self) -> Optional[str]:
        if self.executed:
            return self.ch_id
        else:
            raise RuntimeError(
                "Must run the execute method before passing the chart id"
            )
