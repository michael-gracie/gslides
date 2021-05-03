# -*- coding: utf-8 -*-
"""
Creates the slides and charts in Google slides
"""

from typing import Any, Dict, List, Optional, Tuple, TypeVar, Union

import pandas as pd

from . import creds
from .addchart import Chart
from .colors import translate_color
from .sheetsframe import SheetsFrame
from .utils import (
    black_or_white,
    clean_dtypes,
    clean_nan,
    determine_col_proportion,
    hex_to_rgb,
    optimize_size,
    validate_params_float,
)


class CreatePresentation:
    """The class that creates the presentation.

    :param name: Name of the presentation
    :type name: str
    """

    def __init__(
        self,
        name: str = "Untitled",
    ) -> None:
        """Constructor method"""
        self.name = name
        self.executed = False
        self.pr_id: Optional[str] = None

    def execute(self) -> None:
        """Executes the create presentation slides API call."""
        service: Any = creds.slide_service
        output = service.presentations().create(body={"title": self.name}).execute()
        self.pr_id = output["presentationId"]
        service.presentations().batchUpdate(
            presentationId=output["presentationId"],
            body={"requests": [{"deleteObject": {"objectId": "p"}}]},
        ).execute()
        self.executed = True

    @property
    def presentation_id(self) -> Optional[str]:
        """Returns the presentation_id of the created presentation.

        :raises RuntimeError: Must run the execute method before passing the presentation id
        :return: The presentation_id of the created presentation
        :rtype: str
        """
        if self.executed:
            return self.pr_id
        else:
            raise RuntimeError(
                "Must run the execute method before passing the presentation id"
            )


TLayout = TypeVar("TLayout", bound="Layout")


class Layout:
    def __init__(
        self,
        x_length: float,
        y_length: float,
        layout: Tuple[int, int],
        x_border: float = 0.05,
        y_border: float = 0.01,
        spacing: float = 0.02,
    ):
        """Constructor method"""
        self.x_length = x_length
        self.y_length = y_length
        self.x_objects = layout[0]
        self.y_objects = layout[1]
        self.x_border = x_border
        self.y_border = y_border
        self.spacing = spacing
        self.index = 0
        self.object_size = self._calc_size()
        validate_params_float(self.__dict__)

    def _calc_size(self) -> Tuple[float, float]:
        """Calculates the appropriate size of the chart given dimensions.

        :return: The x and y length of a chart
        :rtype: tuple

        """
        x_size = (
            self.x_length
            - self.x_length * ((self.y_objects - 1) * self.spacing + self.x_border * 2)
        ) / self.y_objects
        y_size = (
            self.y_length
            - self.y_length * ((self.x_objects - 1) * self.spacing + self.y_border * 2)
        ) / self.x_objects
        return (x_size, y_size)

    def __iter__(self: TLayout) -> TLayout:
        """Iterator function

        :return: :class:`Layout`
        :rtype: :class:`Layout`
        """
        return self

    @property
    def coord(self) -> Tuple[int, int]:
        """Calculates the row and column of the given index.

        :return: Row index and column index
        :rtype: tuple

        """
        x_coord = self.index // self.y_objects
        y_coord = self.index % self.y_objects
        return (x_coord, y_coord)

    def __next__(self) -> Tuple[float, float]:
        """Next function

        :return: A tuple for the translate x and translate y value
        :rtype: tuple
        """

        coord = self.coord
        translate_x = (
            self.x_length * self.x_border
            + coord[1] * self.object_size[0]
            + coord[1] * self.x_length * self.spacing
        )
        translate_y = (
            self.y_length * self.y_border
            + coord[0] * self.object_size[1]
            + coord[0] * self.y_length * self.spacing
        )
        if self.index + 1 >= self.x_objects * self.y_objects:
            self.index = 0
        else:
            self.index += 1
        return (translate_x, translate_y)


class Table:
    """The class that creates a table.

    :param data: Data to insert into the table
    :type data: SheetsFrame or pd.DataFrame
    :param font_size: Font size in the unit PT
    :type font_size: int
    :param header: Whether to enable formatting on the header row
    :type header: bool
    :param stub: Whether to enable formatting on the stub (1st) column
    :type stub: bool
    :param header_background_color: A color to set the header background.
        Parameters can either be a hex-code or a named color.
        See gslides.config.color_mapping.keys() for accepted named colors
    :type header_background_color: str
    :param stub_background_color: A color to set the stub background.
        Parameters can either be a hex-code or a named color.
        See gslides.config.color_mapping.keys() for accepted named colors
    :type stub_background_color: str
    """

    def __init__(
        self,
        data: Union[SheetsFrame, pd.DataFrame],
        font_size: int = 12,
        header: bool = True,
        stub: bool = False,
        header_background_color: str = "black",
        stub_background_color: str = "black",
    ) -> None:
        """Constructor method"""
        self.df = self._reset_header(self._resolve_df(data))
        self.font_size = 12
        self.header = header
        self.stub = stub
        self.header_background_color = hex_to_rgb(
            translate_color(header_background_color)
        )
        self.stub_background_color = hex_to_rgb(translate_color(stub_background_color))
        self.header_font_color = black_or_white(self.header_background_color)
        self.stub_font_color = black_or_white(self.stub_background_color)
        self.presentation_id = ""
        self.sl_id = ""
        self.translate_x = 0
        self.translate_y = 0
        self.size = (3000000, 3000000)

    def _resolve_df(self, data: Union[SheetsFrame, pd.DataFrame]):
        """Outputs a cleaned dataframe

        :param data: Data to insert into the table
        :type data: SheetsFrame or pd.DataFrame
        :raises ValueError: Only pd.DataFrame or SheetsFrame accepted
        :return: A cleaned dataframe
        :rtype: pd.DataFrame
        """
        if isinstance(data, SheetsFrame):
            return data.df
        elif isinstance(data, pd.DataFrame):
            df = clean_nan(data)
            df = df.applymap(clean_dtypes)
            return df
        else:
            raise ValueError("Only pd.DataFrame or SheetsFrame accepted")

    def _reset_header(self, df: pd.DataFrame):
        """Transforms a dataframe to set the 1st row as the header

        :param df: Dataframe to transform
        :type df: pd.DataFrame
        :return: Transformed dataframe
        :rtype: pd.DataFrame

        """
        return df.T.reset_index().T

    def render_create_table_json(self, sl_id: str) -> dict:
        """Renders the create table json

        :param sl_id: Slide id
        :type sl_id: str
        :return: json for the API call
        :rtype: dict

        """
        json: Dict[str, Any] = {
            "requests": [
                {
                    "createTable": {
                        "elementProperties": {
                            "pageObjectId": sl_id,
                        },
                        "rows": self.df.shape[0],
                        "columns": self.df.shape[1],
                    }
                },
            ]
        }
        return json

    def _table_move_request(
        self, tbl_id: str, translate_x: float, translate_y: float
    ) -> list:
        """Renders the move table requests

        :param tbl_id: Table id
        :type tbl_id: str
        :param translate_x: The number of EMU to translate the object by
        :type translate_x: float
        :param translate_y: The number of EMU to translate the object by
        :type translate_y: float
        :return: requests for the API call
        :rtype: list
        """
        requests: List[Any] = []
        requests.append(
            {
                "updatePageElementTransform": {
                    "objectId": tbl_id,
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": translate_x,
                        "translateY": translate_y,
                        "unit": "EMU",
                    },
                    "applyMode": "ABSOLUTE",
                },
            }
        )
        return requests

    def _table_add_text_request(self, tbl_id: str) -> list:
        """Renders the add text requests

        :param tbl_id: Table id
        :type tbl_id: str
        :return: requests for the API call
        :rtype: list
        """
        requests: List[Any] = []
        for row_cnt, col in enumerate(self.df.values.tolist()):
            for col_cnt, val in enumerate(col):
                requests.append(
                    {
                        "insertText": {
                            "objectId": tbl_id,
                            "cellLocation": {
                                "rowIndex": row_cnt,
                                "columnIndex": col_cnt,
                            },
                            "text": val,
                            "insertionIndex": 0,
                        }
                    }
                )
        return requests

    def _table_style_text_request(
        self,
        tbl_id: str,
        header: bool,
        stub: bool,
        font_size: int,
        header_font_color: Tuple[float, ...],
        stub_font_color: Tuple[float, ...],
    ) -> List:
        """Renders the style text requests

        :param tbl_id: Table id
        :type tbl_id: str
        :param header: Whether to enable formatting on the header row
        :type header: bool
        :param stub: Whether to enable formatting on the stub (1st) column
        :type stub: bool
        :param font_size: The size of font in the table
        :type font_size: int
        :param header_font_color: A color to set the header font.
        :type header_font_color: tuple
        :param stub_font_color: A color to set the stub font.
        :type stub_font_color: tuple
        :return: requests for the API call
        :rtype: list
        """
        requests: List[Any] = []
        for row_cnt, col in enumerate(self.df.values.tolist()):
            for col_cnt, val in enumerate(col):
                if header and row_cnt == 0:
                    font_color = header_font_color
                    bold = True
                elif stub and col_cnt == 0:
                    font_color = stub_font_color
                    bold = True
                else:
                    font_color = (0, 0, 0)
                    bold = False
                requests.append(
                    {
                        "updateTextStyle": {
                            "objectId": tbl_id,
                            "cellLocation": {
                                "rowIndex": row_cnt,
                                "columnIndex": col_cnt,
                            },
                            "style": {
                                "foregroundColor": {
                                    "opaqueColor": {
                                        "rgbColor": {
                                            "red": font_color[0],
                                            "green": font_color[1],
                                            "blue": font_color[2],
                                        }
                                    }
                                },
                                "bold": bold,
                                "fontFamily": "Arial",
                                "fontSize": {"magnitude": font_size, "unit": "PT"},
                            },
                            "textRange": {"type": "ALL"},
                            "fields": "foregroundColor,bold,fontFamily,fontSize",
                        }
                    }
                )
        return requests

    def _table_update_cell(
        self,
        tbl_id: str,
        header: bool,
        stub: bool,
        header_background_color: Tuple[float, ...],
        stub_background_color: Tuple[float, ...],
    ) -> List:
        """Renders the update cell requests

        :param tbl_id: Table id
        :type tbl_id: str
        :param header: Whether to enable formatting on the header row
        :type header: bool
        :param stub: Whether to enable formatting on the stub (1st) column
        :type stub: bool
        :param header_background_color: A color to set the header font.
        :type header_background_color: tuple
        :param stub_background_color: A color to set the stub font.
        :type stub_background_color: tuple
        :return: requests for the API call
        :rtype: list
        """
        requests: List[Any] = []
        requests.append(
            {
                "updateTableCellProperties": {
                    "objectId": tbl_id,
                    "tableRange": {
                        "location": {"rowIndex": 0, "columnIndex": 0},
                        "rowSpan": self.df.shape[0],
                        "columnSpan": self.df.shape[1],
                    },
                    "tableCellProperties": {"contentAlignment": "MIDDLE"},
                    "fields": "contentAlignment",
                }
            }
        )
        if stub:
            requests.append(
                {
                    "updateTableCellProperties": {
                        "objectId": tbl_id,
                        "tableRange": {
                            "location": {"rowIndex": 0, "columnIndex": 0},
                            "rowSpan": self.df.shape[0],
                            "columnSpan": 1,
                        },
                        "tableCellProperties": {
                            "tableCellBackgroundFill": {
                                "solidFill": {
                                    "color": {
                                        "rgbColor": {
                                            "red": stub_background_color[0],
                                            "green": stub_background_color[1],
                                            "blue": stub_background_color[2],
                                        }
                                    }
                                }
                            }
                        },
                        "fields": "tableCellBackgroundFill.solidFill.color",
                    }
                }
            )
        if header:
            requests.append(
                {
                    "updateTableCellProperties": {
                        "objectId": tbl_id,
                        "tableRange": {
                            "location": {"rowIndex": 0, "columnIndex": 0},
                            "rowSpan": 1,
                            "columnSpan": self.df.shape[1],
                        },
                        "tableCellProperties": {
                            "tableCellBackgroundFill": {
                                "solidFill": {
                                    "color": {
                                        "rgbColor": {
                                            "red": header_background_color[0],
                                            "green": header_background_color[1],
                                            "blue": header_background_color[2],
                                        }
                                    }
                                }
                            }
                        },
                        "fields": "tableCellBackgroundFill.solidFill.color",
                    }
                }
            )
        return requests

    def _table_update_paragraph_style(self, tbl_id: str) -> List:
        """Renders the update paragraph style requests

        :param tbl_id: Table id
        :type tbl_id: str
        :return: requests for the API call
        :rtype: list
        """
        requests: List[Any] = []
        for row_cnt, col in enumerate(self.df.tolist()):
            for col_cnt, val in enumerate(col):
                requests.append(
                    {
                        "updateParagraphStyle": {
                            "objectId": tbl_id,
                            "cellLocation": {
                                "rowIndex": row_cnt,
                                "columnIndex": col_cnt,
                            },
                            "style": {"alignment": "CENTER"},
                            "textRange": {"type": "ALL"},
                            "fields": "alignment",
                        }
                    }
                )
        return requests

    def _table_update_row(self, tbl_id: str, row_height: float) -> List:
        """Renders the update row height request

        :param tbl_id: Table id
        :type tbl_id: str
        :param row_height: The miminum height of a row. To note that this is simply
            the minimum height. Actual height of the row may be greater due to
            font size and text wrapping
        :type row_height: float
        :return: requests for the API call
        :rtype: list
        """
        requests: List[Any] = [
            {
                "updateTableRowProperties": {
                    "tableRowProperties": {
                        "minRowHeight": {"magnitude": row_height, "unit": "EMU"}
                    },
                    "rowIndices": None,
                    "objectId": tbl_id,
                    "fields": "minRowHeight",
                }
            }
        ]
        return requests

    def _table_update_column(self, tbl_id: str, col_widths: Tuple[float, ...]) -> list:
        """Renders the update column width request

        :param tbl_id: Table id
        :type tbl_id: str
        :param col_widths: The width of each column.
        :type col_widths: tuple
        :return: requests for the API call
        :rtype: list
        """
        requests: List[Any] = []
        for cnt, col_width in enumerate(col_widths):
            requests.append(
                {
                    "updateTableColumnProperties": {
                        "tableColumnProperties": {
                            "columnWidth": {"magnitude": col_width, "unit": "EMU"}
                        },
                        "columnIndices": [cnt],
                        "objectId": tbl_id,
                        "fields": "columnWidth",
                    }
                }
            )
        return requests

    def render_update_table_json(self, tbl_id: str) -> dict:
        """Renders the json for the update of table properties.

        :param tbl_id: Table id
        :type tbl_id: str
        :return: json for the API call
        :rtype: dict

        """
        json: Dict[str, Any] = {"requests": []}
        col_widths = self.size[0] * determine_col_proportion(self.df)
        row_height = self.size[1] / (self.df.shape[0])
        json["requests"].extend(
            self._table_move_request(tbl_id, self.translate_x, self.translate_y)
        )
        json["requests"].extend(self._table_add_text_request(tbl_id))
        json["requests"].extend(
            self._table_style_text_request(
                tbl_id,
                self.header,
                self.stub,
                self.font_size,
                self.header_font_color,
                self.stub_font_color,
            )
        )
        json["requests"].extend(
            self._table_update_cell(
                tbl_id,
                self.header,
                self.stub,
                self.header_background_color,
                self.stub_background_color,
            )
        )
        json["requests"].extend(self._table_update_paragraph_style(tbl_id))
        json["requests"].extend(self._table_update_row(tbl_id, row_height))
        json["requests"].extend(self._table_update_column(tbl_id, col_widths))
        return json

    def execute(self):
        """Executes the API call"""
        service = creds.slide_service
        output = (
            service.presentations()
            .batchUpdate(
                presentationId=self.presentation_id,
                body=self.render_create_table_json(self.sl_id),
            )
            .execute()
        )
        (
            service.presentations()
            .batchUpdate(
                presentationId=self.presentation_id,
                body=self.render_update_table_json(
                    output["replies"][0]["createTable"]["objectId"]
                ),
            )
            .execute()
        )


class CreateSlide:
    """The class that creates the presentation.

    :param presentation_id: The presentation_id of the created presentation
    :type presentation_id: str
    :param charts: :class:`Chart` objects that will be created
    :type charts: list
    :param layout: The layout of the chart objects in # of rows by # of columns
    :type layout: tuple
    :param insertion_index: The slide index to insert new slide to.
        The lack of a parameter will insert the slide to the end of the presentation
    :type insertion_index: int, optional
    :param top_margin: The top margin of the presentation in EMU
    :type top_margin: int, optional
    :param bottom_margin: The bottom margin of the presentation in EMU
    :type bottom_margin: int, optional
    :param left_margin: The left margin of the presentation in EMU
    :type left_margin: int, optional
    :param right_margin: The right margin of the presentation in EMU
    :type right_margin: int, optional
    """

    def __init__(
        self,
        presentation_id: str,
        objects: List[Chart],
        layout: Tuple[int, int],
        insertion_index: Optional[int] = None,
        top_margin: int = 1017724,
        bottom_margin: int = 420575,
        left_margin: int = 0,
        right_margin: int = 0,
        title: str = "Title placeholder",
        notes: str = "Notes placeholder",
    ) -> None:
        """Constructor method"""
        self.presentation_id = presentation_id
        self.objects = objects
        self.layout = self._validate_layout(layout)
        self.insertion_index = insertion_index
        self.sl_id: str = ""
        self.sheet_executed = False
        self.slide_executed = False
        self.top_margin = top_margin
        self.bottom_margin = bottom_margin
        self.left_margin = left_margin
        self.right_margin = right_margin
        self.layout_obj = Layout(
            9144000 - self.left_margin - self.right_margin,
            5143500 - self.top_margin - self.bottom_margin,
            layout,
        )
        self.title = title
        self.notes = notes

    def _validate_layout(self, layout: Tuple[int, int]) -> Tuple[int, int]:
        """Validates that the layout of charts is a valide layout.

        :param layout: The layout of the chart objects in # of rows by # of columns
        :type layout: tuple
        :raises ValueError:
        :return: The layout
        :rtype: tuple

        """
        if type(layout) != tuple:
            raise ValueError(
                "Provide tuple where the first index is the number of rows and "
                "the second index is the number of columns"
            )
        elif layout[0] < 1:
            raise ValueError(
                "Provide tuple where each value is an integer greater than 0"
            )
        elif layout[1] < 1:
            raise ValueError(
                "Provide tuple where each value is an integer greater than 0"
            )
        else:
            return layout

    def render_json_create_slide(self) -> dict:
        """Renders the json to create the slide in Google slides.

        :return: The json to do the update
        :rtype: dict
        """
        json: Dict[str, Any] = {
            "requests": [
                {"createSlide": {}},
            ]
        }
        if self.insertion_index:
            json["requests"][0]["createSlide"]["insertionIndex"] = self.insertion_index
        return json

    def render_json_create_textboxes(self, slide_id: str) -> dict:
        """Renders the json to create the textboxes in Google slides.

        :return: The json to do the update
        :rtype: dict
        """
        json = {
            "requests": [
                {
                    "createShape": {
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {
                                "width": {"magnitude": 3000000, "unit": "EMU"},
                                "height": {"magnitude": 3000000, "unit": "EMU"},
                            },
                            "transform": {
                                "scaleX": 2.8402,
                                "scaleY": 0.1909,
                                "translateX": 311700,
                                "translateY": 445025,
                                "unit": "EMU",
                            },
                        },
                    }
                },
                {
                    "createShape": {
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {
                                "width": {"magnitude": 3000000, "unit": "EMU"},
                                "height": {"magnitude": 3000000, "unit": "EMU"},
                            },
                            "transform": {
                                "scaleX": 2.7979,
                                "scaleY": 0.0914,
                                "translateX": 311700,
                                "translateY": 4722925,
                                "unit": "EMU",
                            },
                        },
                    }
                },
            ]
        }
        return json

    def render_json_format_textboxes(
        self, title_box_id: int, notes_box_id: int
    ) -> dict:
        """Renders the json to format the textboxes in Google slides.

        :return: The json to do the update
        :rtype: dict
        """
        json = {
            "requests": [
                {
                    "insertText": {
                        "objectId": title_box_id,
                        "insertionIndex": 0,
                        "text": self.title,
                    }
                },
                {
                    "updateTextStyle": {
                        "objectId": title_box_id,
                        "textRange": {"type": "ALL"},
                        "style": {
                            "bold": True,
                            "fontFamily": "Arial",
                            "fontSize": {"magnitude": 24, "unit": "PT"},
                        },
                        "fields": "bold,fontFamily,fontSize",
                    }
                },
                {
                    "updateShapeProperties": {
                        "objectId": title_box_id,
                        "fields": "contentAlignment",
                        "shapeProperties": {"contentAlignment": "MIDDLE"},
                    }
                },
                {
                    "insertText": {
                        "objectId": notes_box_id,
                        "insertionIndex": 0,
                        "text": self.notes,
                    }
                },
                {
                    "updateTextStyle": {
                        "objectId": notes_box_id,
                        "textRange": {"type": "ALL"},
                        "style": {
                            "bold": True,
                            "fontFamily": "Arial",
                            "fontSize": {"magnitude": 7, "unit": "PT"},
                        },
                        "fields": "bold,fontFamily,fontSize",
                    }
                },
            ]
        }
        return json

    def render_json_copy_chart(
        self,
        chart: Chart,
        size: Tuple[float, float],
        translate_x: float,
        translate_y: float,
    ) -> dict:
        """Renders the json to copy the charts in Google slides.

        :return: The json to do the update
        :rtype: dict
        """
        json = {
            "createSheetsChart": {
                "spreadsheetId": chart.data.spreadsheet_id,
                "chartId": chart.chart_id,
                "linkingMode": "LINKED",
                "elementProperties": {
                    "pageObjectId": self.sl_id,
                    "size": {
                        "width": {"magnitude": size[0], "unit": "EMU"},
                        "height": {"magnitude": size[1], "unit": "EMU"},
                    },
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": self.left_margin + translate_x,
                        "translateY": self.top_margin + translate_y,
                        "unit": "EMU",
                    },
                },
            }
        }
        return json

    def _execute_create_slide(self) -> None:
        """Executes the create slides API call."""
        service: Any = creds.slide_service
        output = (
            service.presentations()
            .batchUpdate(
                presentationId=self.presentation_id,
                body=self.render_json_create_slide(),
            )
            .execute()
        )
        self.sl_id = output["replies"][0]["createSlide"]["objectId"]

    def _execute_create_format_textboxes(self) -> None:
        """Executes the create & format textboxes slides API call."""
        service: Any = creds.slide_service
        output = (
            service.presentations()
            .batchUpdate(
                presentationId=self.presentation_id,
                body=self.render_json_create_textboxes(self.sl_id),
            )
            .execute()
        )
        self.title_bx_id = output["replies"][0]["createShape"]["objectId"]
        self.notes_bx_id = output["replies"][1]["createShape"]["objectId"]
        output = (
            service.presentations()
            .batchUpdate(
                presentationId=self.presentation_id,
                body=self.render_json_format_textboxes(
                    self.title_bx_id, self.notes_bx_id
                ),
            )
            .execute()
        )

    def _execute_populate_objects(self):
        """Executes the population of objects on the slide"""
        json: Dict[str, Any] = {"requests": []}
        service = creds.slide_service
        for obj in self.objects:
            translate_x, translate_y = next(self.layout_obj)
            if obj.__class__.__name__ == "Chart":
                json["requests"].append(
                    self.render_json_copy_chart(
                        obj, self.layout_obj.object_size, translate_x, translate_y
                    )
                )
                (
                    service.presentations()
                    .batchUpdate(presentationId=self.presentation_id, body=json)
                    .execute()
                )
            elif obj.__class__.__name__ == "Table":
                obj.presentation_id = self.presentation_id
                obj.sl_id = self.sl_id
                obj.size = self.layout_obj.object_size
                obj.translate_x = self.left_margin + translate_x
                obj.translate_y = self.top_margin + translate_y
                obj.execute()

    def execute_slide(self) -> None:
        """Executes the slides API call."""
        if self.sheet_executed is False:
            raise RuntimeError(
                "Must run the execute sheet method before running the execute slide method"
            )
        self._execute_create_slide()
        self._execute_create_format_textboxes()
        self._execute_populate_objects()
        self.slide_executed = True

    def execute_sheet(self) -> None:
        """Executes the sheets API call."""
        x_len, y_len = optimize_size(
            self.layout_obj.object_size[1] / self.layout_obj.object_size[0],
            area=222600 / (self.layout[0] * self.layout[1]),
        )
        for obj in self.objects:
            if obj.__class__.__name__ == "Chart":
                obj.size = (int(x_len), int(y_len))
                obj.execute()
        self.sheet_executed = True

    def execute(self) -> None:
        """Executes the sheets & slides API call."""
        self.execute_sheet()
        self.execute_slide()
