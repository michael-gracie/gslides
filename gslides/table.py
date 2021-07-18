# -*- coding: utf-8 -*-
"""
Creates the table in Google slides
"""
import logging
import pprint
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd

from . import creds, package_font
from .colors import translate_color
from .frame import Frame
from .utils import (
    black_or_white,
    clean_dtypes,
    clean_nan,
    determine_col_proportion,
    hex_to_rgb,
)

logger = logging.getLogger(__name__)


class Table:
    """The class that creates a table.

    :param data: Data to insert into the table
    :type data: Frame or pd.DataFrame
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
    :param column_proportions: A list of floats representing the proportion of each
        column
    :type column_proportions:  list of floats
    """

    def __init__(
        self,
        data: Union[Frame, pd.DataFrame],
        font_size: int = 12,
        header: bool = True,
        stub: bool = False,
        header_background_color: str = "black",
        stub_background_color: str = "black",
        column_proportions: Optional[List[float]] = None,
    ) -> None:
        """Constructor method"""
        self.df = self._reset_header(self._resolve_df(data))
        self.font_size = font_size
        self.header = header
        self.stub = stub
        self.header_background_color = hex_to_rgb(
            translate_color(header_background_color)
        )
        self.stub_background_color = hex_to_rgb(translate_color(stub_background_color))
        self.header_font_color = black_or_white(self.header_background_color)
        self.stub_font_color = black_or_white(self.stub_background_color)
        self.column_proportions = column_proportions

    def __repr__(self) -> str:
        """Prints class information.

        :return: String with helpful class infromation
        :rtype: str

        """
        output = f"Table\n" f"{self.df.to_markdown(index = False)}"
        return output

    def _resolve_df(self, data: Union[Frame, pd.DataFrame]):
        """Outputs a cleaned dataframe

        :param data: Data to insert into the table
        :type data: Frame or pd.DataFrame
        :raises ValueError: Only pd.DataFrame or Frame accepted
        :return: A cleaned dataframe
        :rtype: pd.DataFrame
        """
        if isinstance(data, Frame):
            return data.df
        elif isinstance(data, pd.DataFrame):
            df = clean_nan(data)
            df = df.applymap(clean_dtypes)
            return df
        else:
            raise ValueError("Only pd.DataFrame or Frame accepted")

    def _reset_header(self, df: pd.DataFrame):
        """Transforms a dataframe to set the 1st row as the header

        :param df: Dataframe to transform
        :type df: pd.DataFrame
        :return: Transformed dataframe
        :rtype: pd.DataFrame

        """
        return df.T.reset_index().T

    def render_create_table_json(self, sl_id: str) -> Dict[str, Any]:
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
    ) -> List[Any]:
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

    def _table_add_text_request(self, tbl_id: str) -> List[Any]:
        """Renders the add text requests

        :param tbl_id: Table id
        :type tbl_id: str
        :return: requests for the API call
        :rtype: list
        """
        requests: List[Any] = []
        for row_cnt, col in enumerate(
            self.df.applymap(str).values.tolist()
        ):  # fix_later
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
    ) -> List[Any]:
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
                                "fontFamily": package_font.font,
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
    ) -> List[Any]:
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

    def _table_update_paragraph_style(self, tbl_id: str) -> List[Any]:
        """Renders the update paragraph style requests

        :param tbl_id: Table id
        :type tbl_id: str
        :return: requests for the API call
        :rtype: list
        """
        requests: List[Any] = []
        for row_cnt, col in enumerate(self.df.values.tolist()):  # fix later
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

    def render_update_table_json(
        self,
        tbl_id: str,
        size: Tuple[float, float],
        translate_x: float,
        translate_y: float,
    ) -> dict:
        """Renders the json for the update of table properties.

        :param tbl_id: Table id
        :type tbl_id: str
        :param size: Tuple of width and height in EMU
        :type size: tuple
        :param translate_x: The number of EMU to translate the object by
        :type translate_x: float
        :param translate_y: The number of EMU to translate the object by
        :type translate_y: float
        :return: json for the API call
        :rtype: dict

        """
        json: Dict[str, Any] = {"requests": []}
        if self.column_proportions:
            col_widths = size[0] * np.array(self.column_proportions)
        else:
            col_widths = size[0] * determine_col_proportion(self.df)
        row_height = size[1] / (self.df.shape[0])
        json["requests"].extend(
            self._table_move_request(tbl_id, translate_x, translate_y)
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

    def create(
        self,
        presentation_id: str,
        slide_id: str,
        size: Tuple[float, float] = (3000000, 3000000),
        translate_x: float = 0,
        translate_y: float = 0,
    ) -> None:
        """Creates the table in Google slides

        :param presentation_id: The presentation_id of the presentation to create
            to create table in
        :type presentation_id: str
        :param slide_id: The slide_id of the slide to create table in
        :type slide_id: str
        :param size: Tuple of width and height in EMU
        :type size: tuple
        :param translate_x: The number of EMU to translate the object by
        :type translate_x: float
        :param translate_y: The number of EMU to translate the object by
        :type translate_y: float
        """
        service = creds.slide_service
        body = self.render_create_table_json(slide_id)
        logger.info("Executing table creation")
        logger.info(f"Request: {pprint.pformat(body)}")
        output = (
            service.presentations()
            .batchUpdate(
                presentationId=presentation_id,
                body=body,
            )
            .execute()
        )
        logger.info("Table created successfully")
        body = self.render_update_table_json(
            output["replies"][0]["createTable"]["objectId"],
            size,
            translate_x,
            translate_y,
        )
        logger.info("Executing table updates")
        logger.info(f"Request: {pprint.pformat(body)}")
        (
            service.presentations()
            .batchUpdate(
                presentationId=presentation_id,
                body=body,
            )
            .execute()
        )
        logger.info("Table updated successfully")
