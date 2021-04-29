# -*- coding: utf-8 -*-
"""
Manipulates data in google sheets
"""

from typing import Any, List, Optional, Tuple, TypeVar, cast

import pandas as pd

from . import creds
from .utils import (
    cell_to_num,
    clean_dtypes,
    clean_list_of_list,
    clean_nan,
    json_chunk_extract,
    json_val_extract,
    num_to_char,
    validate_cell_name,
)


TSheetsFrame = TypeVar("TSheetsFrame", bound="SheetsFrame")


class SheetsFrame:
    """A table of data in Google sheets

    :param spreadsheet_id: The id associated with the spreadsheet
    :type spreadsheet_id: str
    :param sheet_id: The id associated with the sheet
    :type sheet_id: int
    :param start_row_index: The starting index of the row
    :type start_row_index: int
    :param end_row_index: The ending index of the row
    :type end_row_index: int
    :param start_column_index: The starting index of the column
    :type start_column_index: int
    :param end_column_index: The ending index of the column
    :type end_column_index: int
    :param df: Dataframe representation of the data
    :type df: :class:`pandas.DataFrame`
    """

    def __init__(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        start_row_index: int,
        start_column_index: int,
        end_row_index: int,
        end_column_index: int,
        df: pd.DataFrame,
    ) -> None:
        """Constructor method"""
        self.executed = False
        self.spreadsheet_id = spreadsheet_id
        self.sheet_id = sheet_id
        self.start_row_index = start_row_index
        self.start_column_index = start_column_index
        self.end_row_index = end_row_index
        self.end_column_index = end_column_index
        self.df = df

    @property
    def data(self: TSheetsFrame) -> TSheetsFrame:
        """Returns the :class:`SheetsFrame` object of the data

        :raises RuntimeError: Must run the execute method before passing the data
        :return: The :class:`SheetsFrame` object
        :rtype: :class:`SheetsFrame`
        """
        if self.executed:
            return self
        else:
            raise RuntimeError("Must run the execute method before passing the data")

    def _get_sheet_name(self) -> str:
        """Gets the name of a sheet

        :return: The name associated with the sheet
        :rtype: str
        """
        service: Any = creds.sheet_service
        get_output = (
            service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
        )
        sheet_name = cast(
            str, json_chunk_extract(get_output, "sheetId", self.sheet_id)[0]["title"]
        )
        return sheet_name

    def _get_sheet_data(self, sheet_name: str) -> List[List]:
        """Gets the data from a given groups of cells in a sheet

        :return: A list of lists capturing the data
        :rtype: list
        """
        service: Any = creds.sheet_service
        rng = (
            f"{sheet_name}!{num_to_char(self.start_column_index)}"
            f"{self.start_row_index}:"
            f"{num_to_char(self.end_column_index)}{self.end_row_index}"
        )
        output = (
            service.spreadsheets()
            .values()
            .get(spreadsheetId=self.spreadsheet_id, range=rng)
            .execute()
        )
        if "values" in output.keys():
            return cast(List[List], output["values"])
        else:
            return [[]]


class CreateFrame(SheetsFrame):
    """Class to create data in Google sheets.

    :param df: Data to be created in Google sheets
    :type df: :class:`pandas.DataFrame`
    :param spreadsheet_id: The id associated with the spreadsheet
    :type spreadsheet_id: str
    :param sheet_id: The id associated with the sheet
    :type sheet_id: int
    :param start_row_index: The starting index of the row
    :param overwrite_data: Whether to overwrite the existing data
    :type overwrite_data: bool, optional
    :param anchor_cell: The cell name (e.g. `A5`) that will correspond to the
        top left observation in the dataframe
    :type anchor_cell: str, optional
    """

    def __init__(
        self,
        df: pd.DataFrame,
        spreadsheet_id: str,
        sheet_id: int,
        overwrite_data: bool = False,
        anchor_cell: str = "A1",
    ) -> None:
        """Constructor method"""
        self.df = df
        self.spreadsheet_id = spreadsheet_id
        self.sheet_id = sheet_id
        self.overwrite_data = overwrite_data
        self.anchor_cell = validate_cell_name(anchor_cell.upper())
        self.start_row_index, self.start_column_index = cell_to_num(self.anchor_cell)
        self.end_row_index, self.end_column_index = self._calc_end_index()
        self._clean_df()
        super().__init__(
            spreadsheet_id,
            sheet_id,
            self.start_row_index,
            self.start_column_index,
            self.end_row_index,
            self.end_column_index,
            self.df,
        )

    def _calc_end_index(self) -> Tuple[int, int]:
        """Calculates the ending row and column index based on the anchor cell
        and dataframe size

        :return: Tuple of the ending row and column inde
        :rtype: tuple
        """
        end_row_index = self.start_row_index + self.df.shape[0] + 1
        end_column_index = self.start_column_index + self.df.shape[1]
        return (end_row_index, end_column_index)

    def _clean_df(self) -> None:
        """Cleans the dataframe to convert datatypes into acceptable values for
        the Google Sheets API

        :return: Cleaned :class:`pandas.DataFrame`
        :type df: :class:`pandas.DataFrame`
        """
        self.df = clean_nan(self.df)
        self.df = self.df.applymap(clean_dtypes)

    def render_update_json(self, sheet_name: str) -> dict:
        """Renders the json to update the data in Google sheets

        :param sheet_name: The name of the sheet
        :type sheet_name: str
        :return: The json to do the update
        :rtype: dict
        """
        col_range = (
            f"{sheet_name}!{num_to_char(self.start_column_index)}"
            f"{self.start_row_index}:"
            f"{num_to_char(self.end_column_index)}{self.start_row_index}"
        )
        val_range = (
            f"{sheet_name}!{num_to_char(self.start_column_index)}"
            f"{self.start_row_index+1}:"
            f"{num_to_char(self.end_column_index)}{self.end_row_index}"
        )
        json = {
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": col_range, "values": [self.df.columns.tolist()]},
                {"range": val_range, "values": self.df.values.tolist()},
            ],
        }
        return json

    def execute(self) -> None:
        """Executes the API call

        :return: The json returned by the call
        :rtype: dict

        """
        service: Any = creds.sheet_service
        sheet_name = self._get_sheet_name()
        json = self.render_update_json(sheet_name)
        if self.overwrite_data is False:
            existing_data = self._get_sheet_data(sheet_name)
            if existing_data:
                raise RuntimeError("Create table will overwrite existing data")
        (
            service.spreadsheets()
            .values()
            .batchUpdate(spreadsheetId=self.spreadsheet_id, body=json)
            .execute()
        )
        self.executed = True


class GetFrame(SheetsFrame):
    """Class to get data from Google sheets.

    :param spreadsheet_id: The id associated with the spreadsheet
    :type spreadsheet_id: str
    :param sheet_id: The id associated with the sheet
    :type sheet_id: int
    :param anchor_cell: The cell name (e.g. `A5`) that will correspond to the
        top left observation in the dataframe
    :type anchor_cell: str
    :param bottom_right_cell: The cell name (e.g. `B10`) that will correspond to the
        bottom right observation in the dataframe
    :type bottom_right_cell: str
    """

    def __init__(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        anchor_cell: str,
        bottom_right_cell: str,
    ) -> None:
        """Constructor method"""
        self.spreadsheet_id = spreadsheet_id
        self.sheet_id = sheet_id
        self.anchor_cell = validate_cell_name(anchor_cell.upper())
        self.bottom_right_cell = validate_cell_name(bottom_right_cell.upper())
        self.start_row_index, self.start_column_index = cell_to_num(self.anchor_cell)
        self.end_row_index, self.end_column_index = cell_to_num(self.bottom_right_cell)
        self.df: pd.DataFrame = pd.DataFrame()
        super().__init__(
            spreadsheet_id,
            sheet_id,
            self.start_row_index,
            self.start_column_index,
            self.end_row_index,
            self.end_column_index,
            self.df,
        )

    def execute(self) -> None:
        """Executes the API call

        :return: The json returned by the call
        :rtype: dict

        """
        sheet_name = self._get_sheet_name()
        output = self._get_sheet_data(sheet_name)
        output = clean_list_of_list(output)
        self.df = pd.DataFrame(data=output[1:], columns=output[0])
        self.df = self.df.replace("", None)
        self.executed = True


class CreateSpreadsheet:
    """An object to create a spreadsheet in Google sheets

    :param title: The title of the spreadsheet
    :type title: str, optional
    :param sheet_name: The name of the sheet
    :type title: str, optional
    """

    def __init__(self, title: str = "Untitled", sheet_name: str = "Sheet1") -> None:
        """Constructor method"""
        self.title = title
        self.sheet_name = sheet_name
        self.executed = False
        self.sp_id: Optional[str] = None
        self.sh_id: Optional[int] = None

    def render_json(self) -> dict:
        """Renders the json to create the spreadsheet in Google sheets

        :return: The json to do the update
        :rtype: dict
        """
        json = {
            "properties": {
                "title": self.title,
                "locale": "en_US",
                "autoRecalc": "HOUR",
            },
            "sheets": [{"properties": {"title": self.sheet_name}}],
        }
        return json

    def execute(self) -> None:
        """Executes the API call

        :return: The json returned by the call
        :rtype: dict

        """
        service: Any = creds.sheet_service
        output = service.spreadsheets().create(body=self.render_json()).execute()
        self.sp_id = cast(Optional[str], json_val_extract(output, "spreadsheetId"))
        self.sh_id = cast(Optional[int], json_val_extract(output, "sheetId"))
        self.executed = True

    @property
    def spreadsheet_id(self) -> Optional[str]:
        """Returns the spreadsheet_id of the created spreadhseet.

        :raises RuntimeError: Must run the execute method before obtaining the id
        :return: The spreadsheet_id of the created spreadsheet.
        :rtype: str
        """
        if self.executed:
            return self.sp_id
        else:
            raise RuntimeError("Must run the execute method before obtaining the id")

    @property
    def sheet_id(self) -> Optional[int]:
        """Returns the sheet_id of the created spreadhseet.

        :raises RuntimeError: Must run the execute method before obtaining the id
        :return: The sheet_id of the created sheet.
        :rtype: int
        """
        if self.executed:
            return self.sh_id
        else:
            raise RuntimeError("Must run the execute method before passing the id")


class CreateSheet:
    """Creates a new sheet in an existing spreadsheet

    :param title: The title of the spreadsheet
    :type title: str, optional
    :param sheet_name: The name of the sheet
    :type title: str, optional
    """

    def __init__(self, spreadsheet_id: str, sheet_name: str):
        """Constructor method"""
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.executed = False
        self.sh_id: Optional[int] = None

    def render_json(self) -> dict:
        """Renders the json to create the sheet in Google sheets

        :return: The json to do the update
        :rtype: dict
        """
        json = {"requests": [{"addSheet": {"properties": {"title": self.sheet_name}}}]}
        return json

    def execute(self) -> None:
        """Executes the API call"""
        service: Any = creds.sheet_service
        output = (
            service.spreadsheets()
            .batchUpdate(spreadsheetId=self.spreadsheet_id, body=self.render_json())
            .execute()
        )
        self.sh_id = json_val_extract(output, "sheetId")
        self.executed = True

    @property
    def sheet_id(self) -> Optional[int]:
        """Returns the sheet_id of the created spreadhseet.

        :raises RuntimeError: Must run the execute method before obtaining the id
        :return: The sheet_id of the created sheet.
        :rtype: int
        """
        if self.executed:
            return self.sh_id
        else:
            raise RuntimeError("Must run the execute method before passing the id")
