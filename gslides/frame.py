# -*- coding: utf-8 -*-
"""
Frame class
"""

from typing import Any, List, Tuple, Type, TypeVar, cast

import pandas as pd

from . import creds
from .utils import (
    cell_to_num,
    clean_dtypes,
    clean_list_of_list,
    clean_nan,
    num_to_char,
    validate_cell_name,
)


TFrame = TypeVar("TFrame", bound="Frame")


def get_sheet_data(
    spreadsheet_id,
    sheet_name: str,
    start_column_index: int,
    start_row_index: int,
    end_column_index: int,
    end_row_index: int,
) -> List[List]:

    """Gets the data from a given groups of cells in a sheet

    :param spreadsheet_id: The id of the spreadsheet
    :type spreadsheet_id: str
    :param sheet_name: Sheet name to get data from
    :type sheet_name: str
    :param start_column_index: The index of the starting column of the data
    :type start_column_index: int
    :param start_row_index: The index of the starting row of the data
    :type start_row_index: int
    :param end_column_index: The index of the ending column of the data
    :type end_column_index: int
    :param end_row_index: The index of the ending row of the data
    :type end_row_index: int
    :return: A list of lists capturing the data
    :rtype: list
    """
    service: Any = creds.sheet_service
    rng = (
        f"{sheet_name}!{num_to_char(start_column_index)}"
        f"{start_row_index}:"
        f"{num_to_char(end_column_index)}{end_row_index}"
    )
    output = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=rng)
        .execute()
    )
    if "values" in output.keys():
        return cast(List[List], output["values"])
    else:
        return [[]]


class CreateFrame:
    """Class to create data in Google sheets.

    :param df: Data to be created in Google sheets
    :type df: :class:`pandas.DataFrame`
    :param spreadsheet_id: The id associated with the spreadsheet
    :type spreadsheet_id: str
    :param sheet_name: The name associated with the sheet
    :type sheet_name: str
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
        sheet_name: str,
        overwrite_data: bool = False,
        anchor_cell: str = "A1",
    ) -> None:
        """Constructor method"""
        self.df = df
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.overwrite_data = overwrite_data
        self.anchor_cell = validate_cell_name(anchor_cell.upper())
        self.start_row_index, self.start_column_index = cell_to_num(self.anchor_cell)
        self.end_row_index, self.end_column_index = self._calc_end_index()
        self._clean_df()

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

    def render_update_json(self) -> dict:
        """Renders the json to update the data in Google sheets

        :return: The json to do the update
        :rtype: dict
        """
        col_range = (
            f"{self.sheet_name}!{num_to_char(self.start_column_index)}"
            f"{self.start_row_index}:"
            f"{num_to_char(self.end_column_index)}{self.start_row_index}"
        )
        val_range = (
            f"{self.sheet_name}!{num_to_char(self.start_column_index)}"
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

    def execute(self) -> bool:
        """Executes the API call

        :return: Whether the function executed
        :rtype: bool
        """

        service: Any = creds.sheet_service
        json = self.render_update_json()
        if self.overwrite_data is False:
            existing_data = get_sheet_data(
                self.spreadsheet_id,
                self.sheet_name,
                self.start_column_index,
                self.start_row_index,
                self.end_column_index,
                self.end_row_index,
            )
            if existing_data:
                raise RuntimeError("Create table will overwrite existing data")
        (
            service.spreadsheets()
            .values()
            .batchUpdate(spreadsheetId=self.spreadsheet_id, body=json)
            .execute()
        )
        return True


class GetFrame:
    """Class to get data from Google sheets.

    :param spreadsheet_id: The id associated with the spreadsheet
    :type spreadsheet_id: str
    :param sheet_name: The name associated with the sheet
    :type sheet_name: str
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
        sheet_name: str,
        anchor_cell: str,
        bottom_right_cell: str,
    ) -> None:
        """Constructor method"""
        self.spreadsheet_id = spreadsheet_id
        self.sheet_id = sheet_id
        self.sheet_name = sheet_name
        self.anchor_cell = validate_cell_name(anchor_cell.upper())
        self.bottom_right_cell = validate_cell_name(bottom_right_cell.upper())
        self.start_row_index, self.start_column_index = cell_to_num(self.anchor_cell)
        self.end_row_index, self.end_column_index = cell_to_num(self.bottom_right_cell)
        self.df: pd.DataFrame = pd.DataFrame()

    def execute(self) -> bool:
        """Executes the API call

        :return: Whether the function executed
        :rtype: bool
        """
        output = get_sheet_data(
            self.spreadsheet_id,
            self.sheet_name,
            self.start_column_index,
            self.start_row_index,
            self.end_column_index,
            self.end_row_index,
        )
        output = clean_list_of_list(output)
        self.df = pd.DataFrame(data=output[1:], columns=output[0])
        self.df = self.df.replace("", None)
        return True


class Frame:
    """An object that represents a table of data in Google sheets. Initialize the
    object through either the :class:`Frame.get or :class:`Frame.create` class method.

    :param title: The title of the spreadsheet
    :type title: str, optional
    :param sheet_name: The name of the sheet
    :type title: str, optional

    :param df: The dataframe
    :type df: :class:`pd.DataFrame`, optional
    :param spreadsheet_id: The id of the spreadsheet
    :type spreadsheet_id: str, optional
    :param sheet_id: The id associated with the sheet
    :type sheet_id: int, optional
    :param sheet_name: The name associated with the sheet
    :type sheet_name: str, optional
    :param start_column_index: The index of the starting column of the data
    :type start_column_index: int
    :param start_row_index: The index of the starting row of the data
    :type start_row_index: int
    :param end_column_index: The index of the ending column of the data
    :type end_column_index: int
    :param end_row_index: The index of the ending row of the data
    :type end_row_index: int
    :param initialized: A boolean whether the class has been initialized
    :type initialized: bool, optional
    """

    def __init__(
        self,
        df: pd.DataFrame = pd.DataFrame(),
        spreadsheet_id: str = "",
        sheet_id: int = 0,
        sheet_name: str = "",
        start_column_index: int = 0,
        start_row_index: int = 0,
        end_column_index: int = 0,
        end_row_index: int = 0,
        initialized: bool = False,
    ) -> None:
        """Constructor method"""
        self.df = df
        self.spreadsheet_id = spreadsheet_id
        self.sheet_id = sheet_id
        self.sheet_name = sheet_name
        self.start_column_index = start_column_index
        self.start_row_index = start_row_index
        self.end_column_index = end_column_index
        self.end_row_index = end_row_index
        self.initialized = initialized

    @classmethod
    def create(
        cls: Type[TFrame],
        df: pd.DataFrame,
        spreadsheet_id: str,
        sheet_id: int,
        sheet_name: str,
        overwrite_data: bool = False,
        anchor_cell: str = "A1",
    ) -> TFrame:
        """Creates the table of data in Google sheets

        :param df: The dataframe
        :type df: :class:`pd.DataFrame`
        :param spreadsheet_id: The id of the spreadsheet
        :type spreadsheet_id: str
        :param sheet_id: The id associated with the sheet
        :type sheet_id: int
        :param sheet_name: The name associated with the sheet
        :type sheet_name: str
        :param overwrite_data: Whether to overwrite the existing data
        :type overwrite_data: bool, optional
        :param anchor_cell: The cell name (e.g. `A5`) that will correspond to the
            top left observation in the dataframe
        :type anchor_cell: str
        :return: :class:`gslides.Frame` object
        :rtype: :class:`gslides.Frame`

        """
        frame = CreateFrame(
            df,
            spreadsheet_id,
            sheet_name,
            overwrite_data=overwrite_data,
            anchor_cell=anchor_cell,
        )
        initialized = frame.execute()
        return cls(
            df,
            spreadsheet_id,
            sheet_id,
            sheet_name,
            frame.start_column_index,
            frame.start_row_index,
            frame.end_column_index,
            frame.end_row_index,
            initialized,
        )

    @classmethod
    def get(
        cls: Type[TFrame],
        spreadsheet_id: str,
        sheet_id: int,
        sheet_name: str,
        anchor_cell: str,
        bottom_right_cell: str,
    ) -> TFrame:
        """Gets the table of data in Google sheets

        :param spreadsheet_id: The id of the spreadsheet
        :type spreadsheet_id: str
        :param sheet_id: The id associated with the sheet
        :type sheet_id: int
        :param sheet_name: The name associated with the sheet
        :type sheet_name: str
        :param anchor_cell: The cell name (e.g. `A5`) that will correspond to the
            top left observation in the dataframe
        :type anchor_cell: str
        :param bottom_right_cell: The cell name (e.g. `B7`) that will correspond
            to the bottom right observation in the dataframe
        :type bottom_right_cell: str
        :return: :class:`gslides.Frame` object
        :rtype: :class:`gslides.Frame`

        """
        frame = GetFrame(
            spreadsheet_id,
            sheet_id,
            sheet_name,
            anchor_cell,
            bottom_right_cell,
        )
        initialized = frame.execute()
        return cls(
            frame.df,
            spreadsheet_id,
            sheet_id,
            sheet_name,
            frame.start_column_index,
            frame.start_row_index,
            frame.end_column_index,
            frame.end_row_index,
            initialized,
        )

    @property
    def data(self: TFrame) -> TFrame:
        """Returns the :class:`Frame` object of the data

        :raises Must run the create or get method before passing the data
        :return: The :class:`Frame` object
        :rtype: :class:`Frame`
        """
        if self.initialized:
            return self
        else:
            raise RuntimeError(
                "Must run the create or get method before passing the data"
            )