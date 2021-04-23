# -*- coding: utf-8 -*-
from typing import TYPE_CHECKING, List, Optional, Tuple, cast

import pandas as pd

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


if TYPE_CHECKING:
    from googleapiclient.discovery import Resource


class SheetsFrame:
    def __init__(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        start_row_index: int,
        start_column_index: int,
        end_row_index: int,
        end_column_index: int,
    ) -> None:
        self.executed = False
        self.spreadsheet_id = spreadsheet_id
        self.sheet_id = sheet_id
        self.start_row_index = start_row_index
        self.start_column_index = start_column_index
        self.end_row_index = end_row_index
        self.end_column_index = end_column_index

    @property
    def data(self) -> SheetsFrame:  # noqa
        if self.executed:
            return self
        else:
            raise RuntimeError("Must run the execute method before passing the data")

    def get_sheet_name(self, service: Resource) -> str:
        get_output = (
            service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
        )
        sheet_name = cast(
            str, json_chunk_extract(get_output, "sheetId", self.sheet_id)[0]["title"]
        )
        return sheet_name

    def get_sheet_data(self, service: Resource, sheet_name: str) -> List[List]:
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
    def __init__(
        self,
        df: pd.DataFrame,
        spreadsheet_id: str,
        sheet_id: int,
        overwrite_data: bool = False,
        anchor_cell: str = "A1",
    ) -> None:
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
        )

    def _calc_end_index(self) -> Tuple[int, int]:
        end_row_index = self.start_row_index + self.df.shape[0] + 1
        end_column_index = self.start_column_index + self.df.shape[1]
        return (end_row_index, end_column_index)

    def _clean_df(self) -> None:
        self.df = clean_nan(self.df)
        self.df = self.df.applymap(clean_dtypes)

    def render_update_json(self, sheet_name: str) -> dict:
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

    def execute(self, service: Resource) -> None:
        sheet_name = self.get_sheet_name(service)
        json = self.render_update_json(sheet_name)
        if self.overwrite_data is False:
            existing_data = self.get_sheet_data(service, sheet_name)
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
    def __init__(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        anchor_cell: str,
        bottom_right_cell: str,
    ) -> None:
        self.spreadsheet_id = spreadsheet_id
        self.sheet_id = sheet_id
        self.anchor_cell = validate_cell_name(anchor_cell.upper())
        self.bottom_right_cell = validate_cell_name(bottom_right_cell.upper())
        self.start_row_index, self.start_column_index = cell_to_num(self.anchor_cell)
        self.end_row_index, self.end_column_index = cell_to_num(self.bottom_right_cell)
        self.df: Optional[pd.DataFrame] = None
        super().__init__(
            spreadsheet_id,
            sheet_id,
            self.start_row_index,
            self.start_column_index,
            self.end_row_index,
            self.end_column_index,
        )

    def execute(self, service: Resource) -> None:
        sheet_name = self.get_sheet_name(service)
        output = self.get_sheet_data(service, sheet_name)
        output = clean_list_of_list(output)
        self.df = pd.DataFrame(data=output[1:], columns=output[0])
        self.df = self.df.replace("", None)
        self.executed = True


class CreateSheet:
    def __init__(self, title: str = "Untitled", tab_name: str = "Sheet1") -> None:
        self.title = title
        self.tab_name = tab_name
        self.executed = False
        self.sp_id: Optional[str] = None
        self.sh_id: Optional[int] = None

    def render_json(self) -> dict:
        json = {
            "properties": {
                "title": self.title,
                "locale": "en_US",
                "autoRecalc": "HOUR",
            },
            "sheets": [{"properties": {"title": self.tab_name}}],
        }
        return json

    def execute(self, service: Resource) -> None:
        output = service.spreadsheets().create(body=self.render_json()).execute()
        self.sp_id = json_val_extract(output, "spreadsheetId")
        self.sh_id = json_val_extract(output, "sheetId")
        self.executed = True

    @property
    def spreadsheet_id(self) -> Optional[str]:
        if self.executed:
            return self.sp_id
        else:
            raise RuntimeError("Must run the execute method before obtaining the id")

    @property
    def sheet_id(self) -> Optional[int]:
        if self.executed:
            return self.sh_id
        else:
            raise RuntimeError("Must run the execute method before passing the id")


class CreateTab:
    def __init__(self, spreadsheet_id: str, tab_name: str):
        self.spreadsheet_id = spreadsheet_id
        self.tab_name = tab_name
        self.executed = False
        self.sh_id: Optional[int] = None

    def render_json(self) -> dict:
        json = {"requests": [{"addSheet": {"properties": {"title": self.tab_name}}}]}
        return json

    def execute(self, service: Resource) -> None:
        output = (
            service.spreadsheets()
            .batchUpdate(spreadsheetId=self.spreadsheet_id, body=self.render_json())
            .execute()
        )
        self.sh_id = json_val_extract(output, "sheetId")
        self.executed = True

    @property
    def sheet_id(self) -> Optional[int]:
        if self.executed:
            return self.sh_id
        else:
            raise RuntimeError("Must run the execute method before passing the id")
