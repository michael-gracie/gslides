# -*- coding: utf-8 -*-
"""
Spreadsheet class
"""

import logging
import pprint
from typing import Any, Dict, List, Tuple, Type, TypeVar, cast

from . import creds
from .utils import json_dict_extract, json_val_extract

TSpreadsheet = TypeVar("TSpreadsheet", bound="Spreadsheet")

logger = logging.getLogger(__name__)


class CreateSpreadsheet:
    """An object to create a spreadsheet in Google sheets"""

    def render_json(self, title: str, sheet_names: List[str]) -> dict:
        """Renders the json to create the spreadsheet in Google sheets

        :param title: The title of the spreadsheet
        :type title: str
        :param sheet_names: The list of sheet names
        :type sheet_names: list
        :return: The json to do the update
        :rtype: dict
        """
        sheets = []
        for sheet in sheet_names:
            sheets.append({"properties": {"title": sheet}})
        json = {
            "properties": {
                "title": title,
                "locale": "en_US",
                "autoRecalc": "HOUR",
            },
            "sheets": sheets,
        }
        return json

    def execute(self, title: str, sheet_names: List[str]) -> Tuple[Any, ...]:
        """Executes the API call

        :param title: The title of the spreadsheet
        :type title: str
        :param sheet_names: The list of sheet names
        :type sheet_names: list
        :return: The json returned by the call
        :rtype: dict

        """
        service: Any = creds.sheet_service
        body = self.render_json(title, sheet_names)
        logger.info("Creating the spreadsheet")
        logger.info(f"Request: {pprint.pformat(body)}")
        output = service.spreadsheets().create(body=body).execute()
        logger.info("Spreadsheet created successfully")
        sp_id = cast(str, json_val_extract(output, "spreadsheetId")[0])
        sht_ids = cast(List[int], json_val_extract(output, "sheetId"))
        return (sp_id, sht_ids, True)


class GetSpreadsheet:
    """An object to get a spreadsheet in Google sheets"""

    def execute(self, spreadsheet_id: str) -> Tuple[Any, ...]:
        """Executes the API call

        :param spreadsheet_id: The id of the spreadsheet
        :type spreadsheet_id: str
        :return: The json returned by the call
        :rtype: dict

        """
        service: Any = creds.sheet_service
        logger.info("Retreiving spreadsheet")
        output = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        logger.info("Spreadsheet successfully retreived")
        title = output["properties"]["title"]
        sht_ids = cast(Dict[str, int], json_dict_extract(output, ("title", "sheetId")))
        return (title, sht_ids, True)


class AddSheet:
    """An object that adds a sheet in Google sheets"""

    def render_json(self, sheet_names: List[str]) -> dict:
        """Renders the json to create a list of sheets in Google sheets

        :param sheet_names: The list of sheet names
        :type sheet_names: list
        :return: The json to do the update
        :rtype: dict
        """
        json: Dict[str, Any] = {"requests": []}
        for sheet in sheet_names:
            json["requests"].append({"addSheet": {"properties": {"title": sheet}}})
        return json

    def execute(self, spreadsheet_id: str, sheet_names: List[str]) -> Dict[str, int]:
        """Executes the API call

        :param spreadsheet_id: The id of the spreadsheet
        :type spreadsheet_id: str
        :param sheet_names: The list of sheet names
        :type sheet_names: list
        :return: A dictionary of the sheet names and ids
        :rtype: dict

        """
        service: Any = creds.sheet_service
        body = self.render_json(sheet_names)
        logger.info("Executing sheet creation")
        logger.info(f"Request: {pprint.pformat(body)}")
        output = (
            service.spreadsheets()
            .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
            .execute()
        )
        logger.info("Sheet created successfully")
        sht = cast(List[int], json_val_extract(output, "sheetId"))
        sht_ids = dict(zip(sheet_names, sht))
        return sht_ids


class RemoveSheet:
    """An object that deletes a sheets from Google sheets"""

    def render_json(self, sheet_ids: List[int]) -> dict:
        """Renders the json to remove a list of sheets in Google sheets

        :param sheet_ids: The list of sheet ids
        :type sheet_ids: list
        :return: The json to do the update
        :rtype: dict
        """
        json: Dict[str, Any] = {"requests": []}
        for sheet_id in sheet_ids:
            json["requests"].append({"deleteSheet": {"sheetId": sheet_id}})
        return json

    def execute(self, spreadsheet_id: str, sheet_ids: List[int]) -> List[int]:
        """Executes the API call

        :param spreadsheet_id: The id of the spreadsheet
        :type spreadsheet_id: str
        :param sheet_ids: The list of sheet ids
        :type sheet_ids: list
        :return: List of deleted ids
        :rtype: list

        """
        service: Any = creds.sheet_service
        body = self.render_json(sheet_ids)
        logger.info("Deleting sheet")
        logger.info(f"Request: {pprint.pformat(body)}")
        (
            service.spreadsheets()
            .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
            .execute()
        )
        logger.info("Sheet deleted successfully")
        return sheet_ids


class Spreadsheet:
    """An object that represents a spreadsheet in Google sheets. Initialize the
    object through either the :class:`Spreadsheet.get` or :class:`Spreadsheet.create`
    class method.

    :param sp_id: The id of the spreadsheet
    :type sp_id: str, optional
    :param title: The title of the spreadsheet
    :type title: str, optional
    :param sheet_names: The list of sheet names
    :type sheet_names: list, optional
    :param initialized: A boolean whether the class has been initialized
    :type initialized: bool, optional
    """

    def __init__(
        self,
        sp_id: str = "",
        title: str = "",
        sht_ids: dict = {},
        initialized: bool = False,
    ) -> None:
        """Constructor method"""
        self.sp_id = sp_id
        self.title = title
        self.sht_nms = sht_ids
        self.initialized = initialized

    def __repr__(self) -> str:
        """Prints class information.

        :return: String with helpful class infromation
        :rtype: str

        """
        output = (
            f"Spreadsheet\n"
            f" - spreadsheet_id = {self.spreadsheet_id}\n"
            f" - title = {self.title}"
        )
        return output

    @classmethod
    def create(
        cls: Type[TSpreadsheet],
        title: str = "Untitled",
        sheet_names: List[str] = ["Sheet1"],
    ) -> TSpreadsheet:
        """Creates a new spreadsheet.

        :param title: The title of the spreadsheet
        :type title: str, optional
        :param sheet_names: The list of sheet names
        :type sheet_names: list, optional
        :return: A Spreadsheet object
        :rtype: :class:`gslides.Spreadsheet`

        """
        sp_id, sht_ids, initialized = CreateSpreadsheet().execute(title, sheet_names)
        sht_ids = dict(zip(sheet_names, sht_ids))
        return cls(sp_id, title, sht_ids, initialized)

    @classmethod
    def get(cls: Type[TSpreadsheet], spreadsheet_id: str) -> TSpreadsheet:
        """Gets an existing spreadsheet.

        :param spreadsheet_id: The id of the spreadsheet
        :type spreadsheet_id: str, optional
        :return: A Spreadsheet object
        :rtype: :class:`gslides.Spreadsheet`

        """
        title, sht_ids, initialized = GetSpreadsheet().execute(spreadsheet_id)
        return cls(spreadsheet_id, title, sht_ids, initialized)

    def add_sheets(self, sheet_names: List[str]) -> None:
        """Adds sheets to a spreadsheet

        :param sheet_names: The list of sheet names
        :type sheet_names: list
        :return: A Spreadsheet object
        :rtype: gslides.Spreadsheet

        """
        new_sht_ids = AddSheet().execute(self.spreadsheet_id, sheet_names)
        self.sht_nms.update(new_sht_ids)

    def rm_sheets(self, sheet_names: List[str]) -> None:
        """Removes sheets from a spreadsheet

        :param sheet_names: The list of sheet names
        :type sheet_names: list
        :return: A Spreadsheet object
        :rtype: gslides.Spreadsheet

        """
        sht_ids = [self.sht_nms[id] for id in sheet_names]
        RemoveSheet().execute(self.spreadsheet_id, sht_ids)
        for nm in sheet_names:
            self.sht_nms.pop(nm)

    @property
    def get_method(self) -> str:
        """Returns the corresponding get initialization method.

        :return: Get initialization method
        :rtype: str

        """
        return f"=Spreadsheet.get(spreadsheet_id='{self.spreadsheet_id}')"

    @property
    def spreadsheet_id(self) -> str:
        """Returns the spreadsheet_id of the created spreadsheet.

        :raises RuntimeError: Must run the create or get method before obtaining the id
        :return: The spreadsheet_id of the created spreadsheet.
        :rtype: str
        """
        if self.initialized:
            return self.sp_id
        else:
            raise RuntimeError(
                "Must run the create or get method before obtaining the id"
            )

    @property
    def sheet_names(self) -> Dict[int, Any]:
        """Returns the sheets names and ids for the spreadsheet

        :raises RuntimeError: Must run the create or get method before obtaining the id
        :return: A dictionary of the sheet_ids and names for a spreadsheet
        :rtype: dict
        """
        if self.initialized:
            return self.sht_nms
        else:
            raise RuntimeError(
                "Must run the create or get method before obtaining the id"
            )
