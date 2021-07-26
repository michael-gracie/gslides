# -*- coding: utf-8 -*-
"""
Creates the slides and charts in Google slides
"""
import logging
import pprint
from typing import Any, Dict, List, Optional, Tuple, Type, TypeVar, Union

import requests
from IPython.display import Image

from . import creds, package_font
from .chart import Chart
from .config import PRESENTATION_PARAMS
from .table import Table
from .utils import json_chunk_key_extract, optimize_size, validate_params_float

TLayout = TypeVar("TLayout", bound="Layout")
TPresentation = TypeVar("TPresentation", bound="Presentation")

logger = logging.getLogger(__name__)


class Layout:
    """A class that manages the layout of objects on a canvas

    :param float x_length: The width of the canvase
    :type x_length: float
    :param float y_length: The height of the canvase
    :type y_length: float
    :param layout: The grid layout of objects on the canvase (rows x columns)
    :type layout: tuple
    :param float x_border: The % margin for the x border
    :type x_border: float
    :param float y_border: The % margin for the y border
    :type y_border: float
    :param float spacing: The % spacing between objects
    :type spacing: float

    """

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


class AddSlide:
    """The class that adds a slide to the presentation.

    :param presentation_id: The presentation_id of the created presentation
    :type presentation_id: str
    :param objects: List of :class:`Chart` or :class:`Table` objects
    :type objects: list
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
    :param page_size: Tuple of the width and height of the presentation in EMU
    :type page_size: tuple
    :param title: The text for the title textbox
    :type title: str, optional
    :param notes: The text for the notes textbox
    :type notes: str, optional
    """

    def __init__(
        self,
        presentation_id: str,
        objects: List[Union[Chart, Table]],
        layout: Tuple[int, int],
        insertion_index: Optional[int] = None,
        top_margin: int = 1017724,
        bottom_margin: int = 420575,
        left_margin: int = 0,
        right_margin: int = 0,
        page_size: Tuple[int, int] = (9144000, 5143500),
        title: str = "Title placeholder",
        notes: str = "Notes placeholder",
    ) -> None:
        """Constructor method"""
        self.presentation_id = presentation_id
        self.objects = self._validate_objects(objects)
        self.layout = self._validate_layout(layout)
        self.insertion_index = insertion_index
        self.sl_id: str = ""
        self.ch_ids: dict = {}
        self.sheet_executed = False
        self.slide_executed = False
        self.top_margin = top_margin
        self.bottom_margin = bottom_margin
        self.left_margin = left_margin
        self.right_margin = right_margin
        self.page_size = page_size
        self.layout_obj = Layout(
            page_size[0] - self.left_margin - self.right_margin,
            page_size[1] - self.top_margin - self.bottom_margin,
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

    def _validate_objects(
        self, objects: List[Union[Chart, Table]]
    ) -> List[Union[Chart, Table]]:
        """Validates that there is a list of objects

        :param objects: List of :class:`Chart` or :class:`Table` objects
        :type objects: list
        :raises ValueError:
        :return: The objects
        :rtype: list

        """
        if type(objects) != list:
            raise ValueError("Objects only accepts a list")
        else:
            return objects

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

        :param slide_id: The slide_id of the slide to create textbox in
        :type slide_id: str
        :return: The json to do the update
        :rtype: dict
        """
        start_textbox_x = self.page_size[0] * 0.05
        scale_textbox_x = self.page_size[0] * 0.9 / 3000000
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
                                "scaleX": scale_textbox_x,
                                "scaleY": 0.1909,
                                "translateX": start_textbox_x,
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
                                "scaleX": scale_textbox_x,
                                "scaleY": 0.0914,
                                "translateX": start_textbox_x,
                                "translateY": self.page_size[1] - self.bottom_margin,
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

        :param title_box_id: The id of the title box
        :type title_box_id: int
        :param notes_box_id: The id of the notes box
        :type notes_box_id: int
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
                            "fontFamily": package_font.font,
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
                            "fontFamily": package_font.font,
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

        :param chart: The chart to copy
        :type chart: :class:`Chart`
        :param size: Tuple of width and height in EMU
        :type size: tuple
        :param translate_x: The number of EMU to translate the object by
        :type translate_x: float
        :param translate_y: The number of EMU to translate the object by
        :type translate_y: float
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
        logger.info("Creating slide")
        output = (
            service.presentations()
            .batchUpdate(
                presentationId=self.presentation_id,
                body=self.render_json_create_slide(),
            )
            .execute()
        )
        logger.info("Slide created successfully")
        self.sl_id = output["replies"][0]["createSlide"]["objectId"]

    def _execute_create_format_textboxes(self) -> None:
        """Executes the create & format textboxes slides API call."""
        service: Any = creds.slide_service
        body = self.render_json_create_textboxes(self.sl_id)
        logger.info("Executing textbox creation")
        logger.info(f"Request: {pprint.pformat(body)}")
        output = (
            service.presentations()
            .batchUpdate(
                presentationId=self.presentation_id,
                body=self.render_json_create_textboxes(self.sl_id),
            )
            .execute()
        )
        logger.info("Textboxes created successfully")
        self.title_bx_id = output["replies"][0]["createShape"]["objectId"]
        self.notes_bx_id = output["replies"][1]["createShape"]["objectId"]
        body = self.render_json_format_textboxes(self.title_bx_id, self.notes_bx_id)
        logger.info("Executing textbox creation")
        logger.info(f"Request: {pprint.pformat(body)}")
        output = (
            service.presentations()
            .batchUpdate(
                presentationId=self.presentation_id,
                body=body,
            )
            .execute()
        )
        logger.info("Textboxes formatted successfully")

    def _execute_populate_objects(self) -> None:
        """Executes the population of objects on the slide"""
        json: Dict[str, Any] = {"requests": []}
        service = creds.slide_service
        for obj in self.objects:
            translate_x, translate_y = next(self.layout_obj)
            if isinstance(obj, Chart):
                json["requests"].append(
                    self.render_json_copy_chart(
                        obj, self.layout_obj.object_size, translate_x, translate_y
                    )
                )
                logger.info("Populating charts in google slides")
                logger.info(f"Request: {pprint.pformat(json)}")
                output = (
                    service.presentations()
                    .batchUpdate(presentationId=self.presentation_id, body=json)
                    .execute()
                )
                self.ch_ids[
                    output["replies"][0]["createSheetsChart"]["objectId"]
                ] = obj.title
                logger.info("Charts successfully populated")
            elif isinstance(obj, Table):
                obj.create(
                    self.presentation_id,
                    self.sl_id,
                    self.layout_obj.object_size,
                    self.left_margin + translate_x,
                    self.top_margin + translate_y,
                )

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
            if isinstance(obj, Chart):
                obj.create((int(x_len), int(y_len)))
        self.sheet_executed = True

    def execute(self) -> Tuple[str, Dict[Any, Any]]:
        """Executes the sheets & slides API call."""
        self.execute_sheet()
        self.execute_slide()
        return self.sl_id, self.ch_ids


class Presentation:
    """An object that represents a presentation in Google slides. Initialize the
    object through either the :class:`Presentation.get` or
    :class:`Presentation.create` class method.

    :param name: Name of the presentation
    :type name: str
    :param pr_id: The id of the presentation
    :type pr_id: str
    :param sl_ids: A list of the slide ids
    :type sl_ids: list
    :param ch_ids: A dictionary of the charts objects id's and their corresponding
        title
    :type ch_ids: dict
    :param page_size: Tuple of the width and height of the presentation in EMU
    :type page_size: tuple
    :param initialized: Whether to object has been initialized
    :type initialized: bool
    """

    def __init__(
        self,
        name: str = "",
        pr_id: str = "",
        sl_ids: list = [],
        ch_ids: dict = {},
        page_size: Tuple[int, int] = (9144000, 5143500),
        initialized: bool = False,
    ) -> None:
        """Constructor method"""
        self.name = name
        self.pr_id = pr_id
        self.sl_ids = sl_ids
        self.ch_ids = ch_ids
        self.page_size = page_size
        self.initialized = initialized

    def __repr__(self) -> str:
        """Prints class information.

        :return: String with helpful class infromation
        :rtype: str

        """
        output = f"Presentation\n" f" - presentation_id = {self.presentation_id}"
        return output

    @classmethod
    def create(
        cls: Type[TPresentation],
        name: str = "Untitled",
    ) -> TPresentation:
        """Class method that creates a new presentation. To note, due to an issue
        in the API page size is currently not supported.

        :param name: Name of the presentation
        :type name: str
        :return: A presentation object
        :rtype: :class:`Presentation`

        """
        service: Any = creds.slide_service
        logger.info("Creating presentation")
        output = service.presentations().create(body={"title": name}).execute()
        pr_id = output["presentationId"]
        service.presentations().batchUpdate(
            presentationId=output["presentationId"],
            body={"requests": [{"deleteObject": {"objectId": "p"}}]},
        ).execute()
        logger.info("Presentation successfully created")
        return cls(name, pr_id, [], {}, (9144000, 5143500), True)

    @classmethod
    def get(cls: Type[TPresentation], presentation_id: str) -> TPresentation:
        """Class method that gets a presentation.

        :param presentation_id: Id of the presentation
        :type presentation_id: str
        :return: A presentation object
        :rtype: :class:`Presentation`

        """
        service: Any = creds.slide_service
        logger.info("Retrieving presentation")
        output = service.presentations().get(presentationId=presentation_id).execute()
        logger.info("Presentation successfully retrieved")
        name = output["title"]
        page_size = (
            output["pageSize"]["width"]["magnitude"],
            output["pageSize"]["height"]["magnitude"],
        )
        if "slides" in output.keys():
            sl_ids = [sl["objectId"] for sl in output["slides"]]
        else:
            sl_ids = []
        chart_ids = {}
        charts = json_chunk_key_extract(output, "sheetsChart")
        for chart in charts:
            chart_ids[chart["objectId"]] = chart["title"]
        return cls(name, presentation_id, sl_ids, chart_ids, page_size, True)

    def add_slide(
        self,
        objects: List[Union[Chart, Table]],
        layout: Tuple[int, int],
        insertion_index: Optional[int] = None,
        top_margin: int = 1017724,
        bottom_margin: int = 420575,
        left_margin: int = 0,
        right_margin: int = 0,
        title: str = "Title placeholder",
        notes: str = "Notes placeholder",
    ) -> None:
        """Add a slide to the presentation.

        :param objects: List of :class:`Chart` or :class:`Table` objects
        :type objects: list
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
        :param title: The text for the title textbox
        :type title: str, optional
        :param notes: The text for the notes textbox
        :type notes: str, optional
        """
        sl = AddSlide(
            self.presentation_id,
            objects,
            layout,
            insertion_index,
            top_margin,
            bottom_margin,
            left_margin,
            right_margin,
            self.page_size,
            title,
            notes,
        )
        new_sl_id, new_ch_ids = sl.execute()
        if insertion_index is None:
            self.sl_ids.append(new_sl_id)
        else:
            self.sl_ids.insert(insertion_index, new_sl_id)
        self.ch_ids = {**self.ch_ids, **new_ch_ids}

    def rm_slide(self, slide_id: str) -> None:
        """Removes a slide based on a slide id.

        :param slide_id: The slide_id of the slide to delete
        :type slide_id: str
        """
        service: Any = creds.slide_service
        logger.info("Deleting slide")
        service.presentations().batchUpdate(
            presentationId=self.presentation_id,
            body={"requests": [{"deleteObject": {"objectId": slide_id}}]},
        ).execute()
        logger.info("Slide successfully deleted")
        self.sl_ids.remove(slide_id)

    def template(self, mapping: dict, slide_ids: list = []) -> None:
        """Replaces all text encaspulated with `{{ <TEXT> }}` with input.

        :param mapping: Dictionary mapping old text to new text
        :type mapping: dict
        :param slide_ids: The slides to apply template on. If none, then all slides
            will be considered.
        :type slide_ids: list, optional
        """

        requests = []
        for key, val in mapping.items():
            json = {
                "replaceAllText": {
                    "replaceText": val,
                    "pageObjectIds": slide_ids,
                    "containsText": {"text": f"{{{{ {key} }}}}", "matchCase": False},
                }
            }
            requests.append(json)
        service: Any = creds.slide_service
        logger.info("Templating data")
        service.presentations().batchUpdate(
            presentationId=self.presentation_id,
            body={"requests": requests},
        ).execute()
        logger.info("Data successfully templated")

    def update_charts(self) -> None:
        """Updates all the charts in the slides deck with refreshed underlying
        data.
        """
        requests = []
        for key in self.chart_ids.keys():
            json = {
                "refreshSheetsChart": {
                    "objectId": key,
                }
            }
            requests.append(json)
        service: Any = creds.slide_service
        logger.info("Update charts")
        service.presentations().batchUpdate(
            presentationId=self.presentation_id,
            body={"requests": requests},
        ).execute()
        logger.info("Charts successfully updated")

    def _validate_image_size(self, image_size):
        """Validate that the image size configuration is valid

        :param image_size: String to configure the image size
        :type image_size: str
        :raises ValueError:
        """
        if image_size not in PRESENTATION_PARAMS["data_label_placement"]["params"]:
            raise ValueError(
                f"{image_size} is not a valid parameter for image_size. "
                f"Accepted parameters include"
                f"{', '.join(PRESENTATION_PARAMS['data_label_placement']['params'])}. See"
                f"{PRESENTATION_PARAMS['data_label_placement']['url']} for further documentation."
            )

    def show_slide(self, slide_id: str, image_size: str = "LARGE") -> Image:
        """Displays a given slide in a Jupyter notebook.

        :param slide_id: The id of the slide to show
        :type slide_id: str
        :param image_size: String to configure the image size
        :type image_size: str
        :return: ipython Image object
        :rtype: Image

        """
        self._validate_image_size(image_size)
        service: Any = creds.slide_service
        img_info = (
            service.presentations()
            .pages()
            .getThumbnail(
                presentationId=self.presentation_id,
                pageObjectId=slide_id,
                thumbnailProperties_thumbnailSize=image_size,
            )
            .execute()
        )
        return Image(requests.get(img_info["contentUrl"]).content)

    def download_slide(
        self, slide_id: str, path: str, image_size: str = "LARGE"
    ) -> None:
        """Downloads a given slide to a file in png format

        :param slide_id: The id of the slide to show
        :type slide_id: str
        :param path: Path to write png file to
        :type path: str
        :param image_size: String to configure the image size
        :type image_size: str
        """
        self._validate_image_size(image_size)
        service: Any = creds.slide_service
        img_info = (
            service.presentations()
            .pages()
            .getThumbnail(
                presentationId=self.presentation_id,
                pageObjectId=slide_id,
                thumbnailProperties_thumbnailSize=image_size,
            )
            .execute()
        )
        with open(path, "wb") as f:
            f.write(requests.get(img_info["contentUrl"]).content)
        return None

    @property
    def get_method(self) -> str:
        """Returns the corresponding get initialization method.

        :return: Get intialization method
        :rtype: str

        """
        return f"=Presentation.get(presentation_id='{self.presentation_id}')"

    @property
    def presentation_id(self) -> str:
        """Returns the presentation_id of the created presentation.

        :raises RuntimeError: Must run the create or get method before passing the presentation id
        :return: The presentation_id of the created presentation
        :rtype: str
        """
        if self.initialized:
            return self.pr_id
        else:
            raise RuntimeError(
                "Must run the create or get method before passing the presentation id"
            )

    @property
    def slide_ids(self) -> List[str]:
        """Returns the slide_ids of the created presentation.

        :raises RuntimeError: Must run the create or get method before passing the slides ids
        :return: The slide ids of the created presentation
        :rtype: str
        """
        if self.initialized:
            return self.sl_ids
        else:
            raise RuntimeError(
                "Must run the create or get method before passing the slide ids"
            )

    @property
    def chart_ids(self) -> dict:
        """Returns the chart_ids of the created presentation.

        :raises RuntimeError: Must run the create or get method before passing the slides ids
        :return: The chart ids of the created presentation
        :rtype: dict
        """
        if self.initialized:
            return self.ch_ids
        else:
            raise RuntimeError(
                "Must run the create or get method before passing the slide ids"
            )
