# -*- coding: utf-8 -*-
"""
Creates the slides and charts in Google slides
"""

from typing import Any, Dict, List, Optional, Tuple, Type, TypeVar, Union

from . import creds, package_font
from .chart import Chart
from .table import Table
from .utils import optimize_size, validate_params_float

TLayout = TypeVar("TLayout", bound="Layout")
TPresentation = TypeVar("TPresentation", bound="Presentation")


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

        :param slide_id: The slide_id of the slide to create textbox in
        :type slide_id: str
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
                (
                    service.presentations()
                    .batchUpdate(presentationId=self.presentation_id, body=json)
                    .execute()
                )
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

    def execute(self) -> str:
        """Executes the sheets & slides API call."""
        self.execute_sheet()
        self.execute_slide()
        return self.sl_id


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
    :param initialized: Whether to object has been initialized
    :type initialized: bool
    """

    def __init__(
        self,
        name: str = "",
        pr_id: str = "",
        sl_ids: list = [],
        initialized: bool = False,
    ) -> None:
        """Constructor method"""
        self.name = name
        self.pr_id = pr_id
        self.sl_ids = sl_ids
        self.initialized = initialized

    @classmethod
    def create(cls: Type[TPresentation], name: str = "Untitled") -> TPresentation:
        """Class method that creates a new presentation.

        :param name: Name of the presentation
        :type name: str
        :return: A presentation object
        :rtype: :class:`Presentation`

        """
        service: Any = creds.slide_service
        output = service.presentations().create(body={"title": name}).execute()
        pr_id = output["presentationId"]
        service.presentations().batchUpdate(
            presentationId=output["presentationId"],
            body={"requests": [{"deleteObject": {"objectId": "p"}}]},
        ).execute()
        return cls(name, pr_id, [], True)

    @classmethod
    def get(cls: Type[TPresentation], presentation_id: str) -> TPresentation:
        """Class method that gets a presentation.

        :param presentation_id: Id of the presentation
        :type presentation_id: str
        :return: A presentation object
        :rtype: :class:`Presentation`

        """
        service: Any = creds.slide_service
        output = service.presentations().get(presentationId=presentation_id).execute()
        name = output["title"]
        if "slides" in output.keys():
            sl_ids = [sl["objectId"] for sl in output["slides"]]
        else:
            sl_ids = []
        return cls(name, presentation_id, sl_ids, True)

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
            title,
            notes,
        )
        output = sl.execute()
        if insertion_index is None:
            self.sl_ids.append(output)
        else:
            self.sl_ids.insert(insertion_index, output)

    def rm_slide(self, slide_id: str) -> None:
        """Removes a slide based on a slide id.

        :param slide_id: The slide_id of the slide to delete
        :type slide_id: str
        """
        service: Any = creds.slide_service
        service.presentations().batchUpdate(
            presentationId=self.presentation_id,
            body={"requests": [{"deleteObject": {"objectId": slide_id}}]},
        ).execute()
        self.sl_ids.remove(slide_id)

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
        :return: The presentation_id of the created presentation
        :rtype: str
        """
        if self.initialized:
            return self.sl_ids
        else:
            raise RuntimeError(
                "Must run the create or get method before passing the slide ids"
            )
