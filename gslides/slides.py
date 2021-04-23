# -*- coding: utf-8 -*-
from typing import Any, Dict, List, Optional, Tuple, TypeVar

from googleapiclient.discovery import Resource

from .addchart import Chart
from .utils import optimize_size, validate_params_float


class CreatePresentation:
    def __init__(
        self,
        name: str = "Untitled",
    ) -> None:
        self.name = name
        self.executed = False
        self.pr_id: Optional[str] = None

    def execute(self, service: Resource) -> None:
        output = service.presentations().create(body={"title": self.name}).execute()
        self.pr_id = output["presentationId"]
        service.presentations().batchUpdate(
            presentationId=output["presentationId"],
            body={"requests": [{"deleteObject": {"objectId": "p"}}]},
        ).execute()
        self.executed = True

    @property
    def presentation_id(self) -> Optional[str]:
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
        return self

    @property
    def coord(self) -> Tuple[int, int]:
        x_coord = self.index // self.y_objects
        y_coord = self.index % self.y_objects
        return (x_coord, y_coord)

    def __next__(self) -> Tuple[float, float]:
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


class CreateSlide:
    def __init__(
        self,
        presentation_id: str,
        charts: List[Chart],
        layout: Tuple[int, int],
        insertion_index: Optional[int] = None,
        top_margin: int = 1017724,
        bottom_margin: int = 420575,
        left_margin: int = 0,
        right_margin: int = 0,
    ) -> None:
        self.presentation_id = presentation_id
        self.charts = charts
        self.layout = self._validate_layout(layout)
        self.insertion_index = insertion_index
        self.sl_id: Optional[int] = None
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

    def _validate_layout(self, layout: Tuple[int, int]) -> Tuple[int, int]:
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

    def render_json_create_slide(self) -> Dict[str, Any]:
        json: Dict[str, Any] = {
            "requests": [
                {"createSlide": {}},
            ]
        }
        if self.insertion_index:
            json["requests"][0]["createSlide"]["insertionIndex"] = self.insertion_index
        return json

    def render_json_create_textboxes(self, slide_id: Optional[int]) -> dict:
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
        json = {
            "requests": [
                {
                    "insertText": {
                        "objectId": title_box_id,
                        "insertionIndex": 0,
                        "text": "Title Placeholder",
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
                        "text": "Notes Placeholder",
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

    def execute_slide(self, service: Resource) -> None:
        if self.sheet_executed is False:
            raise RuntimeError(
                "Must run the execute sheet method before running the execute slide method"
            )
        output = (
            service.presentations()
            .batchUpdate(
                presentationId=self.presentation_id,
                body=self.render_json_create_slide(),
            )
            .execute()
        )
        self.sl_id = output["replies"][0]["createSlide"]["objectId"]
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
        json: Dict[str, Any] = {"requests": []}
        for ch in self.charts:
            translate_x, translate_y = next(self.layout_obj)
            json["requests"].append(
                self.render_json_copy_chart(
                    ch, self.layout_obj.object_size, translate_x, translate_y
                )
            )
        output = (
            service.presentations()
            .batchUpdate(presentationId=self.presentation_id, body=json)
            .execute()
        )
        self.slide_executed = True

    def execute_sheet(self, service: Resource) -> None:
        x_len, y_len = optimize_size(
            self.layout_obj.object_size[1] / self.layout_obj.object_size[0],
            area=222600 / (self.layout[0] * self.layout[1]),
        )
        for ch in self.charts:
            ch.size = (int(x_len), int(y_len))
            ch.execute(service)
        self.sheet_executed = True
