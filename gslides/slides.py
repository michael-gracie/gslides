from .utils import optimize_size, validate_params_float


class CreatePresentation:
    def __init__(
        self,
        name="Untitled",
    ):
        self.name = name
        self.executed = False
        self.pr_id = None

    def execute(self, service):
        output = service.presentations().create(body={"title": self.name}).execute()
        self.pr_id = output["presentationId"]
        service.presentations().batchUpdate(
            presentationId=output["presentationId"],
            body={"requests": [{"deleteObject": {"objectId": "p"}}]},
        ).execute()
        self.executed = True

    @property
    def presentation_id(self):
        if self.executed:
            return self.pr_id
        else:
            raise RuntimeError(
                "Must run the execute method before passing the presentation id"
            )


class Layout:
    def __init__(
        self, x_length, y_length, layout, x_border=0.05, y_border=0.01, spacing=0.02
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

    def _calc_size(self):
        x_size = (
            self.x_length
            - self.x_length * ((self.y_objects - 1) * self.spacing + self.x_border * 2)
        ) / self.y_objects
        y_size = (
            self.y_length
            - self.y_length * ((self.x_objects - 1) * self.spacing + self.y_border * 2)
        ) / self.x_objects
        return (x_size, y_size)

    def __iter__(self):
        return self

    @property
    def coord(self):
        x_coord = self.index // self.y_objects
        y_coord = self.index % self.y_objects
        return (x_coord, y_coord)

    def __next__(self):
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
        presentation_id,
        charts=None,
        layout=None,
        insertion_index=None,
        top_margin=1017724,
        bottom_margin=420575,
        left_margin=0,
        right_margin=0,
    ):
        self.presentation_id = presentation_id
        self.charts = charts
        self.layout = self._validate_layout(layout)
        self.insertion_index = insertion_index
        self.sl_id = None
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

    def _validate_layout(self, layout):
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

    def render_json_create_slide(self):
        json = {
            "requests": [
                {"createSlide": {}},
            ]
        }
        if self.insertion_index:
            json["requests"][0]["createSlide"]["insertionIndex"] = self.insertion_index
        return json

    def render_json_create_textboxes(self, slide_id):
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

    def render_json_format_textboxes(self, title_box_id, notes_box_id):
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

    def render_json_copy_chart(self, chart, size, translate_x, translate_y):
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

    def execute_slide(self, service):
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
        json = {"requests": []}
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

    def execute_sheet(self, service):
        x_len, y_len = optimize_size(
            self.layout_obj.object_size[1] / self.layout_obj.object_size[0],
            area=222600 / (self.layout[0] * self.layout[1]),
        )
        for ch in self.charts:
            ch.size = (int(x_len), int(y_len))
            ch.execute(service)
        self.sheet_executed = True
