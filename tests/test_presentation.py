import numpy as np
import pytest

from gslides.presentation import AddSlide, Layout, Presentation


class MockService:
    def presentations(self, **kwargs):
        return self

    def pages(self, **kwargs):
        return self

    def getThumbnail(self, **kwargs):
        return self

    def batchUpdate(self, **kwargs):
        return self

    def create(self, **kwargs):
        return self

    def get(self, **kwargs):
        return self

    def values(self, **kwargs):
        return self

    def execute(self, **kwargs):
        return self


class TestLayout:
    def setup(self):
        self.object = Layout(x_length=10, y_length=10, layout=(1, 2))

    def test_calc_size(self):
        assert self.object._calc_size() == (4.4, 9.8)

    def test_iter(self):
        assert self.object.__iter__() == self.object

    def test_coord(self):
        assert self.object.coord == (0, 0)
        next(self.object)
        assert self.object.coord == (0, 1)

    def test_next(self):
        assert self.object.__next__() == (0.5, 0.1)
        np.testing.assert_allclose(self.object.__next__(), (5.1, 0.1))


class MockFrame:
    def __init__(self, spreadsheet_id):
        self.spreadsheet_id = spreadsheet_id


class MockChart:
    def __init__(self, data, chart_id, size):
        self.data = data
        self.chart_id = chart_id
        self.size = size

    def create(self, *args, **kwargs):
        return None


class TestAddSlide:
    def setup(self):
        sh = MockFrame("zyxw")
        self.ch = MockChart(sh, 1234, (4.4, 9.8))
        self.object = AddSlide(
            presentation_id="abcd",
            objects=[self.ch, self.ch],
            layout=(1, 2),
            insertion_index=2,
        )

    @pytest.mark.parametrize(
        "input,expected",
        [
            pytest.param([1, 2], None, marks=pytest.mark.xfail(reason=ValueError)),
            pytest.param((0, 2), None, marks=pytest.mark.xfail(reason=ValueError)),
            pytest.param((1, 0), None, marks=pytest.mark.xfail(reason=ValueError)),
            ((1, 2), (1, 2)),
        ],
    )
    def test_validate_layout(self, input, expected):
        assert self.object._validate_layout(input) == expected

    def test_render_json_create_slide(self):
        assert (
            self.object.render_json_create_slide()["requests"][0]["createSlide"][
                "insertionIndex"
            ]
            == 2
        )

    def test_render_json_create_textboxes(self):
        assert (
            self.object.render_json_create_textboxes(1234)["requests"][0][
                "createShape"
            ]["shapeType"]
            == "TEXT_BOX"
        )

    def test_render_json_format_textboxes(self):
        assert (
            self.object.render_json_format_textboxes(56, 78)["requests"][0][
                "insertText"
            ]["objectId"]
            == 56
        )

    def test_render_json_copy_chart(self):
        assert (
            self.object.render_json_copy_chart(self.ch, (4.4, 9.8), 0.5, 0.1)[
                "createSheetsChart"
            ]["spreadsheetId"]
            == "zyxw"
        )

    def test_execute_create_slide(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.slide_service", property(mock_service)
        )

        def mock_return(self):
            return {"replies": [{"createSlide": {"objectId": 1111}}]}

        monkeypatch.setattr(MockService, "execute", mock_return)
        self.object._execute_create_slide()
        assert self.object.sl_id == 1111

    def test_execute_create_format_textboxes(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.slide_service", property(mock_service)
        )
        self.object.sl_id == 1111

        def mock_return(self):
            return {
                "replies": [
                    {"createShape": {"objectId": 2222}},
                    {"createShape": {"objectId": 3333}},
                ]
            }

        monkeypatch.setattr(MockService, "execute", mock_return)
        self.object._execute_create_format_textboxes()
        assert self.object.title_bx_id == 2222
        assert self.object.notes_bx_id == 3333

    def test_execute_populate_objects(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.slide_service", property(mock_service)
        )
        self.object._execute_populate_objects()
        assert True

    def test_execute_slide(self, monkeypatch):
        self.object.sheet_executed = True

        def mock_return():
            return None

        monkeypatch.setattr(self.object, "_execute_create_slide", mock_return)
        monkeypatch.setattr(
            self.object, "_execute_create_format_textboxes", mock_return
        )
        monkeypatch.setattr(self.object, "_execute_populate_objects", mock_return)
        self.object.execute_slide()
        assert self.object.slide_executed

    def test_execute_sheet(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.sheet_service", property(mock_service)
        )

        def mock_return(self):
            return None

        monkeypatch.setattr(MockService, "execute", mock_return)
        self.object.execute_sheet()
        assert self.object.sheet_executed

    def test_execute(self, monkeypatch):
        def mock_return():
            return None

        monkeypatch.setattr(self.object, "execute_sheet", mock_return)
        monkeypatch.setattr(self.object, "execute_slide", mock_return)
        self.object.execute()
        assert True


class TestPresentation:
    def setup(self):
        self.object = Presentation(
            name="test",
            pr_id="abcd",
            sl_ids=[1111, 2222, 3333],
            ch_ids={"a1b2c3d4": "Test Chart"},
            initialized=True,
        )

    def test_repr(self):
        self.object.__repr__()
        assert True

    def test_create(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.slide_service", property(mock_service)
        )

        def mock_return(self):
            return {"presentationId": "abcd"}

        monkeypatch.setattr(MockService, "execute", mock_return)
        assert Presentation.create(name="test").pr_id == "abcd"

    def test_get(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.slide_service", property(mock_service)
        )

        def mock_return(self):
            return {
                "slides": [{"objectId": 1111}, {"objectId": 2222}, {"objectId": 3333}],
                "title": "test",
                "pageSize": {
                    "width": {"magnitude": 9144000, "unit": "EMU"},
                    "height": {"magnitude": 5143500, "unit": "EMU"},
                },
            }

        monkeypatch.setattr(MockService, "execute", mock_return)
        assert Presentation.get(presentation_id="abcd").sl_ids == [1111, 2222, 3333]

    def test_add_slide(self, monkeypatch):
        def mock_return(self):
            return (4444, {"a1b2c3d4": "Test Chart"})

        monkeypatch.setattr(AddSlide, "execute", mock_return)
        self.object.add_slide(objects=[], layout=(1, 1))
        assert self.object.sl_ids[-1] == 4444

    def test_rm_slide(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.slide_service", property(mock_service)
        )

        def mock_return(self):
            return None

        monkeypatch.setattr(MockService, "execute", mock_return)
        self.object.rm_slide(slide_id=3333)
        assert self.object.sl_ids == [1111, 2222]

    def test_template(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.slide_service", property(mock_service)
        )

        self.object.template({"old": "new"})
        assert True

    def test_update_charts(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.slide_service", property(mock_service)
        )

        self.object.update_charts()
        assert True

    def test_presentation_id(self):
        assert self.object.presentation_id == "abcd"

    def test_slide_ids(self):
        assert self.object.slide_ids == [1111, 2222, 3333]

    def test_chart_ids(self):
        assert self.object.chart_ids == {"a1b2c3d4": "Test Chart"}
