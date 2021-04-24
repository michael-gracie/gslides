import numpy as np
import pytest

from gslides.slides import CreatePresentation, CreateSlide, Layout


class MockService:
    def presentations(self, **kwargs):
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


class TestCreatePresentation:
    def setup(self):
        self.object = CreatePresentation(name="pytest")

    def test_execute(self, monkeypatch):
        def mock_return(self):
            return {"presentationId": "abcd"}

        monkeypatch.setattr(MockService, "execute", mock_return)
        service = MockService()
        self.object.execute(service)
        assert self.object.executed == True

    @pytest.mark.xfail(reason=RuntimeError)
    def test_presentation_id(self):
        assert self.object.presentation_id == None


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


class MockSheetsFrame:
    def __init__(self, spreadsheet_id):
        self.spreadsheet_id = spreadsheet_id


class MockChart:
    def __init__(self, data, chart_id, size):
        self.data = data
        self.chart_id = chart_id
        self.size = size

    def execute(self, *args, **kwargs):
        return None


class TestCreateSlide:
    def setup(self):
        sh = MockSheetsFrame("zyxw")
        self.ch = MockChart(sh, 1234, (4.4, 9.8))
        self.object = CreateSlide(
            presentation_id="abcd",
            charts=[self.ch, self.ch],
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
        self.object.sheet_executed = True

        def mock_return(self):
            return {"replies": [{"createSlide": {"objectId": 1111}}]}

        monkeypatch.setattr(MockService, "execute", mock_return)
        service = MockService()
        self.object._execute_create_slide(service)
        assert self.object.sl_id == 1111

    def test_execute_create_slide(self, monkeypatch):
        self.object.sheet_executed = True

        def mock_return(self):
            return {"replies": [{"createSlide": {"objectId": 1111}}]}

        monkeypatch.setattr(MockService, "execute", mock_return)
        service = MockService()
        self.object._execute_create_slide(service)
        assert self.object.sl_id == 1111

    def test_execute_create_format_textboxes(self, monkeypatch):
        self.object.sl_id == 1111

        def mock_return(self):
            return {
                "replies": [
                    {"createShape": {"objectId": 2222}},
                    {"createShape": {"objectId": 3333}},
                ]
            }

        monkeypatch.setattr(MockService, "execute", mock_return)
        service = MockService()
        self.object._execute_create_format_textboxes(service)
        assert self.object.title_bx_id == 2222
        assert self.object.notes_bx_id == 3333

    def test_execute_slide(self, monkeypatch):
        def mock_return(self):
            return None

        monkeypatch.setattr(self.object, "_execute_create_slide", mock_return)
        monkeypatch.setattr(
            self.object, "_execute_create_format_textboxes", mock_return
        )
        service = MockService()
        self.object.execute_slide(service)
        assert self.object.slide_executed

    def test_execute_sheet(self, monkeypatch):
        def mock_return(self):
            return None

        monkeypatch.setattr(MockService, "execute", mock_return)
        service = MockService()
        self.object.execute_sheet(service)
        assert self.object.sheet_executed
