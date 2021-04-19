import pandas as pd
import pytest

from gslides.addchart import Area, Chart, Column, Histogram, Line, Scatter
from gslides.sheetsframe import GetFrame


def test_df():
    data = [
        ["Object", "Blue", "Red", "Grand Total"],
        ["Ball", "6", "1", "7"],
        ["Cube", "6", "4", "10"],
        ["Stick", "7", "5", "12"],
    ]
    return pd.DataFrame(columns=data[0], data=data[1:])


class MockService:
    def spreadsheets(self, **kwargs):
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


def test_series_inits():
    assert Line().__class__.__name__ == "Line"
    assert Column().__class__.__name__ == "Column"
    assert Area().__class__.__name__ == "Area"
    assert Scatter().__class__.__name__ == "Scatter"
    assert Histogram().__class__.__name__ == "Histogram"


class TestLine:
    def setup(self):
        self.object = Line(
            y_columns=["a", "b", "c"],
            line_style="MEDIUM_DASHED",
            line_width=1,
            point_enabled=True,
            point_shape="DIAMOND",
            point_size=10,
            data_label_enabled=True,
            data_label_placement="BELOW",
        )

    def test_repr(self):
        assert self.object.__repr__()[13:17] == "Line"

    def test_check_point_args(self):
        assert self.object._check_point_args(self.object.__dict__) == None

    def test_check_data_label_args(self):
        assert self.object._check_data_label_args(self.object.__dict__) == None

    def test_render_basic_chart_json(self):
        output = self.object.render_basic_chart_json(
            None, 1234, 1, 1, 2, 2, type="Line"
        )
        assert output["type"] == "Line"


class TestChart:
    def setup(self):
        l = Line(
            y_columns=["Blue"],
            line_style="MEDIUM_DASHED",
            line_width=1,
            point_enabled=True,
            point_shape="DIAMOND",
            point_size=10,
            data_label_enabled=True,
            data_label_placement="BELOW",
        )

        frame = GetFrame(
            spreadsheet_id="abc123",
            sheet_id=1234,
            anchor_cell="A1",
            bottom_right_cell="B4",
        )

        frame.df = test_df()

        self.object = Chart(
            frame,
            [l],
            x_column="Object",
            stacking=False,
            title="pytest",
            x_axis_label="X LABEL",
            y_axis_label="Y LABEL",
            x_min=0,
            x_max=100,
            y_min=-10,
            y_max=15,
            palette=None,
            legend_position="RIGHT_LEGEND",
        )

    @pytest.mark.xfail(reason=RuntimeError)
    def test_sheet_id(self):
        assert self.object.chart_id == None

    def test_determine_chart_type(self):
        assert self.object._determine_chart_type([Line()]) == "LINE"
        assert self.object._determine_chart_type([Line(), Column()]) == "COMBO"

    def test_check_stacking(self):
        assert self.object._check_stacking([Line()], False) == None
        assert self.object._check_stacking([Column(), Line()], True) == None

    def test_resolve_series(self):
        assert self.object._resolve_series()["Blue"].__class__.__name__ == "Line"

    def test_render_basic_chart_json(self):
        assert (
            self.object.render_basic_chart_json()["chart"]["spec"]["title"] == "pytest"
        )

    def test_execute(self, monkeypatch):
        def mock_return(self):
            return {"chartId": 11111}

        monkeypatch.setattr(MockService, "execute", mock_return)
        service = MockService()
        self.object.execute(service)
        assert self.object.ch_id == 11111
        assert self.object.executed
