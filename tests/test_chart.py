import pandas as pd
import pytest

from gslides.chart import Chart, Series
from gslides.frame import Frame


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
    assert Series.line().type == "Line"
    assert Series.column().type == "Column"
    assert Series.area().type == "Area"
    assert Series.scatter().type == "Scatter"
    assert Series.histogram().type == "Histogram"


class TestLine:
    def setup(self):
        self.object = Series.line(
            series_columns=["a", "b", "c"],
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


class TestHistogram:
    def setup(self):
        self.object = Series.histogram(
            series_columns=["a", "b", "c"],
            bucket_size=10,
            outlier_percentage=0.5,
        )

    def test_render_histogram_chart_json(self):
        output = self.object.render_histogram_chart_json(None, 1234, 1, 1, 2, 2)
        assert output["data"]["sourceRange"]["sources"][0]["sheetId"] == 1234


class TestChart:
    def setup(self):
        l = Series.line(
            series_columns=["Blue"],
            line_style="MEDIUM_DASHED",
            line_width=1,
            point_enabled=True,
            point_shape="DIAMOND",
            point_size=10,
            data_label_enabled=True,
            data_label_placement="BELOW",
        )

        frame = Frame(
            df=test_df(),
            spreadsheet_id="abc123",
            sheet_id=1234,
            sheet_name="first",
            start_column_index=1,
            start_row_index=1,
            end_column_index=5,
            end_row_index=5,
            initialized=True,
        )

        self.object = Chart(
            frame,
            "Object",
            [l],
            stacking=False,
            title="pytest",
            x_axis_label="X LABEL",
            y_axis_label="Y LABEL",
            x_min=0,
            x_max=100,
            y_min=-10,
            y_max=15,
            x_axis_format="0.00",
            y_axis_format="CURRENCY",
            palette=None,
            legend_position="RIGHT_LEGEND",
        )

    def test_repr(self):
        self.object.__repr__()
        assert True

    @pytest.mark.xfail(reason=RuntimeError)
    def test_chart_id(self):
        assert self.object.chart_id == None

    def test_determine_chart_type(self):
        assert self.object._determine_chart_type([Series.line()]) == "LINE"
        assert (
            self.object._determine_chart_type([Series.line(), Series.column()])
            == "COMBO"
        )

    def test_check_stacking(self):
        assert self.object._check_stacking([Series.line()], False) == None
        assert (
            self.object._check_stacking([Series.line(), Series.column()], True) == None
        )

    def test_resolve_series(self):
        assert self.object._resolve_series()["Blue"].type == "Line"

    def test_render_basic_chart_json(self):
        assert (
            self.object.render_basic_chart_json((600, 371))["chart"]["spec"]["title"]
            == "pytest"
        )

    def test_render_histogram_chart_json(self):
        assert (
            self.object.render_histogram_chart_json((600, 371))["chart"]["spec"][
                "title"
            ]
            == "pytest"
        )

    def test_create(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.sheet_service", property(mock_service)
        )

        def mock_execute(self):
            return {"chartId": 11111}

        monkeypatch.setattr(MockService, "execute", mock_execute)

        def mock_format_frame(self, column_mapping):
            return

        monkeypatch.setattr("gslides.frame.Frame.format_frame", mock_format_frame)

        self.object.create()
        assert self.object.ch_id == 11111
