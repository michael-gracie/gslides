import pandas as pd
import pytest

from gslides.table import Table


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


def test_df():
    data = [
        ["Object", "Blue", "Red", "Grand Total"],
        ["Ball", "6", "1", "7"],
        ["Cube", "6", "4", "10"],
        ["Stick", "7", "5", "12"],
    ]
    return pd.DataFrame(columns=data[0], data=data[1:])


class TestTable:
    def setup(self):
        self.object = Table(
            data=test_df(),
        )

    def test_repr(self):
        self.object.__repr__()
        assert True

    def test_resolve_df(self):
        pd.testing.assert_frame_equal(self.object._resolve_df(test_df()), test_df())
        assert True

    def test_reset_header(self):
        assert list(self.object._reset_header(test_df()).columns) == [0, 1, 2, 3]

    def test_render_create_table_json(self):
        assert (
            self.object.render_create_table_json(sl_id=1111)["requests"][0][
                "createTable"
            ]["rows"]
            == 4
        )

    def test_table_move_request(self):
        assert (
            self.object._table_move_request(1111, 0, 0)[0][
                "updatePageElementTransform"
            ]["objectId"]
            == 1111
        )

    def test_table_add_text_request(self):
        assert (
            self.object._table_add_text_request(1111)[0]["insertText"]["objectId"]
            == 1111
        )

    def test_table_style_text_request(self):
        assert (
            self.object._table_style_text_request(
                1111,
                True,
                False,
                12,
                (0, 0, 0),
                (0, 0, 0),
            )[0]["updateTextStyle"]["objectId"]
            == 1111
        )

    def test_table_update_cell(self):
        assert (
            self.object._table_update_cell(1111, True, False, (0, 0, 0), (0, 0, 0),)[0][
                "updateTableCellProperties"
            ]["objectId"]
            == 1111
        )

    def test_table_update_paragraph_style(self):
        assert (
            self.object._table_update_paragraph_style(1111)[0]["updateParagraphStyle"][
                "objectId"
            ]
            == 1111
        )

    def test_table_update_row(self):
        assert (
            self.object._table_update_row(1111, 1)[0]["updateTableRowProperties"][
                "objectId"
            ]
            == 1111
        )

    def test_table_update_column(self):
        assert (
            self.object._table_update_column(1111, [0.5, 0.125, 0.125, 0.25])[0][
                "updateTableColumnProperties"
            ]["objectId"]
            == 1111
        )

    def test_render_update_table_json(self):
        assert (
            self.object.render_update_table_json(1111, (300, 300), 0, 0)["requests"][0][
                "updatePageElementTransform"
            ]["objectId"]
            == 1111
        )

    def test_create(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.slide_service", property(mock_service)
        )

        def mock_return(self):
            return {"replies": [{"createTable": {"objectId": 1111}}]}

        monkeypatch.setattr(MockService, "execute", mock_return)

        self.object.create("abcd", 1111)
        assert True
