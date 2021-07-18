import pandas as pd
import pytest

from gslides.frame import CreateFrame, Frame, GetFrame, format_type, get_sheet_data


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


def test_get_sheet_data(monkeypatch):
    def mock_service(self):
        return MockService()

    monkeypatch.setattr("gslides.config.Creds.sheet_service", property(mock_service))

    def mock_return(self):
        return {"values": [["test"], ["0"], ["1"]]}

    monkeypatch.setattr(MockService, "execute", mock_return)
    assert get_sheet_data("abc123", "first", 0, 0, 2, 0) == [["test"], ["0"], ["1"]]


@pytest.mark.parametrize(
    "input,expected",
    [
        ("$0.00", {"type": "NUMBER", "pattern": "$0.00"}),
        ("CURRENCY", {"type": "NUMBER", "pattern": "$0.00"}),
        ("NUMBER", {"type": "NUMBER"}),
    ],
)
def test_format_type(input, expected):
    assert format_type(input) == expected


class TestGetFrame:
    def setup(self):
        self.object = GetFrame(
            spreadsheet_id="abc123",
            sheet_id=1234,
            sheet_name="first",
            anchor_cell="A1",
            bottom_right_cell="B4",
        )

    def test_execute(self, monkeypatch):
        def mock_data_return(
            spreadsheet_id,
            sheet_name,
            start_column_index,
            start_row_index,
            end_column_index,
            end_row_index,
        ):
            return [["test"], ["0"], ["1"]]

        monkeypatch.setattr("gslides.frame.get_sheet_data", mock_data_return)
        assert self.object.execute() == True


class TestCreateFrame:
    def setup(self):
        self.object = CreateFrame(
            df=test_df(), spreadsheet_id="abc123", sheet_name="first"
        )

    def test_calc_end_index(self):
        assert self.object._calc_end_index() == (5, 5)

    def test_render_update_json(self):
        assert self.object.render_update_json()["data"][0]["range"] == "first!A1:E1"

    @pytest.mark.xfail(reason=RuntimeError)
    def test_execute_overwrite(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.sheet_service", property(mock_service)
        )

        def mock_data_return(
            spreadsheet_id,
            sheet_name,
            start_column_index,
            start_row_index,
            end_column_index,
            end_row_index,
        ):
            return [["test"], ["0"], ["1"]]

        monkeypatch.setattr("get_sheet_data", mock_data_return)
        self.object.execute()

    def test_execute_no_overwrite(self, monkeypatch):
        self.object.overwrite_data = True

        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.sheet_service", property(mock_service)
        )

        def mock_data_return(
            spreadsheet_id,
            sheet_name,
            start_column_index,
            start_row_index,
            end_column_index,
            end_row_index,
        ):
            return [["test"], ["0"], ["1"]]

        monkeypatch.setattr("gslides.frame.get_sheet_data", mock_data_return)
        assert self.object.execute() == True


class TestFrame:
    def setup(self):
        self.object = Frame(
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

    def test_repr(self):
        self.object.__repr__()
        assert True

    def test_create(self, monkeypatch):
        def mock_return(self):
            return True

        monkeypatch.setattr(CreateFrame, "execute", mock_return)
        assert (
            dir(
                Frame.create(
                    df=test_df(),
                    spreadsheet_id="abc123",
                    sheet_id=1234,
                    sheet_name="first",
                )
            )
            == dir(self.object)
        )

    def test_get(self, monkeypatch):
        def mock_return(self):
            return True

        monkeypatch.setattr(GetFrame, "execute", mock_return)
        assert (
            Frame.get(
                spreadsheet_id="abc123",
                sheet_id=1234,
                sheet_name="first",
                anchor_cell="A1",
                bottom_right_cell="B4",
            ).initialized
            == True
        )

    def test_format_frame(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.sheet_service", property(mock_service)
        )
        self.object.format_frame({"Blue": "CURRENCY"})
        assert True

    def test_render_format_frame(self):
        assert (
            self.object.render_format_frame({"Blue": "CURRENCY"})["requests"][0][
                "updateCells"
            ]["rows"][0]["values"][0]["userEnteredFormat"]["numberFormat"]["pattern"]
            == "$0.00"
        )

    def test_data(self):
        assert self.object.data == self.object
