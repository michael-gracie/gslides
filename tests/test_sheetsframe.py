import pandas as pd
import pytest

from gslides.sheetsframe import (
    CreateFrame,
    CreateSheet,
    CreateSpreadsheet,
    GetFrame,
    SheetsFrame,
)


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


class TestCreateSheet:
    def setup(self):
        self.object = CreateSheet(
            spreadsheet_id="abc123",
            sheet_name="Test",
        )

    def test_render_json(self):
        assert self.object.render_json() == {
            "requests": [{"addSheet": {"properties": {"title": "Test"}}}]
        }

    @pytest.mark.xfail(reason=RuntimeError)
    def test_sheet_id(self):
        assert self.object.sheet_id == None

    def test_execute(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.sheet_service", property(mock_service)
        )

        def mock_return(self):
            return {"sheetId": 1234}

        monkeypatch.setattr(MockService, "execute", mock_return)
        self.object.execute()
        assert self.object.sh_id == 1234
        assert self.object.executed


class TestCreateSpreadsheet:
    def setup(self):
        self.object = CreateSpreadsheet(
            title="Test",
            sheet_name="first",
        )

    def test_render_json(self):
        json = {
            "properties": {
                "title": "Test",
                "locale": "en_US",
                "autoRecalc": "HOUR",
            },
            "sheets": [{"properties": {"title": "first"}}],
        }
        assert self.object.render_json() == json

    @pytest.mark.xfail(reason=RuntimeError)
    def test_sheet_id(self):
        assert self.object.sheet_id == None

    @pytest.mark.xfail(reason=RuntimeError)
    def test_spreadsheet_id(self):
        assert self.object.spreadsheet_id == None

    def test_execute(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.sheet_service", property(mock_service)
        )

        def mock_return(self):
            return {"sheetId": 1234, "spreadsheetId": "abc123"}

        monkeypatch.setattr(MockService, "execute", mock_return)
        self.object.execute()
        assert self.object.sh_id == 1234
        assert self.object.sp_id == "abc123"
        assert self.object.executed


class TestGetFrame:
    def setup(self):
        self.object = GetFrame(
            spreadsheet_id="abc123",
            sheet_id=1234,
            anchor_cell="A1",
            bottom_right_cell="B4",
        )

    def test_execute(self, monkeypatch):
        def mock_name_return(self):
            return "first"

        def mock_data_return(self, sheet_name):
            return [["test"], ["0"], ["1"]]

        monkeypatch.setattr(GetFrame, "_get_sheet_name", mock_name_return)
        monkeypatch.setattr(GetFrame, "_get_sheet_data", mock_data_return)
        self.object.execute()
        assert self.object.executed


class TestCreateFrame:
    def setup(self):
        self.object = CreateFrame(
            df=test_df(),
            spreadsheet_id="abc123",
            sheet_id=1234,
        )

    def test_calc_end_index(self):
        assert self.object._calc_end_index() == (5, 5)

    def test_render_update_json(self):
        assert (
            self.object.render_update_json("test")["data"][0]["range"] == "test!A1:E1"
        )

    @pytest.mark.xfail(reason=RuntimeError)
    def test_execute_overwrite(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.sheet_service", property(mock_service)
        )

        def mock_name_return(self):
            return "first"

        def mock_data_return(self, sheet_name):
            return [["test"], ["0"], ["1"]]

        monkeypatch.setattr(CreateFrame, "_get_sheet_name", mock_name_return)
        monkeypatch.setattr(CreateFrame, "_get_sheet_data", mock_data_return)
        service = MockService()
        self.object.execute()

    def test_execute_no_overwrite(self, monkeypatch):
        self.object.overwrite_data = True

        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.sheet_service", property(mock_service)
        )

        def mock_name_return(self):
            return "first"

        def mock_data_return(self, sheet_name):
            return None

        monkeypatch.setattr(CreateFrame, "_get_sheet_name", mock_name_return)
        monkeypatch.setattr(CreateFrame, "_get_sheet_data", mock_data_return)
        self.object.execute()
        assert self.object.executed


class TestSheetsFrame:
    def setup(self):
        self.object = SheetsFrame("abc123", 1234, 1, 1, 2, 2, pd.DataFrame())

    def test_get_sheet_name(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.sheet_service", property(mock_service)
        )

        def mock_return(self):
            return [{"title": "first", "sheetId": 1234}]

        monkeypatch.setattr(MockService, "execute", mock_return)
        assert self.object._get_sheet_name() == "first"

    def test_get_sheet_data(self, monkeypatch):
        def mock_service(self):
            return MockService()

        monkeypatch.setattr(
            "gslides.config.Creds.sheet_service", property(mock_service)
        )

        def mock_return(self):
            return {"values": 0}

        monkeypatch.setattr(MockService, "execute", mock_return)
        assert self.object._get_sheet_data("first") == 0

    @pytest.mark.xfail(reason=RuntimeError)
    def test_data(self):
        assert self.object.data == None
