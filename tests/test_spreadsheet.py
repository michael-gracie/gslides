import pytest

from gslides.spreadsheet import (
    AddSheet,
    CreateSpreadsheet,
    GetSpreadsheet,
    RemoveSheet,
    Spreadsheet,
)


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


def test_create_spreadsheet_render_json():
    json = {
        "properties": {
            "title": "Test",
            "locale": "en_US",
            "autoRecalc": "HOUR",
        },
        "sheets": [{"properties": {"title": "first"}}],
    }
    assert CreateSpreadsheet().render_json(title="Test", sheet_names=["first"]) == json


def test_create_spreadsheet_execute(monkeypatch):
    def mock_service(self):
        return MockService()

    monkeypatch.setattr("gslides.config.Creds.sheet_service", property(mock_service))

    def mock_return(self):
        return {"sheetId": 1234, "spreadsheetId": "abc123"}

    monkeypatch.setattr(MockService, "execute", mock_return)
    assert CreateSpreadsheet().execute(title="Test", sheet_names=["first"]) == (
        "abc123",
        [1234],
        True,
    )


def test_get_spreadsheet_execute(monkeypatch):
    def mock_service(self):
        return MockService()

    monkeypatch.setattr("gslides.config.Creds.sheet_service", property(mock_service))

    def mock_return(self):
        return {
            "properties": {"title": "Test"},
            "sheets": [{"sheetId": 1234, "title": "first"}],
        }

    monkeypatch.setattr(MockService, "execute", mock_return)
    assert GetSpreadsheet().execute(spreadsheet_id="abc123") == (
        "Test",
        {"first": 1234},
        True,
    )


def test_add_sheet_render_json():
    json = {"requests": [{"addSheet": {"properties": {"title": "first"}}}]}
    assert AddSheet().render_json(sheet_names=["first"]) == json


def test_add_sheet_execute(monkeypatch):
    def mock_service(self):
        return MockService()

    monkeypatch.setattr("gslides.config.Creds.sheet_service", property(mock_service))

    def mock_return(self):
        return {"sheets": [{"sheetId": 1234, "title": "first"}]}

    monkeypatch.setattr(MockService, "execute", mock_return)
    assert AddSheet().execute(spreadsheet_id="abc123", sheet_names=["first"]) == {
        "first": 1234
    }


def test_remove_sheet_render_json():
    json = {"requests": [{"deleteSheet": {"sheetId": 1234}}]}
    assert RemoveSheet().render_json(sheet_ids=[1234]) == json


def test_remove_sheet_execute(monkeypatch):
    def mock_service(self):
        return MockService()

    monkeypatch.setattr("gslides.config.Creds.sheet_service", property(mock_service))

    assert RemoveSheet().execute(spreadsheet_id="abc123", sheet_ids=[1234]) == [1234]


class TestSpreadsheet:
    def setup(self):
        self.object = Spreadsheet(
            sp_id="abc123",
            title="Test",
            sht_ids={"first": 1234},
            initialized=True,
        )

    def test_repr(self):
        self.object.__repr__()
        assert True

    def test_create(self, monkeypatch):
        def mock_return(self, title, sheet_names):
            return ("abc123", [1234], True)

        monkeypatch.setattr(CreateSpreadsheet, "execute", mock_return)
        assert dir(Spreadsheet.create(title="Test", sheet_names=["first"])) == dir(
            self.object
        )

    def test_get(self, monkeypatch):
        def mock_return(self, spreadsheet_id):
            return ("Test", {"first": 1234}, True)

        monkeypatch.setattr(GetSpreadsheet, "execute", mock_return)
        assert dir(Spreadsheet.get(spreadsheet_id="abc123")) == dir(self.object)

    def test_add_sheets(self, monkeypatch):
        def mock_return(self, spreadsheet_id, sheet_names):
            return {"second": 2345}

        monkeypatch.setattr(AddSheet, "execute", mock_return)
        self.object.add_sheets(["second"])
        assert self.object.sht_nms == {"first": 1234, "second": 2345}

    def test_rm_sheets(self, monkeypatch):
        def mock_return(self, spreadsheet_id, sheet_names):
            return ["first"]

        monkeypatch.setattr(RemoveSheet, "execute", mock_return)
        self.object.rm_sheets(["first"])
        assert self.object.sht_nms == {}

    def test_sheet_names(self):
        assert self.object.sheet_names == {"first": 1234}

    def test_spreadsheet_id(self):
        assert self.object.spreadsheet_id == "abc123"
