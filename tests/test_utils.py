import pytest

import gslides.utils as utils


json = {
    "chartId": 1234567,
    "spec": {
        "title": "pytest",
        "basicChart": {
            "chartType": "COMBO",
            "legendPosition": "RIGHT_LEGEND",
            "axis": [
                {
                    "position": "BOTTOM_AXIS",
                    "title": "Object",
                    "format": {"fontFamily": "Roboto"},
                    "viewWindowOptions": {},
                },
                {"position": "LEFT_AXIS", "viewWindowOptions": {}},
            ],
            "domains": [
                {
                    "domain": {
                        "sourceRange": {
                            "sources": [
                                {
                                    "sheetId": 1111111,
                                    "startRowIndex": 0,
                                    "endRowIndex": 5,
                                    "startColumnIndex": 4,
                                    "endColumnIndex": 5,
                                }
                            ]
                        }
                    }
                }
            ],
            "series": [
                {
                    "series": {
                        "sourceRange": {
                            "sources": [
                                {
                                    "sheetId": 1111111,
                                    "startRowIndex": 0,
                                    "endRowIndex": 5,
                                    "startColumnIndex": 5,
                                    "endColumnIndex": 6,
                                }
                            ]
                        }
                    },
                    "targetAxis": "LEFT_AXIS",
                    "type": "COLUMN",
                    "dataLabel": {
                        "type": "NONE",
                        "textFormat": {"fontFamily": "Roboto"},
                    },
                },
                {
                    "series": {
                        "sourceRange": {
                            "sources": [
                                {
                                    "sheetId": 1111111,
                                    "startRowIndex": 0,
                                    "endRowIndex": 5,
                                    "startColumnIndex": 6,
                                    "endColumnIndex": 7,
                                }
                            ]
                        }
                    },
                    "targetAxis": "LEFT_AXIS",
                    "type": "COLUMN",
                    "dataLabel": {
                        "type": "NONE",
                        "textFormat": {"fontFamily": "Roboto"},
                    },
                },
                {
                    "series": {
                        "sourceRange": {
                            "sources": [
                                {
                                    "sheetId": 1111111,
                                    "startRowIndex": 0,
                                    "endRowIndex": 5,
                                    "startColumnIndex": 7,
                                    "endColumnIndex": 8,
                                }
                            ]
                        }
                    },
                    "targetAxis": "LEFT_AXIS",
                    "type": "LINE",
                    "dataLabel": {
                        "type": "NONE",
                        "textFormat": {"fontFamily": "Roboto"},
                    },
                },
            ],
            "stackedType": "STACKED",
        },
        "hiddenDimensionStrategy": "SKIP_HIDDEN_ROWS_AND_COLUMNS",
        "titleTextFormat": {"fontFamily": "Roboto"},
        "fontName": "Roboto",
    },
    "position": {
        "overlayPosition": {
            "anchorCell": {"sheetId": 1111111},
            "offsetXPixels": 5,
            "widthPixels": 100,
            "heightPixels": 200,
        }
    },
}


@pytest.mark.parametrize(
    "key,expected", [("heightPixels", 200), ("chartId", 1234567), ("slideId", None)]
)
def test_json_val_extract(key, expected):
    assert utils.json_val_extract(json, key) == expected


def test_json_chunk_extract():
    assert utils.json_chunk_extract(json, "pytest")[0]["fontName"] == "Roboto"


@pytest.mark.parametrize(
    "input,expected",
    [
        (3, "C"),
        (27, "AA"),
        pytest.param(700, None, marks=pytest.mark.xfail(reason=ValueError)),
    ],
)
def test_num_to_char(input, expected):
    assert utils.num_to_char(input) == expected


def char_to_num(x):
    total = 0
    for i in range(len(x)):
        total += (ord(x[::-1][i]) - 64) * (26 ** i)
    return total


@pytest.mark.parametrize(
    "input,expected",
    [
        ("C", 3),
        ("AA", 27),
    ],
)
def test_char_to_num(input, expected):
    assert utils.char_to_num(input) == expected


@pytest.mark.parametrize(
    "input,expected",
    [
        ("C10", (10, 3)),
        ("AA100", (100, 27)),
    ],
)
def test_cell_to_num(input, expected):
    assert utils.cell_to_num(input) == expected
