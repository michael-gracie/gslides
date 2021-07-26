from decimal import Decimal

import numpy as np
import pandas as pd
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
    "key,expected", [("heightPixels", [200]), ("chartId", [1234567]), ("slideId", [])]
)
def test_json_val_extract(key, expected):
    assert utils.json_val_extract(json, key) == expected


def test_json_chunk_extract():
    assert utils.json_chunk_extract(json, "title", "pytest")[0]["fontName"] == "Roboto"


def test_json_chunk_key_extract():
    assert utils.json_chunk_key_extract(json, "spec")[0]["chartId"] == 1234567


def test_json_dict_extract():
    assert utils.json_dict_extract(json, ("title", "fontName")) == {"pytest": "Roboto"}


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


@pytest.mark.parametrize(
    "input,expected",
    [
        ("#66ff66", "#66ff66"),
        pytest.param("#61ab6", None, marks=pytest.mark.xfail(reason=ValueError)),
    ],
)
def test_validate_hex_color_code(input, expected):
    assert utils.validate_hex_color_code(input) == expected


def test_hex_to_rgb():
    assert utils.hex_to_rgb("#61ab96") == (
        0.3803921568627451,
        0.6705882352941176,
        0.5882352941176471,
    )


def test_emu_to_px():
    assert utils.emu_to_px(914400) == 220


def test_optimize_size():
    assert utils.optimize_size(2) == (333.61654635224556, 667.2330927044911)


def test_clean_list_of_list():
    data = [
        ["Object", "Blue", "Red", "Grand Total"],
        ["Ball", "6"],
        ["Cube", "6", "4", "10"],
        ["Stick", "7", "5", "12"],
    ]
    data = utils.clean_list_of_list(data)
    assert data[1][3] == None


def test_clean_nan():
    df = pd.DataFrame({"test": [np.nan, 1]})
    df = utils.clean_nan(df)
    assert df["test"][0] == None


@pytest.mark.parametrize(
    "input,expected",
    [
        (1, 1),
        ("string", "string"),
        (0.1, 0.1),
        (Decimal(0.1), 0.1),
        (pd.Timestamp("2020-01-01"), "2020-01-01 00:00:00"),
        (pd.Timestamp("2020-01-01").date(), "2020-01-01"),
    ],
)
def test_clean_dtypes(input, expected):
    assert utils.clean_dtypes(input) == expected


@pytest.mark.xfail(reason=ValueError)
def test_validate_params_list():
    utils.validate_params_list({"legend_position": "outside"})


@pytest.mark.xfail(reason=ValueError)
def test_validate_params_int():
    utils.validate_params_int({"line_width": -1})


@pytest.mark.xfail(reason=ValueError)
def test_validate_params_int():
    utils.validate_params_int({"outlier_percentage": 1.5})


@pytest.mark.xfail(reason=ValueError)
def test_validate_series_columns():
    utils.validate_series_columns("col")


@pytest.mark.xfail(reason=ValueError)
def test_validate_cell_name():
    utils.validate_cell_name("1A:4")


def test_df():
    data = [
        ["Object", "Blue", "Red", "Grand Total"],
        ["Ball", "6", "1", "7"],
        ["Cube", "6", "4", "10"],
        ["Stik", "7", "5", "12"],
    ]
    return pd.DataFrame(columns=data[0], data=data[1:])


def test_determine_col_proportion():
    assert list(utils.determine_col_proportion(test_df())) == [0.5, 0.125, 0.125, 0.25]


def test_black_or_white():
    assert utils.black_or_white((0.9, 0.9, 0.9)) == (0, 0, 0)
