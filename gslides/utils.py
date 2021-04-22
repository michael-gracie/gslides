import datetime
import re

from decimal import Decimal

import numpy as np
import pandas as pd

from .config import CHART_PARAMS


def json_val_extract(obj, key):
    """Recursively fetch chunks from nested JSON."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == key:
                return v
            else:
                if json_val_extract(v, key):
                    return json_val_extract(v, key)
    elif isinstance(obj, list):
        for item in obj:
            if json_val_extract(item, key):
                return json_val_extract(item, key)
    else:
        return None


def json_chunk_extract(obj, key, val):
    """Recursively fetch chunks from nested JSON."""
    arr = []

    def extract(obj, arr, val):
        """Recursively search for keys in JSON tree."""
        if isinstance(obj, dict):
            if (key, val) in obj.items():
                arr.append(obj)
            else:
                for k, v in obj.items():
                    extract(v, arr, val)
        elif isinstance(obj, list):
            for item in obj:
                extract(item, arr, val)
        return arr

    values = extract(obj, arr, val)
    return values


def num_to_char(x):
    if x <= 26:
        return chr(x + 64).upper()
    elif x <= 26 * 26:
        return f"{chr(x//26+64).upper()}{chr(x%26+64).upper()}"
    else:
        raise ValueError("Integer too high to be converted")


def char_to_num(x):
    total = 0
    for i in range(len(x)):
        total += (ord(x[::-1][i]) - 64) * (26 ** i)
    return total


def cell_to_num(x):
    regex = re.compile("([A-Z]*)([0-9]*)")
    column = regex.search(x).group(1)
    row = int(regex.search(x).group(2))
    return (row, char_to_num(column))


def validate_hex_color_code(x):
    match = re.search("^#(?:[0-9a-fA-F]{3}){1,2}$", x)
    if match:
        return x
    else:
        raise ValueError("Input a hexadecimal color code")


def hex_to_rgb(x):
    x = x[1:]
    return tuple(int(x[i : i + 2], 16) / 255 for i in (0, 2, 4))  # noqa


def emu_to_px(x):
    return int(x * 220 / (914400))


def optimize_size(y_scale, area=222600):
    x_length = (area / (y_scale)) ** 0.5
    return (x_length, x_length * y_scale)


def clean_list_of_list(x):
    max_len = max([len(i) for i in x])
    for i in x:
        i.extend([None] * (max_len - len(i)))
    return x


def clean_nan(df):
    return df.replace({np.nan: None})


def clean_dtypes(x):
    if type(x) in [pd._libs.tslibs.timestamps.Timestamp, datetime.date]:
        return str(x)
    elif type(x) in [Decimal]:
        return float(str(x))
    elif type(x) in [str, int, float, np.int64, np.float64, type(None)]:
        return x
    else:
        raise TypeError(
            f"{type(x)} is not an accepted datatype. Type must conform to "
            "str, int, float, NoneType, decimal.Decimal, pd.Timestamp, datetime.date"
        )


def validate_params_list(params):
    for key, val in CHART_PARAMS.items():
        if key in params.keys() and params[key]:
            if params[key] not in val["params"]:
                raise ValueError(
                    f"{ params[key]} is not a valid parameter for {key}. "
                    f"Accepted parameters include {', '.join(val['params'])}. See "
                    f"{val['url']} for further documentation."
                )


def validate_params_int(params):
    variables = ["line_width", "point_size", "bucket_size"]
    for var in variables:
        if var in params.keys() and params[var]:
            if type(params[var]) != int or params[var] < 0:
                raise ValueError(
                    f"{params[var]} is not a valid parameter for {var}. "
                    f"Accepted values are any integer greater than 0"
                )


def validate_params_float(params):
    variables = ["outlier_percentage", "x_border", "y_border", "spacing"]
    for var in variables:
        if var in params.keys() and params[var]:
            if params[var] >= 1 or params[var] < 0:
                raise ValueError(
                    f"{params[var]} is not a valid parameter for {var}. "
                    f"Accepted values are any float between 0 and 1"
                )


def validate_cell_name(x):
    pattern = re.compile("([A-Z]{1,2})([0-9]+)")
    if pattern.search(x)[0] == x:
        return x
    else:
        raise ValueError("Invalid cell name.")
