# -*- coding: utf-8 -*-
import datetime
import re

from decimal import Decimal
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd

from .config import CHART_PARAMS


def json_val_extract(obj: Dict[str, Any], key: str) -> Optional[Any]:
    """Recursively fetch chunks from nested JSON."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == key:
                return v
            else:
                if json_val_extract(v, key):
                    return json_val_extract(v, key)
        return None
    elif isinstance(obj, list):
        for item in obj:
            if json_val_extract(item, key):
                return json_val_extract(item, key)
        return None
    else:
        return None


def json_chunk_extract(
    obj: Dict[str, Any], key: str, val: Union[str, int, float]
) -> List:
    """Recursively fetch chunks from nested JSON."""
    arr: List = []

    def extract(obj: Dict[str, Any], arr: List, val: Union[str, int, float]) -> List:
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


def num_to_char(x: int) -> str:
    if x <= 26:
        return chr(x + 64).upper()
    elif x <= 26 * 26:
        return f"{chr(x//26+64).upper()}{chr(x%26+64).upper()}"
    else:
        raise ValueError("Integer too high to be converted")


def char_to_num(x: str) -> int:
    total = 0
    for i in range(len(x)):
        total += (ord(x[::-1][i]) - 64) * (26 ** i)
    return total


def cell_to_num(x: str) -> Tuple[int, int]:
    regex = re.compile("([A-Z]*)([0-9]*)")
    output = regex.search(x)
    if output:
        column = output.group(1)
        row = int(output.group(2))
        return (row, char_to_num(column))
    else:
        raise ValueError("Invalid cell format")


def validate_hex_color_code(x: str) -> str:
    match = re.search("^#(?:[0-9a-fA-F]{3}){1,2}$", x)
    if match:
        return x
    else:
        raise ValueError("Input a hexadecimal color code")


def hex_to_rgb(x: str) -> Tuple[float, ...]:
    x = x[1:]
    return tuple(int(x[i : i + 2], 16) / 255 for i in (0, 2, 4))  # noqa


def emu_to_px(x: int) -> int:
    return int(x * 220 / (914400))


def optimize_size(y_scale: float, area: float = 222600) -> Tuple[float, float]:
    x_length = (area / (y_scale)) ** 0.5
    return (x_length, x_length * y_scale)


def clean_list_of_list(x: List[List]) -> List[List]:
    max_len = max([len(i) for i in x])
    for i in x:
        i.extend([None] * (max_len - len(i)))
    return x


def clean_nan(df: pd.DataFrame) -> pd.DataFrame:
    return df.replace({np.nan: None})


def clean_dtypes(x: Any) -> Union[str, float, int, np.int64, np.float64, None]:
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


def validate_params_list(params: dict) -> None:
    for key, val in CHART_PARAMS.items():
        if key in params.keys() and params[key]:
            if params[key] not in val["params"]:
                raise ValueError(
                    f"{ params[key]} is not a valid parameter for {key}. "
                    f"Accepted parameters include {', '.join(val['params'])}. See "
                    f"{val['url']} for further documentation."
                )


def validate_params_int(params: dict) -> None:
    variables = ["line_width", "point_size", "bucket_size"]
    for var in variables:
        if var in params.keys() and params[var]:
            if type(params[var]) != int or params[var] < 0:
                raise ValueError(
                    f"{params[var]} is not a valid parameter for {var}. "
                    f"Accepted values are any integer greater than 0"
                )


def validate_params_float(params: dict) -> None:
    variables = ["outlier_percentage", "x_border", "y_border", "spacing"]
    for var in variables:
        if var in params.keys() and params[var]:
            if params[var] >= 1 or params[var] < 0:
                raise ValueError(
                    f"{params[var]} is not a valid parameter for {var}. "
                    f"Accepted values are any float between 0 and 1"
                )


def validate_cell_name(x: str) -> str:
    pattern = re.compile("([A-Z]{1,2})([0-9]+)")
    output = pattern.search(x)
    if output:
        if output[0] == x:
            return x
        else:
            raise ValueError("Invalid cell name.")
    else:
        raise ValueError("Invalid cell name.")
