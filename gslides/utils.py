import re


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


def json_chunk_extract(obj, val):
    """Recursively fetch chunks from nested JSON."""
    arr = []

    def extract(obj, arr, val):
        """Recursively search for keys in JSON tree."""
        if isinstance(obj, dict):
            if val in obj.values():
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
