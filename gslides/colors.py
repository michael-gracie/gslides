# -*- coding: utf-8 -*-
"""
Manages the color configuration
"""

import os
from typing import Dict, List, Optional, Tuple, TypeVar

import yaml

from .utils import hex_to_rgb, validate_hex_color_code

CURR_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(CURR_DIR, "config/color_mapping.yaml"), "r") as f:
    color_mapping: Dict[str, str] = yaml.safe_load(f)

with open(os.path.join(CURR_DIR, "config/base_palettes.yaml"), "r") as f:
    base_palettes: Dict[str, List[str]] = yaml.safe_load(f)

custom_palettes_path = os.path.join(
    os.path.expanduser("~"), ".gslides/custom_palettes.yaml"
)
if os.path.isfile(custom_palettes_path):
    with open(custom_palettes_path, "r") as f:
        base_palettes.update(yaml.safe_load(f))


def translate_color(color: str) -> str:
    """Translates a color from a named color to a hex code

    :param color: Named color
    :type color: str
    :return: Hex code corresponding to the color
    :rtype: str
    :raises ValueError:

    """
    if color.lower() in color_mapping.keys():
        return color_mapping[color.lower()]
    elif color[0] == "#":
        return color
    else:
        raise ValueError(f"{color} is not a valid hex or named color")


TPalette = TypeVar("TPalette", bound="Palette")


class Palette:
    """An iterator for a list of colors.

    :param palette: Name of a palette
    :type palette: str
    """

    def __init__(self, palette: Optional[str] = None) -> None:
        """Constructor method"""
        if palette:
            if palette in base_palettes.keys():
                self.colors = base_palettes[palette]
            else:
                raise ValueError(
                    f'{palette} is not in available palettes: {", .".join(base_palettes.keys())}'
                )
        else:
            self.colors = base_palettes["black"]
        self._clean_palette()
        self.index = 0

    def load_palette(self, name: str) -> None:
        """Loads a given palette

        :param palette: Name of a palette
        :type palette: str

        """
        self.colors = base_palettes[name]
        self._clean_palette()
        self.index = 0

    def _clean_palette(self) -> None:
        """Cleans and transletes colors in the palette"""
        self.colors = [
            validate_hex_color_code(translate_color(color)) for color in self.colors
        ]

    def __iter__(self: TPalette) -> TPalette:
        """Iterator function

        :return: :class:`Palette`
        :rtype: :class:`Palette`
        """
        return self

    def __next__(self) -> Tuple[float, ...]:
        """Next function

        :return: A tuple for the r, g, b values of the next color in the iterator
        :rtype: tuple
        """
        color = self.colors[self.index]
        if self.index + 1 >= len(self.colors):
            self.index = 0
        else:
            self.index += 1
        return hex_to_rgb(color)
