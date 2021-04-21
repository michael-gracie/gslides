import os

import yaml

from .utils import hex_to_rgb, validate_hex_color_code


CURR_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(CURR_DIR, "config/color_mapping.yaml"), "r") as f:
    color_mapping = yaml.safe_load(f)

with open(os.path.join(CURR_DIR, "config/base_palettes.yaml"), "r") as f:
    base_palettes = yaml.safe_load(f)


def translate_color(color):
    if color in color_mapping.keys():
        return color_mapping[color]
    else:
        return color


class Palette:
    def __init__(self, palette=None):
        if palette:
            self.colors = base_palettes[palette]
        else:
            self.colors = base_palettes["black"]
        self._clean_palette()
        self.index = 0

    def load_palette(self, name):
        self.colors = base_palettes[name]
        self._clean_palette
        self.index = 0

    def _clean_palette(self):
        self.colors = [
            validate_hex_color_code(translate_color(color)) for color in self.colors
        ]

    def __iter__(self):
        return self

    def __next__(self):
        color = self.colors[self.index]
        if self.index + 1 >= len(self.colors):
            self.index = 0
        else:
            self.index += 1
        return hex_to_rgb(color)
