# -*- coding: utf-8 -*-

"""Top-level package for gslides."""

__author__ = """Michael Gracie"""
__email__ = ""
__version__ = "0.1.0"

from .addchart import Area, Chart, Column, Histogram, Line, Scatter
from .colors import Palette
from .sheetsframe import CreateFrame, CreateSheet, CreateTab, GetFrame
from .slides import CreatePresentation, CreateSlide
