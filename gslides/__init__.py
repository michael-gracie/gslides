# -*- coding: utf-8 -*-

"""Top-level package for gslides."""

__author__ = """Michael Gracie"""
__email__ = ""
__version__ = "0.1.0"

from .config import Creds


creds = Creds()


def intialize_credentials(credentials):
    creds.set_credentials(credentials)


from .addchart import Area, Chart, Column, Histogram, Line, Scatter  # noqa
from .colors import Palette  # noqa
from .sheetsframe import CreateFrame, CreateSheet, CreateTab, GetFrame  # noqa
from .slides import CreatePresentation, CreateSlide  # noqa
