# -*- coding: utf-8 -*-

"""Top-level package for gslides."""

__author__ = """Michael Gracie"""
__email__ = ""
__version__ = "0.1.1"

from typing import Optional

from google.oauth2.credentials import Credentials

from .config import CHART_PARAMS, Creds, Font, PackagePalette

creds = Creds()
package_font = Font()
package_palette = PackagePalette()


def initialize_credentials(credentials: Optional[Credentials]) -> None:
    """Intializes credentials for all classes in the package.

    :param credentials: Credentials to build api connection
    :type credentials: google.oauth2.credentialsCredentials
    """
    creds.set_credentials(credentials)


def set_font(font: str) -> None:
    """Sets the font for all objects

    :param font: Font
    :type font: str
    """
    package_font.set_font(font)


def set_palette(palette: str) -> None:
    """Sets the palette for all charts

    :param palette: The palette to use
    :type palette: str
    """
    package_palette.set_palette(palette)


from .chart import Chart, Series  # noqa
from .colors import Palette  # noqa
from .frame import Frame  # noqa
from .presentation import Presentation  # noqa
from .spreadsheet import Spreadsheet  # noqa
from .table import Table  # noqa
