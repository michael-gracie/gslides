# -*- coding: utf-8 -*-

"""Top-level package for gslides."""

__author__ = """Michael Gracie"""
__email__ = ""
__version__ = "0.1.0"

from typing import Optional

from google.oauth2.credentials import Credentials

from .config import Creds


creds = Creds()


def intialize_credentials(credentials: Optional[Credentials]) -> None:
    """Intializes credentials for all classes in the package.

    :param credentials: Credentials to build api connection
    :type credentials: google.oauth2.credentialsCredentials
    """
    creds.set_credentials(credentials)


from .addchart import Area, Chart, Column, Histogram, Line, Scatter  # noqa
from .colors import Palette  # noqa
from .sheetsframe import CreateFrame, CreateSheet, CreateTab, GetFrame  # noqa
from .slides import CreatePresentation, CreateSlide  # noqa
