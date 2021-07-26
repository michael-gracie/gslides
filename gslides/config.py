# -*- coding: utf-8 -*-
import logging
import os
from typing import Dict, Optional

import yaml
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import Resource, build

logger = logging.getLogger(__name__)

CURR_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(CURR_DIR, "config/chart_params.yaml"), "r") as f:
    CHART_PARAMS: Dict[str, Dict] = yaml.safe_load(f)

with open(os.path.join(CURR_DIR, "config/presentation_params.yaml"), "r") as f:
    PRESENTATION_PARAMS: Dict[str, Dict] = yaml.safe_load(f)


class Creds:
    """The credentials object to build the connections to the APIs"""

    def __init__(self) -> None:
        """Constructor method"""
        self.crdtls: Optional[Credentials] = None
        self.sht_srvc: Optional[Resource] = None
        self.sld_srvc: Optional[Resource] = None

    def set_credentials(self, credentials: Optional[Credentials]) -> None:
        """Sets the credentials

        :param credentials: :class:`google.oauth2.credentials.Credentials`
        :type credentials: :class:`google.oauth2.credentials.Credentials`

        """
        self.crdtls = credentials
        logger.info("Building sheets connection")
        self.sht_srvc = build("sheets", "v4", credentials=credentials)
        logger.info("Built sheets connection")
        logger.info("Building slides connection")
        self.sld_srvc = build("slides", "v1", credentials=credentials)
        logger.info("Built slides connection")

    @property
    def sheet_service(self) -> Resource:
        """Returns the connects to the sheets API

        :raises RuntimeError: Must run set_credentials before executing method
        :return: API connection
        :rtype: :class:`googleapiclient.discovery.Resource`
        """
        if self.sht_srvc:
            return self.sht_srvc
        else:
            raise RuntimeError("Must run set_credentials before executing method")

    @property
    def slide_service(self) -> Resource:
        """Returns the connects to the slides API

        :raises RuntimeError: Must run set_credentials before executing method
        :return: API connection
        :rtype: :class:`googleapiclient.discovery.Resource`
        """
        if self.sht_srvc:
            return self.sld_srvc
        else:
            raise RuntimeError("Must run set_credentials before executing method")


class Font:
    """The credentials object to build the connections to the APIs"""

    def __init__(self) -> None:
        """Constructor method"""
        self.font: str = "Arial"

    def set_font(self, font: str) -> None:
        """Sets the font

        :param font: Font
        :type font: str

        """
        self.font = font


class PackagePalette:
    """The credentials object to build the connections to the APIs"""

    def __init__(self) -> None:
        """Constructor method"""
        self.palette: Optional[str] = None

    def set_palette(self, palette: Optional[str]) -> None:
        """Sets the palette

        :param palette: Palette
        :type palette: str

        """
        self.palette = palette
