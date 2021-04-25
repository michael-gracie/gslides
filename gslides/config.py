# -*- coding: utf-8 -*-
import os

from typing import Dict, Optional

import yaml

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import Resource, build


CURR_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(CURR_DIR, "config/chart_params.yaml"), "r") as f:
    CHART_PARAMS: Dict[str, Dict] = yaml.safe_load(f)


class Creds:
    def __init__(self) -> None:
        self.crdtls: Optional[Credentials] = None
        self.sht_srvc: Optional[Resource] = None
        self.sld_srvc: Optional[Resource] = None

    def set_credentials(self, credentials: Optional[Credentials]) -> None:
        self.crdtls = credentials
        self.sht_srvc = build("sheets", "v4", credentials=credentials)
        self.sld_srvc = build("slides", "v1", credentials=credentials)

    @property
    def sheet_service(self) -> Optional[Resource]:
        if self.sht_srvc:
            return self.sht_srvc
        else:
            raise RuntimeError("Must run set_credentials before executing method")

    @property
    def slide_service(self) -> Optional[Resource]:
        if self.sht_srvc:
            return self.sld_srvc
        else:
            raise RuntimeError("Must run set_credentials before executing method")
