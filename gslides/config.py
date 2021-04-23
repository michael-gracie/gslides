# -*- coding: utf-8 -*-
import os

from typing import Dict

import yaml


CURR_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(CURR_DIR, "config/chart_params.yaml"), "r") as f:
    CHART_PARAMS: Dict[str, Dict] = yaml.safe_load(f)
