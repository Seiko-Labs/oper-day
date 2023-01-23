import datetime
from dataclasses import dataclass
from typing import Dict


@dataclass
class Credentials:
    usr: str
    psw: str


@dataclass
class Process:
    name: str
    path: str
