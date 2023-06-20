from dataclasses import dataclass


@dataclass
class Credentials:
    usr: str
    psw: str


@dataclass
class Process:
    name: str
    path: str


@dataclass
class Dimension:
    width: int
    height: int
