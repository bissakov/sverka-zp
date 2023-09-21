from dataclasses import dataclass


@dataclass
class Credentials:
    user: str
    password: str


@dataclass
class Process:
    name: str
    path: str


@dataclass
class Dimension:
    width: int
    height: int
