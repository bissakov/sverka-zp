import os
from dataclasses import dataclass
from typing import List


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


@dataclass
class ExcelInfo:
    path: str
    name: str


@dataclass
class ArchiveInfo:
    zip_dir: str
    zip_name: str
    file_name: str = None

    def __post_init__(self) -> None:
        self.file_name = os.path.splitext(self.zip_name)[0]


@dataclass
class EmailInfo:
    email_list: List[str]
    subject: str = None
    body: str = None
    attachment: str = None


@dataclass
class Data:
    usr: str
    psw: str
    process_name: str
    process_path: str
    excel_path: str
    excel_name: str
    zip_dir: str
    zip_file: str
    email_list: List[str]

    data: dict = None

    def __post_init__(self) -> None:
        self.attachment = self.zip_file
        self.data = {
            'credentials': Credentials(usr=self.usr, psw=self.psw),
            'process': Process(name=self.process_name, path=self.process_path),
            'excel': ExcelInfo(path=self.excel_path, name=self.excel_name),
            'archive_info': ArchiveInfo(zip_dir=self.zip_dir, zip_name=self.zip_file),
            'email_info': EmailInfo(email_list=self.email_list)
        }
