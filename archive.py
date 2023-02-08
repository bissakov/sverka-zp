import os
import shutil
from data_structures import ArchiveInfo


class Archive:
    def __init__(self, info: ArchiveInfo) -> None:
        self.zip = info

    def run(self) -> None:
        if os.path.isfile(self.zip.zip_name):
            os.unlink(self.zip.zip_name)
        shutil.make_archive(self.zip.file_name, 'zip', self.zip.zip_dir)
