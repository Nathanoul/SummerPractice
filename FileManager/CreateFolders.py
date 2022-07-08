'''
Creates folders if they are not exist
'''
from shutil import copy2
from os import path, remove, listdir, rmdir, mkdir


def CreateFolders(pathnames: list):
    for pathname in pathnames:
        Path = pathname
        if not path.exists(Path):
            mkdir(Path)
