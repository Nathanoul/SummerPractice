'''
Copies files from one folder to another
'''
from shutil import copy2
from os import path, remove, listdir, rmdir, mkdir


def Copy2Folder(frompath: str, topath: str):
    # copy every directory in the folder frompath to the folder topath
    for filename in listdir(frompath):
        copy2(path.join(frompath, filename), path.join(topath, filename))
