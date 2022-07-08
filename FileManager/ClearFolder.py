'''
Recursive function clearing folder from files and
calling itself when clearing folder within folder
'''
from shutil import copy2
from os import path, remove, listdir, rmdir, mkdir

def ClearFolder(folderpath: str):
    # check every directory inside chosen folder
    for dirname in listdir(folderpath):
        if path.isdir(path.join(folderpath, dirname)):
            # clear and remove directory if it's a folder
            ClearFolder(path.join(folderpath, dirname))
            rmdir(path.join(folderpath, dirname))
        else:
            # remove directory if it's not a folder
            remove(path.join(folderpath, dirname))
