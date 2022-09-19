import os
from os.path import join


class utils:

    def __init__(self):
        ...

    def recursiveReadDir(self, directory):
        _lst_txt = lst_txt = []
        for root, directories, filenames in os.walk(directory):
            for directory in directories:
                join(root, directory)
            for filename in filenames:
                if filename[-3:].upper() == 'PDF':
                    lst_txt.append(join(root, filename))
        return lst_txt