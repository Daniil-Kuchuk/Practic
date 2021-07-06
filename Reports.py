from pptx import Presentation
import os.path
import sys


class Reports:
    def __init__(self, path='C:/'):
        self.__prefix = {}
        self.__path = path if os.path.exists(path) else None
        self.__walk = [item for item in os.walk(path)]
        self.__numbers_row = len(self.__walk[0][1])
        self.__numbers_columns = len(self.__walk[1][1])
        self.__numbers_slide = 0

    def __stop_program(self):
        if self.__path is None:
            print("a;ldf")
            sys.exit(-1)

    def create_slide(self):
        self.__stop_program()

        self.__template()

    def add_prefix(self, prefix):
        for (key, value) in prefix.items():
            self.__prefix[key] = value
        self.__numbers_slide = len(self.__prefix)

    def __template(self):
        print(self.__walk)
        slide = {}
        for prefix in self.__walk[2][2]:
            prefix = prefix.split('_')[0]
            slide[self.__prefix[prefix]] = {}

        for grid in slide.values():
            for a in self.__walk[0][1]:
                grid[a] = {}

        for item in slide.keys():
            for grid in os.listdir(self.__path):
                for q in os.listdir(self.__path + grid):
                    pass


        print(slide)