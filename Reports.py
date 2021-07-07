from pptx import Presentation
from PIL import Image
import os.path
import sys

from pptx.util import Inches


class Reports:
    def __init__(self, path='C:/'):
        self.__prefix = {}
        self.__path = path if os.path.exists(path) else None
        self.__walk = [item for item in os.walk(path)]
        self.__numbers_row = len(self.__walk[0][1])
        self.__numbers_columns = len(self.__walk[1][1])
        self.__numbers_slide = 0
        self.__images = {}

    def __stop_program(self):
        if self.__path is None:
            print("a;ldf")
            sys.exit(-1)

    def create_slide(self):
        self.__stop_program()

        self.__template()

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        s = slide.shapes.add_picture('D:\\test\\сетка1\\q1\\Pn_field_x=2.png', Inches(1.0), Inches(0.5))
        prs.save('test.pptx')

    def add_prefix(self, prefix):
        for (key, value) in prefix.items():
            self.__prefix[key] = value
        self.__numbers_slide = len(self.__prefix)

    def __template(self):
        for i, grid in enumerate(os.listdir(self.__path)):
            for q in os.listdir(f'{self.__path}\\{grid}'):
                for img in os.listdir(f'{self.__path}\\{grid}\\{q}'):
                    self.__images[f'{img.split("_")[0]}_grid{i+1}_{q}'] = f'{self.__path}\\{grid}\\{q}\\{img}'

        for item, val in self.__images.items():
            # with Image.open(val) as img:
            #     img = img.resize((60, 51))
            #     img.save(val)
            print(f'{item}: {val}')
