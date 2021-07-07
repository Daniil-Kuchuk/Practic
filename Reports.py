from pptx import Presentation
from PIL import Image
import os.path
import sys

from pptx.util import Inches, Pt


class Reports:
    def __init__(self, path='C:/'):
        self.__prefix = {}
        self.__path = path if os.path.exists(path) else None
        self.__walk = [item for item in os.walk(path)]
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
        numbers_row = len(os.listdir(self.__path))
        numbers_columns = len(os.listdir(os.chdir(self.__path + '\\' + os.listdir(self.__path)[0])))

        for key, value in self.__prefix.items():
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = f'{value}'
            for grid in range(numbers_row):
                txBox = slide.shapes.add_textbox(Pt(30.0), Pt(25.5), Pt(23.0), Pt(22.5))
                txBox.text_frame._bodyPr.attrib.update({'vert': 'vert'})
                tf = txBox.text_frame
                p = tf.add_paragraph()
                p.text = f'сетка {grid+1}'
                for col in range(numbers_columns):
                    q_txt_box = slide.shapes.add_textbox(Pt(12.0), Pt(10.5), Pt(3.0), Pt(2.5))
                    tf = q_txt_box.text_frame.add_paragraph()
                    tf.text = f'q{col+1}'
                    slide.shapes.add_picture(self.__images[f'{key}_grid{grid+1}_q{col+1}'], Inches(5.0), Inches(6.5))

        prs.save('C:\\Users\\kuchu\\PycharmProjects\\Practic\\test.pptx')

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
