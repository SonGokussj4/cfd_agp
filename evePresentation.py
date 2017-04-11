"""Evektor library for all things"""
import os
import sys
import pptx
import colorlog
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from collections import namedtuple

formatter = colorlog.ColoredFormatter(
    '%(log_color)s%(levelname)-8s%(reset)s %(message_log_color)s%(message)s',
    datefmt=None,
    reset=True,
    log_colors={
        'CRITICAL': 'bold_red',
        'ERROR': 'red',
        'WARNING': 'yellow',
        'INFO': 'green',
        'DEBUG': 'cyan',
    },
    secondary_log_colors={
        'message': {
            'CRITICAL': 'bold_red',
            'ERROR': 'red',
            'WARNING': 'yellow',
            'INFO': 'white',
            'DEBUG': 'cyan',
        }
    },
    # style='%',
)

# Initialize LOGGER
handler = colorlog.StreamHandler()
handler.setFormatter(formatter)
logger = colorlog.getLogger(__name__)
logger.addHandler(handler)
logger.setLevel('INFO')


class Presentation:

    def __init__(self, src_prs_path: str):
        logger.info("Loading template: {}".format(src_prs_path))
        self.prs = pptx.Presentation(src_prs_path)
        self.slides = []
        self.variants = []
        self.conf = None
        self.one_image_slides = [2, 3]
        self.two_images_slides = [4, 5, 6]

    def load_config(self, config_file):
        import configparser
        conf = configparser.ConfigParser()
        conf.read(config_file)
        self.conf = conf

    def add_variants(self, variants: list):
        if len(variants) not in range(1, 4):
            raise Exception("Only 1-3 variants possible for now...\n")

        for var in variants:
            # Create a NamedTuple
            variant = namedtuple('variant', ('name', 'fullpath', 'num', 'exists'))
            # Fill the NamedTuple
            variant.name = var.rstrip('/')
            variant.fullpath = os.path.abspath(var)
            variant.folder = '/'.join(variant.fullpath.split('/')[0:-1])
            variant.num = variant.name.split('-')[0]
            variant.exists = os.path.isdir(os.path.abspath(var))

            # Check if variant (folder) exists, if true, add to list of Presentation.variants
            if variant.exists:
                self.variants.append(variant)
            else:
                logger.critical(
                    "Project: {} was not found in: {} \nAre you in the RIGH folder??? "
                    "And please check entered variant names. ".format(variant.name, os.getcwd()))
                sys.exit()

    def get_num_of_slides(self) -> int:
        return len(self.prs.slides)

    def process_slides(self):
        author = self.conf.get('User Settings', 'author')

        for section in self.conf.sections():  # nebo: self.conf.sections()[1:]
            if not self.conf.has_option(section, 'layout'):
                continue

            # Convert argparse list of tuples to dictionary of key: val
            conf = dict(self.conf.items(section))

            # Get variables from config file of actual section
            title = conf.get('title', "{} TITLE MISSING".format(section))
            layout_num = int(conf.get('layout', 2))
            fringebar = conf.get('fringebar', None)
            images = conf.get('images', []).replace(' ', '') .split(',')

            # Add slide as object
            slide_layout = self.prs.slide_layouts[layout_num]
            pptx_slide = self.prs.slides.add_slide(slide_layout)

            self.slides.append(pptx_slide)
            logger.info("Adding slide {}".format(self.get_num_of_slides()))
            logger.debug("Variables: \nTitle: {} \nLayout: {} \nFringebar: {} \nImages: {}".format(
                title, layout_num, fringebar, images))

            # Slide object: add title, fringebar and images
            slide = Slide(self, pptx_slide, layout_num)
            slide.set_title(title)
            slide.set_author(author)
            slide.add_fringebar(fringebar)
            slide.add_images(images)

    def save_presentation(self, output_pres_path):
        if '/' in output_pres_path:
            output_path = output_pres_path
            self.prs.save(output_path)

        else:
            cur_dir = os.path.realpath(os.path.curdir)
            output_path = os.path.join(cur_dir, output_pres_path)
            self.prs.save(output_path)

        logger.info("Presentation saved to: {}".format(output_path))

    def output_placeholders_pptx(self, output_pres_path: str):
        for master_slide in self.prs.slide_masters:
            for lay_id, layout in enumerate(master_slide.slide_layouts):
                # if lay_id == 0:
                #     continue
                slide = self.prs.slides.add_slide(layout)

                for ph in slide.placeholders:

                    # TEXT modification
                    p = ph.text_frame.paragraphs[0]
                    p.alignment = PP_ALIGN.CENTER
                    run = p.add_run()
                    run.font.size = Pt(11)
                    run.font.bold = True

                    # COLOR modification
                    ph.fill.solid()
                    ph.fill.fore_color.rgb = RGBColor(241, 241, 241)

                    if ph.placeholder_format.type == 1:  # TITLE
                        run.text = "Id[{}]: {} - Layout[{}]: [{}]".format(
                            ph.placeholder_format.idx, ph.name, lay_id, layout.name)
                    else:
                        run.text = "Id[{}]: {}".format(ph.placeholder_format.idx, ph.name)

        self.prs.save(output_pres_path)

    def plot_gradients(self):
        import csv
        import numpy as np
        import matplotlib.pyplot as plt
        import pandas as pd

        for sec in self.conf.options('Plots'):  # each plot
            sec_name = self.conf.get('Plots', sec)

            fig = plt.figure()
            # fig.clf()
            plt.style.use('seaborn-notebook')
            plt.grid(True)
            plt.title(sec_name)

            # Gets a len(self.variants) num of curves to 1 plot
            for variant in self.variants:

                # Load x, y data of each variant
                datafile = os.path.join(variant.fullpath, 'PICTURES', sec_name)
                data = pd.read_csv(datafile, index_col=False, skiprows=5, comment='*', header=None, usecols=[0, 1])
                x = data[0].values
                y = data[1].values
                plt.plot(x, y)
                # print("x:", x)
                # print("y:", y)

            fig.savefig('{}.png'.format(sec_name), dpi=1200, format='png')
        plt.show()




class Slide():

    def __init__(self, pres, slide, layout_num):
        self.slide = slide  # slide knows about pptx.slide object
        self.pres = pres  # slide knows about presentation
        self.variants = pres.variants
        self.layout_num = layout_num
        self.slide_num = pres.get_num_of_slides()

    def set_title(self, title: str):
        self.slide.shapes.title.text = title

    def set_author(self, author: str):
        self.slide.placeholders[20].text = author

    def add_images(self, images):

        # Should be 1 image but more were specified in config file
        if self.layout_num in self.pres.one_image_slides and len(images) > 1:
            logger.critical("You've specified [{} images] for layout[{}]. Should be [{} image]. "
                            "Fix the config file.".format(len(images), self.layout_num, 1))
            sys.exit()

        # Should be more images but less was specivied in config file
        elif self.layout_num in self.pres.two_images_slides and len(images) < 2:
            logger.critical("You've specified only [{} image] for layout[{}]. Should be [{} images]. "
                            "Fix the config file.".format(len(images), self.layout_num, 2))
            sys.exit()

        for idx, variant in enumerate(self.variants):

            # TEXT
            self.slide.placeholders[17 + idx].text = variant.num  # Variant number

            # IMAGES
            # 1st image is in all slide layouts
            img1_path = os.path.join(variant.fullpath, 'PICTURES', images[0])
            if os.path.isfile(img1_path):
                self.slide.placeholders[11 + idx].insert_picture(img1_path)
            else:
                logger.error("Image: {} does not exist in \n         {}".format(
                    os.path.basename(img1_path), os.path.dirname(img1_path)))

            # 2nd additional image is only in layout 2, 4, 5
            if self.layout_num in self.pres.two_images_slides:
                img2_path = os.path.join(variant.fullpath, 'PICTURES', images[1])

                if os.path.isfile(img2_path):
                    self.slide.placeholders[14 + idx].insert_picture(img2_path)
                else:
                    logger.error("Image: {} does not exist in \n         {}".format(
                        os.path.basename(img2_path), os.path.dirname(img2_path)))

    def add_fringebar(self, fringebar: str):
        if self.layout_num in [2, 4] and fringebar is None:
            logger.critical("Slide [{}] with Layout[{}] has to have fringebar but none was specified "
                            "in config file. Aborting script...".format(self.slide_num, self.layout_num))
            sys.exit()

        elif self.layout_num not in [2, 4] and fringebar is not None:
            logger.error("Slide [{}] with Layout[{}] should not have FRINGEBAR assigned. "
                         "Please see the config file and check this slide.".format(self.slide_num, self.layout_num))
            return None

        if fringebar is not None:
            fringebar_path = os.path.join(self.variants[0].fullpath, 'PICTURES', fringebar)
            if os.path.isfile(fringebar_path):
                self.slide.placeholders[10].insert_picture(fringebar_path)
            else:
                logger.warning("Fringebar: {} does not exist in \n         {}".format(
                    os.path.basename(fringebar_path), os.path.dirname(fringebar_path)))
