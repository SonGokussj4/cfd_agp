"""Evektor library for all things"""
import os
import pptx
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


class Presentation:

    def __init__(self, src_prs_path):
        print("[ INFO    ] Loading presentation teplate: {}".format(src_prs_path))
        self.prs = pptx.Presentation(src_prs_path)
        self.variants = []

    def add_variants(self, variants):
        if len(variants) not in range(1, 4):
            raise Exception("Only 1 or 3 variants possible for now...\n")

        self.variants = variants

    def get_num_of_slides(self) -> int:
        return len(self.prs.slides)

    def save_presentation(self, output_pres_path):
        if '/' in output_pres_path:
            output_path = output_pres_path
            self.prs.save(output_path)

        else:
            cur_dir = os.path.realpath(os.path.curdir)
            output_path = os.path.join(cur_dir, output_pres_path)
            self.prs.save(output_path)

        print("[ INFO    ] Presentation saved to: {}".format(output_path))

    def output_placeholders(self, output_pres_path):
        for master_slide in self.prs.slide_masters:
            for lay_id, layout in enumerate(master_slide.slide_layouts):
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
                    # elif ph.placeholder_format.type == 18:  # PICTURE
                    #     run.text = "Id[{}]: {}".format(ph.placeholder_format.idx, ph.name)
                    else:
                        run.text = "Id[{}]: {}".format(ph.placeholder_format.idx, ph.name)

        self.prs.save(output_pres_path)

    def add_slide(self, title, layout_num):
        slide_layout = self.prs.slide_masters[1].slide_layouts[layout_num]
        slide = self.prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = title
        return Slide(self, slide, layout_num)


class Slide():

    def __init__(self, pres, slide, layout_num):
        self.slide = slide
        self.pres = pres
        self.variants = pres.variants
        self.layout_num = layout_num
        self.slide_num = pres.get_num_of_slides()

    def add_images(self, *images):

        PROJECT_PATH = os.path.realpath(os.path.curdir)

        for idx, variant in enumerate(self.variants):
            variant_num = variant.split('-')[0]

            # TEXT
            self.slide.placeholders[17 + idx].text = variant_num

            # IMAGES
            if self.layout_num in [1, 3]:
                img_path = os.path.join(PROJECT_PATH, variant, 'PICTURES', images[0])
                self.slide.placeholders[11 + idx].insert_picture(img_path)

            elif self.layout_num in [2, 4, 5]:
                img1_path = os.path.join(PROJECT_PATH, variant, 'PICTURES', images[0])
                img2_path = os.path.join(PROJECT_PATH, variant, 'PICTURES', images[1])
                self.slide.placeholders[11 + idx].insert_picture(img1_path)
                self.slide.placeholders[14 + idx].insert_picture(img2_path)

    def add_fringebar(self, fringebar):
        PROJECT_PATH = os.path.realpath(os.path.curdir)

        if self.layout_num not in [1, 2]:
            raise Exception("Layout num: {} has NO FRINGEBAR".format(self.layout_num))
        try:
            fringebar_path = os.path.join(PROJECT_PATH, self.variants[0], 'PICTURES', fringebar)
            self.slide.placeholders[10].insert_picture(fringebar_path)
        except FileNotFoundError as e:
            print("[ WARNING ] Fringebar not found...", e)
