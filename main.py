#!/usr/bin/env python3
import os
import sys
# sys.path.append(os.path.join(os.path.dirname(__file__), 'libs'))
sys.path.append(os.path.join(os.path.split(os.path.abspath(os.path.realpath(sys.argv[0])))[0], 'libs'))
from evePresentation import Presentation
import cli


def main():
    """
    Main function

    # layouts:
    #   - [0] ... Title, 2 Content  (X)
    #   - [1] ... Fringebar, 3 normal pictures  (1 fringebar.jpeg + 1 car.jpeg)
    #   - [2] ... Fringebar, 6 normal pictures  (1 fringebar.jpeg + 2 car.jpeg)
    #   - [3] ... 3 normal pictures  (1 car.jpeg)
    #   - [4] ... 6 normal pictures  (2 car.jpeg)
    #   - [5] ... 6 wider pictures  (2 car.jpeg)
    """
    # Get parser parameters from CLI
    args = cli.get_parser()

    # Load Presentation Template
    pr = Presentation(src_prs_path=args.input_pptx)

    # Add Variants
    pr.add_variants(args.variants)
    # pr.output_placeholders('PLACEHOLDERS.pptx')

    # Add Slides
    # SLIDE 3: TLAKOVÝ GRADIENT
    slide = pr.add_slide(title="TLAKOVÝ GRADIENT", layout_num=2)
    # slide.add_fringebar('GRADP_fringebar.jpeg')
    slide.add_images('GRADP1.jpeg',
                     'GRADP2.jpeg')

    # SLIDE 4: TLAKOVÝ GRADIENT REZ Z=0.81
    slide = pr.add_slide("TLAKOVÝ GRADIENT REZ Z=0.81", 1)
    # slide.add_fringebar('GRADP-REZ_fringebar.jpeg')
    slide.add_images('GRADP-REZ-Z081.jpeg')

    # SLIDE 5: VORTICITY X
    slide = pr.add_slide("VORTICITY X", 2)
    # slide.add_fringebar('VORTICITY_X_a_fringebar.jpeg')
    slide.add_images('VORTICITY_X_a2.jpeg',
                     'VORTICITY_X_a1.jpeg')

    # SLIDE 6: VORTICITY X
    slide = pr.add_slide("VORTICITY X", 2)
    # slide.add_fringebar('VORTICITY_X_b_fringebar.jpeg')
    slide.add_images('VORTICITY_X_b1.jpeg',
                     'VORTICITY_X_b2.jpeg')

    # SLIDE 7: VORTICITY X
    slide = pr.add_slide("VORTICITY X", 2)
    # slide.add_fringebar('VORTICITY_X_c_fringebar.jpeg')
    slide.add_images('VORTICITY_X_c2.jpeg',
                     'VORTICITY_X_c1.jpeg')

    # SLIDE 8: KINETICKÁ ENERGIE MEZNÍ VRSTVY
    slide = pr.add_slide("KINETICKÁ ENERGIE MEZNÍ VRSTVY", 2)
    # slide.add_fringebar('kineticka_energie_fringebar.jpeg')
    slide.add_images('kineticka_energie3.jpeg',
                     'kineticka_energie1.jpeg')

    # SLIDE 9: KINETICKÁ ENERGIE MEZNÍ VRSTVY
    slide = pr.add_slide("KINETICKÁ ENERGIE MEZNÍ VRSTVY", 2)
    # slide.add_fringebar('kineticka_energie_fringebar.jpeg')
    slide.add_images('kineticka_energie4.jpeg',
                     'kineticka_energie2.jpeg')

    # SLIDE 10: RYCHLOST MEZNÍ VRSTVY
    slide = pr.add_slide("RYCHLOST MEZNÍ VRSTVY", 2)
    # slide.add_fringebar('kineticka_energie_fringebar.jpeg')
    slide.add_images('RYCHLOST_MEZNI_VRSTVY1.jpeg',
                     'RYCHLOST_MEZNI_VRSTVY2.jpeg')

    # SLIDE 11: KINETICKÁ ENERGIE MEZNÍ VRSTVY
    slide = pr.add_slide("RYCHLOST MEZNÍ VRSTVY", 2)
    # slide.add_fringebar('kineticka_energie_fringebar.jpeg')
    slide.add_images('RYCHLOST_MEZNI_VRSTVY4.jpeg',
                     'RYCHLOST_MEZNI_VRSTVY3.jpeg')

    # SLIDE 12: KINETICKÁ ENERGIE MEZNÍ VRSTVY
    slide = pr.add_slide("KINETICKÁ ENERGIE MEZNÍ VRSTVY", 2)
    slide.add_images('iso_plocha_0ms1.jpeg', 'iso_plocha_0ms2.jpeg')

    # SLIDE 13: KINETICKÁ ENERGIE MEZNÍ VRSTVY
    slide = pr.add_slide("ISO PLOCHA – RYCHLOST PROUDU 0 m/s V OSE X", 5)
    slide.add_images('iso_plocha_0ms3.jpeg',
                     'iso_plocha_0ms4.jpeg')

    # Save Presentation
    pr.save_presentation(args.output_pptx)


if __name__ == '__main__':
    main()
