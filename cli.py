import os
import argparse

__version__ = 20170404


def get_parser():
    """Parse user selected / default attributes."""
    parser = argparse.ArgumentParser()
    parser.formatter_class = CustomHelpFormatter
    parser.description = """
    Make CFD presentation for comparison up to three variants.

    Script has to be launched in the FOLDER with VARIANTS.
    Example:
        YETI:  /ST/SkodaAuto/AEROAKUSTIKA/PRJ/SK326-0/
        RAPID: /ST/SkodaAuto/AEROAKUSTIKA/PRJ/SK370-3/STACIONARNI-VYPOCET/"""

    # VERSION
    parser.add_argument('--version',
                        action='version',
                        version='%(prog)s: v{}'.format(__version__))

    parser.add_argument(dest='variants',
                        metavar='VARIANT',
                        type=str,
                        nargs='*',
                        help="full FOLDER NAME of variant(s)")

    parser.add_argument('-o', '--output',
                        dest='output_pptx',
                        metavar='output_pptx',
                        type=str,
                        default=os.path.join(os.path.realpath(os.path.curdir), 'OUTPUT.pptx'),
                        help='Specify full path with name for output presentation: /path/to/output.pptx\n')

    parser.add_argument('-i', '--input',
                        dest='input_pptx',
                        metavar='input_pptx',
                        type=str,
                        default=os.path.join(
                            os.path.dirname(__file__), 'TEMPLATES', 'SABLONA-RAPID-AEROAKUSTIKA.pptx'),
                        help='Specify full path for custom input presentation\n')

    parser.add_argument('-c', '--config',
                        dest='cfg_file',
                        metavar='cfg_file',
                        type=str,
                        default=os.path.join(os.path.dirname(__file__), 'slides.cfg'),
                        help='Optional user slides config file\n')


    parser.add_argument('--show_placeholders',
                        dest='show_placeholders',
                        action='store_true',
                        help='Generate PPTX that shows IDs and Names of placeholders\n')

    # In main functions, use:
    # parser = get_parser()
    # args = parser.parse_args()
    return parser


class CustomHelpFormatter(argparse.ArgumentDefaultsHelpFormatter, argparse.RawTextHelpFormatter):
    """ArgParse custom formatter that has LONGER LINES and RAW DescriptionHelp formatting."""

    def __init__(self, prog):
        super(CustomHelpFormatter, self).__init__(prog, max_help_position=60, width=80)
