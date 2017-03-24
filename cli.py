import os
import argparse

__version__ = 20170324


def get_parser():
    """Parse user selected / default attributes."""
    parser = argparse.ArgumentParser()
    parser.formatter_class = CustomHelpFormatter
    parser.description = """
    Make CFD presentation for comparison up to three variants.

    Script HAS TO BE LAUNCHED in the folder with VARIANTS.
    For example in: /ST/SkodaAuto/AEROAKUSTIKA/PRJ/SK326-0/"""

    # VERSION
    parser.add_argument('--version',
                        action='version',
                        version='%(prog)s: v{}'.format(__version__))

    parser.add_argument(dest='variants',
                        metavar='VARIANT',
                        nargs='+',
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
                        default=os.path.join(os.path.dirname(__file__), '_FILES', 'TEST.pptx'),
                        help='Specify full path for input presentation\n')

    return parser.parse_args()  # In main functions, use: `args = get_parser()`


class CustomHelpFormatter(argparse.ArgumentDefaultsHelpFormatter, argparse.RawTextHelpFormatter):
    """ArgParse custom formatter that has LONGER LINES and RAW DescriptionHelp formatting."""

    def __init__(self, prog):
        super(CustomHelpFormatter, self).__init__(prog, max_help_position=80, width=80)
