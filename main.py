#!/usr/bin/python3
import os
import sys
# sys.path.append(os.path.join(os.path.dirname(__file__), 'libs'))
sys.path.append(os.path.join(os.path.split(os.path.abspath(os.path.realpath(sys.argv[0])))[0], 'libs'))
import evePresentation
import cli
import colorlog
import better_exceptions

# Initialize LOGGER
handler = colorlog.StreamHandler()
handler.setFormatter(evePresentation.formatter)
logger = colorlog.getLogger(__name__)
logger.addHandler(handler)
logger.setLevel('DEBUG')


def main():
    """Main function"""
    # Get parser parameters from CLI
    parser = cli.get_parser()
    args = parser.parse_args()

    # Load Presentation Template
    logger.info("Starting...")
    pr = evePresentation.Presentation(src_prs_path=args.input_pptx)

    # Arg option: --show_placeholders
    if args.show_placeholders:
        pres_name = 'PLACEHOLDERS.pptx'
        pr.output_placeholders_pptx(pres_name)
        logger.info("Created {}".format(os.path.join(os.getcwd(), pres_name)))
        sys.exit()

    # MAIN SCRIPT
    if not args.variants:
        logger.error("You have to specify variants (folder names...)\n")
        parser.print_help()
        sys.exit()

    # Load config file for slides
    pr.load_config(args.cfg_file)

    # Add user selected variants
    pr.add_variants(args.variants)

    # Process slides
    pr.process_slides()

    # Save Presentation
    pr.save_presentation(args.output_pptx)


if __name__ == '__main__':
    main()
