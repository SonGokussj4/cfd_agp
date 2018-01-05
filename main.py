#!/usr/bin/python3
import os
import sys
# sys.path.append(os.path.join(os.path.dirname(__file__), 'libs'))
script_dir = os.path.dirname(os.path.realpath(sys.argv[0]))
sys.path.append(os.path.join(script_dir, 'libs'))
import evePresentation
import cli
import colorlog
import better_exceptions
import configparser
import os.path

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

    # Check if user entered variants
    if args.variants:
        if os.path.isfile(args.variants[0]):
            config = configparser.ConfigParser()
            config.read(args.variants[0])
            args.variants = [config[section] for section in config.sections()]
            args.cfg_file = config.get('DEFAULT', 'cfg_file', fallback=args.cfg_file)
            args.input_pptx = config.get('DEFAULT', 'input_pptx', fallback=args.input_pptx)
            args.output_pptx = config.get('DEFAULT', 'output_pptx', fallback=args.output_pptx)

    # Load Presentation Template
    logger.info("Starting...")
    pr = evePresentation.Presentation(src_prs_path=args.input_pptx)

    # Arg option: --show_placeholders
    if args.show_placeholders:
        pres_name = 'PLACEHOLDERS.pptx'
        pr.output_placeholders_pptx(pres_name)
        logger.info("Created {}".format(os.path.join(os.getcwd(), pres_name)))
        exit()

    # Arg option: --readme
    if args.readme:
        import subprocess
        readme_file = os.path.join(script_dir, 'README.md')
        subprocess.call(['sublime', readme_file])
        exit()
    # Arg option: -g --gradients
    if args.gradients:
        pr.gradients_from_file(args.gradients)
        exit()

    # Check if user entered variants
    if not args.variants:
        logger.error("You have to specify variants (folder names...)\n")
        parser.print_help()
        sys.exit()

    # Load config file for section [Slide \d]
    pr.load_config(args.cfg_file)

    # Add user selected variants
    pr.add_variants(args.variants)

    # Arg options: --plots
    if args.plots:
        pr.plot_gradients()
        exit()

    # Process slides
    pr.process_slides()

    # Save Presentation
    pr.save_presentation(args.output_pptx)


if __name__ == '__main__':
    main()
