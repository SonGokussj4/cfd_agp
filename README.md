# cfd_agp

Automatic Generated Presentation for CFD results using python-pptx


### Examples:
$ cfd_agp S100-BASIC-MIRROR-PR2/
    outputs OUTPUT.pptx in current directory using default template

$ cfd_agp S100-BASIC-MIRROR-PR2/ -i ~/my-template.pptx -o ~/output/here/my-output-name.pptx
    outputs my-output-name.pptx in ~/output/here using my-tamplate.pptx

$ cfd_agp S100-BASIC-MIRROR-PR2/ cfd_agp S200-BASIC-MIRROR-PR2/ -c ~/MyAGP/my_config.cfg
    outputs OUTPUT.pptx in current directory (comparing 2 variants) using user config file my_config.cfd
    placed in ~/MyAGP folder


### cfd_agp --help
```
    usage: cfd_agp [-h] [--version] [-o output_pptx] [-i input_pptx] [-c cfg_file]
                 [--show_placeholders]
                 [VARIANT [VARIANT ...]]

      Make CFD presentation for comparison up to three variants.

      Script has to be launched in the FOLDER with VARIANTS.
      Example:
          YETI:  /ST/SkodaAuto/AEROAKUSTIKA/PRJ/SK326-0/
          RAPID: /ST/SkodaAuto/AEROAKUSTIKA/PRJ/SK370-3/STACIONARNI-VYPOCET/

  positional arguments:
    VARIANT                               full FOLDER NAME of variant(s) (default: None)

  optional arguments:
    -h, --help                            show this help message and exit
    --version                             show program's version number and exit
    -o output_pptx, --output output_pptx  Specify full path with name for output presentation: /path/to/output.pptx
                                           (default: /ST/SkodaAuto/AEROAKUSTIKA/PRJ/SK370-3/STACIONARNI-VYPOCET/OUTPUT.pptx)
    -i input_pptx, --input input_pptx     Specify full path for custom input presentation
                                           (default: /expSW/SOFTWARE/skripty/cfd_agp/TEMPLATES/SABLONA-RAPID-AEROAKUSTIKA.pptx)
    -c cfg_file, --config cfg_file        Optional user slides config file
                                           (default: /expSW/SOFTWARE/skripty/cfd_agp/slides.cfg)
    --show_placeholders                   Generate PPTX that shows IDs and Names of placeholders
                                           (default: False)
```

