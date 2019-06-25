# README

PDF stamper binary using Pymupdf library with GUI for linux.

Within the folder you will find the executable `Pymupdf-batch-incremental-GUI-V.0.5` run that to stamp or multistamp:

**stamp** - Stamp all pages of one pdf with corresponding stamp pdf file. Enter full paths to folder of to be stamped files, folder containing stamp files and output folder.
        Filenames must start with 3 characters corresponding to the stamp filename you want to stamp with.
        Example Filenames:
                    To be stamped file: 001-abc.pdf
                    Stamp file: 001-xyz.pdf

**multistamp** - Stamp all pages of all files with one stamp.pdf. Enter full path to folder containing to be stamped files, full path to stamp FILE and output folder.
             The stamp file need not have follow the naming convention as described in stamp.
