# DocMergePDF-Merge-All-Your-PDFs-Docs-PPTs-and-PPTXs-to-a-Single-PDF
The repo is a Python project that allows users to merge multiple PDFs, Word documents, and PowerPoint presentations into a single PDF file. Users can simply move the files they want to merge to the current directory, optionally rename them, and then run the program. 

# PDF, DOC, PPT and PPTX merger

This is a Python script that merges all PDFs, DOCs, PPTs or PPTXs in a directory into a single PDF file.

## Requirements

Before running the script, make sure you have installed the required packages by running the following command:

``pip install -r requirements.txt``

## Code

The Python script can be found in the file `merge.py`.

## Usage

1. Move all the PDFs, DOCs, PPTs to the current directory
2. If the files are numbered 1-,2-,3-.. etc, enter 'y'when prompted to rename the files, else enter 'n'
3. The files needs to be renamed lexicographically to maintain the order
4. un the program

Run the script by executing the following command in the terminal:

``python merge.py``

The output PDF file will be saved in the `output` directory. The input files will be deleted after the merge process is completed.

## License

This project is licensed under the MIT License. See the LICENSE file for more information.
