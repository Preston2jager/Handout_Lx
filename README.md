# Handout_Lx

A quick step by step handout generator running with LaTex.

## Requirements

### For users just want handouts in default layout:
The dist is packed with all required libs, the only requirement is XeLatex. A full installation of Texlive is recommanded.

### For users who know what they are doing:
##### To do...
For compiling your own, the following packages are required:
- python-doc
- fitz

Use the following line to compile:
~~~ bash
pyinstaller --icon=logo.ico --onefile --name=HandoutGen main.py
~~~

## Usage

### Deployment

Clone the repo to local, rename folders as you like. The exe file reads the basic.yaml in subfolders to identify different Handouts. It requires a folder structure as below:
~~~
Handout_folder
├── Project_folder_1
├── Project_folder_2
├── Project_folder_3
├── Project_folder_4
|   ....
└── HandoutGen.exe
~~~

### Handout Project Setup

Project_folder_X
├── Latex             <-- The Latex folder is managed automatically.
|   ├── Resources
|   |   ├── figure_generated
|   |   ├── logo
|   |   |   └── logo.png   <-- Replace with your own. Keep the same file name.
|   |   └── page
|   ├── Handout_Lx.cls
|   └── main.tex        
├── basic.yaml        <-- Project information, related to headers and footers.
└── content.docx      <-- All your handout content, 

Refer the sample_handout.pdf for handout syntax.

### How it work

The program the fetch everything from content.docx and basic.yaml and convert them into .tex files. Then it will call Xelatex to complie a combined PDF.

- Images are fetched from docx file and handled automatically.
    - You don't need to re-size the image in docx file.
    - Images height in PDF are limited to 1/3 of the page height.
    - By default the docx file will not compress any image.
    - Support multiple images for a single step.
- Handout steps and pages numbers are handled automatically.
- Support sub-title line.
