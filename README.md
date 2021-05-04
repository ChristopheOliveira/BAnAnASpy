![alt tag](https://user-images.githubusercontent.com/80535244/112607815-e4ddcd80-8e19-11eb-9520-b5e1f27aeb09.png) 
# BAnAnASpy
BAtch 'n Automated 'n Aligner Sequences python program


## Setup:
Before run the program for the first time, you need to setup some variables located at the start of 'main()' method.\
No need of additional programs, you just need some python libraries

&nbsp;
### Libraries to install:
    'biopython' 'xlsxwriter'

>Linux: `pip install 'package'` or `pip3 install 'package'`\
>Windows: `python pip install 'package'` or `py -m pip install 'package'`

    'tkinter' if necessary
>Python2.X: `sudo apt-get install python-tk`\
>Python3.X: `sudo apt-get install python3-tk`

&nbsp;
### Things to know about filename of SEQ before sequencing:
* Use dash '-' to separate each part name
* First part must be the patient id
* Second part must be the gene
* The order of other parts is free
* The primer part must finish by 'F' or 'S' for forward sequences
* The primer part must finish by 'R' for reverse sequences
* For multiplex or partial exon names, separate each number by underscore '_' or dot '.' or slash '/'

Examples:

    filenames before sequencing: 1201645-ABCB4-8F, 1123546-ABCB11-DPN-8_9R
    ab1 files: A01_1201645-ABCB4-8F.ab1, H02_1123546-ABCB11-DPN-8_9R.ab1
You can also use the program `FormatSeqName` for renaming your .ab1 files to the format supported by BAnAnASpy.
<br/>
<br/>
<br/>
Coding with Python3.7, Biopython1.77, XlsxWriter1.3.3
