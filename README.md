# BAnAnASpy
BAtch 'n Automated 'n Aligner Sequences python program


## Setup:
Before run the program for the first time, you need to setup some variables located at the start of 'main()' method.
No need external programs, you just need some python libraries

### Libraries to install:
    'biopython' 'xlsxwriter'
    
    
Linux: pip install 'package' or pip3 install 'package'
Windows: python or py -m pip install 'package'

sudo apt-get install python-tk for Python2.X
sudo apt-get install python3-tk for Python3.X

### Things to know about filename of SEQ before sequencing:
*Use '-'(6) to separate each part name
*Fist part must be the patient id
*Second part must be the gene
*The other parts order is free
*The primer part must finish by 'F' or 'S' for forward sequences
*The primer part must finish by 'R' for reverse sequences
*For multiplex or partial exon names, separate each number by '_'(8) ou '.' ou '/'

    Before sequencing:
    Exp filenames: 1201645-ABCB4-8F, 1123546-ABCB11-DPN-8_9R
    After sequencing:
    Exp files Ab1: A01_1201645-ABCB4-8F.ab1, H02_1123546-ABCB11-DPN-8_9R.ab1

Coding with Python3.7, Biopython1.77, XlsxWriter1.3.3
