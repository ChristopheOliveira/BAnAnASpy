#!/usr/bin/python
# -*- coding:utf-8 -*-

# Script by Christophe Oliveira
# Contact : christophe.oliveira@aphp.fr, christophe.oliveira33@gmail.com

# FormatSeqName_v0.3
# Program for renaming .ab1 files to the format supported by BAnAnASpy

#####  UPDATES  #####
# Changement de méthode de récupération des fichiers
# Rajout des tests pour garantir la bonne execution du 'rename()'
# Gestion des fichiers non modifiables
# lever d'exception avant de rename les fichiers
# Gestion des erreurs possibles de nom

#####  Reste à faire  #####


from glob import glob
from os import rename, system
from os.path import isdir, dirname, basename, splitext
from re import compile
import tkinter as tk
from tkinter.filedialog import askdirectory


# Choice of work folder #
root = tk.Tk()
root.withdraw()
path = askdirectory(title='Please Select your work folder')

# Make list #
def listdirectory(path):
    files = []
    list = glob(path+'\\*')
    for i in list:
        if isdir(i):
            files.extend(listdirectory(i))
        else:
            files.append(i)
    return files

path_files = listdirectory(path)
regex = compile('.ab1')
path_oldFiles = [path_file for path_file in path_files if regex.search(path_file)]
# print(path_oldFiles)

path_filesDir = []
oldFiles = []
for file in path_oldFiles:
    path_filesDir.append(dirname(file))
    oldFiles.append(basename(splitext(file)[0]))
# print(path_filesDir)
# print(oldFiles)

# Calculation of minParts #
# Sert à éliminer les fichiers contenant un '_' ponctuel dans le nom de la seq
minParts = 99999
for file in oldFiles:
    if len(file.split('_')) < minParts:
        minParts = len(file.split('_'))
# print(minParts)

# Make newFiles list #
newFiles = []
unmodifiedFiles = []
delFilesIndex = []
wellRegex = compile(r'^[A-Z][0-9][0-9]$')
seqRegex = compile(r'[a-zA-Z]+')
for i, file in enumerate(oldFiles):
    wellIndex = []
    seqNameIndex = []
    fileSplit = file.split('_')
    # print(fileSplit)
    if len(fileSplit) == minParts:
        for j, part in enumerate(fileSplit):
            if wellRegex.search(part):
                wellIndex.append(j)
            if '-' in part and seqRegex.search(part):
                seqNameIndex.append(j)
        # print(wellIndex)
        # print(seqNameIndex)
        if len(wellIndex) == 1 and len(seqNameIndex) == 1:
            if wellIndex[0] != 0 and seqNameIndex[0] != 1:
                newFiles.append(fileSplit[wellIndex[0]]+'_'+fileSplit[seqNameIndex[0]])
            else:
                delFilesIndex.append((i, file))
                print(file, ': No need to rename')
        else:
            delFilesIndex.append((i, file))
            if len(wellIndex) != 1:
                print(file, ': 2 wells possible. Check your fileName')
            if len(seqNameIndex) != 1:
                print(file, ': 2 seqNames possible. Check your fileName')
    else:
        delFilesIndex.append((i, file))
        print(file, ": Longer fileName. '_' present in the seqName or rename separately")

# Deletion of unmodified files #
if len(delFilesIndex) != 0:
    # print(delFilesIndex)
    delFilesIndex.reverse()
    # print(delFilesIndex)
    for tuple in delFilesIndex:
        # print(tuple[0])
        # print(tuple[1])
        unmodifiedFiles.append(tuple[1])
        del path_filesDir[tuple[0]]
        del oldFiles[tuple[0]]
    unmodifiedFiles.reverse()

# Renaming files #
countRename = 0
if len(path_filesDir) == len(oldFiles) and len(oldFiles) == len(newFiles):
    i = 0
    while i < len(oldFiles):
        rename(path_filesDir[i] + '\\' + oldFiles[i] + '.ab1', path_filesDir[i] + '\\' + newFiles[i] + '.ab1')
        countRename += 1
        i += 1
else:
    raise ValueError('Incoherent list')

print('')
# print('Original:', path_oldFiles)
# print(len(path_oldFiles))
# print('path_filesDir:', path_filesDir)
# print(len(path_filesDir))
print('oldFiles:', oldFiles)
# print(len(oldFiles))
print('newFiles:', newFiles)
# print(len(newFiles))
# print('unmodifiedFiles:', unmodifiedFiles)
# print(len(unmodifiedFiles))

print('')
print(str(countRename) + '/' + str(len(path_oldFiles)) + ' files modified')
if len(unmodifiedFiles) != 0:
    print(str(len(unmodifiedFiles)) + ' file(s) have not been modified')
    print('')
    print('unmodifiedFiles:')
    print(unmodifiedFiles)

print('')
system('pause')
