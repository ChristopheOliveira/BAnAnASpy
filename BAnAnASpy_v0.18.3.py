#!/usr/bin/python
# -*- coding:utf-8 -*-

# Script by Christophe Oliveira
# Contact : christophe.oliveira@aphp.fr, christophe.oliveira33@gmail.com

# BAnAnASpy_v0.18
#####  UPDATES  #####
# Modification des couleurs de Results de la methode cellFormatEdition() et du merge pb
# Correction 'feuil1.freeze_panes(1, 2)' 'feuil4.freeze_panes(1, 1)' 'feuil5.freeze_panes(1, 1)'
# Ajout d'un path par défaut pour la DB 'C:/' et création si besoin du dossier
# Changement report name 'BAnAnASpy_Report'
# Changement dans le log pour Exon_Totalcoverage et fusion des coverage 'partial' et 'incomplete'
# Retrait de la der var inutile de generateSynthesisList 'variantsListCDS' pour les pb de merge
# Rajout de la posAb1 dans la synthèse
# Réglage proxy en automatique
# Génération d'une genesList avant l'analyse de la DB
# Print format de la synthesis
# Changement de la lettre du serveur PAILLASSES

#####  Reste à faire  #####
# TODO: Modif FDT pour remplacer '-' par '.' ou '_' ds seq name
# TODO: Test sur Linux OK mais pb avec fonction os.system('pause')
# TODO: Nettoyage du code inutile
# TODO: Recherche des gb et fa avec erreur (MPC2)
# TODO: Alignement sur exon duplex (UQCRC2)
# TODO: pb Extraction des info exons pour duplex
# TODO: pb mut sur 2e partie exon non créée (exp B402-UQCRC2-9F) + disparition du 10R des listes
# TODO: Disparition de SERPINA1-2R et 5R (bug du reverse complement vide)
# TODO: Gestion des exon 02-03F avec tiret '-'
# TODO: Gestion de l'ADN mito
# TODO: Rajout de la possibilité de voir les mutations avec EMBOSS abiview et les afficher dans le report
# TODO: refaire un test de couverture sur le consensus en cas de couv partielle ou incomplete
# TODO: Revoir la méthode de fusion des variants G&C avec les seq en amont et aval
# TODO: Recalculer les pos des mut en cas d'INS homo
# TODO: Décalage du stop lors du découpage exon de l'alignement en cas d'INS
# TODO: Faire une methode createVariantListErrorSheet pour masquage des colonnes différentes
# TODO: Tester le programme sur des indels
# TODO: Verrouiller les fichiers de la DB en lecture seule et empêcher la suppression des fichiers tout en permettant l'écriture
# TODO: Retirer tout ce qui concerne le NM
# TODO: Faire des tests sur du CDNA Si test non concluant créer un dossier CDNA/Fasta CDNA/Genbank CDNA/CDS CDNA/Transcript
# TODO: Gestion des qv pour la seq consensus > fastq
# TODO: Pour la DB, ne travailler qu'avec les gb sans fasta

#####  Setup auto-py-to-exe  #####
# One-File / Console-Based
# TODO: Erreur win32com Installer win32com
# TODO: Revoir les réglages de la fenetre console pour l'exe

#####  Libraries to install  #####
# 'biopython' 'xlsxwriter' ('xlutils' 'lxml')
# Linux: pip install 'package' ou pip3 install 'package'
# Windows: python ou py -m pip install 'package'

# 'tkinter' if necessary
# sudo apt-get install python-tk pour Python2.X
# sudo apt-get install python3-tk pour Python3.X

#####  Things to know about filename of SEQ  #####
# Use dash '-' to separate each part name
# Fist part must be the patient id
# Second part must be the gene
# The order of other parts is free
# The primer part must finish by 'F' or 'S' for forward sequences
# The primer part must finish by 'R' for reverse sequences
# For multiplex or partial exon names, separate each number by underscore '_' or dot '.' or slash '/'
# Exp filenames before sequencing: 1201645-ABCB4-8F, 1123546-ABCB11-DPN-8_9R
# Exp ab1 files: A01_1201645-ABCB4-8F.ab1, H02_1123546-ABCB11-DPN-8_9R.ab1

import glob
import os
import sys
from datetime import datetime
# from time import strftime
import logging
logger = logging.Logger('catch_all')
import re
from math import floor
# from operator import itemgetter
# from lxml import etree
from xml.dom import minidom
import xlsxwriter
from xlrd import open_workbook
# from xlutils.copy import copy
from Bio import SeqIO
from Bio import Align
from Bio.Seq import Seq
from Bio.SeqRecord import SeqRecord
from Bio.Data.IUPACData import ambiguous_dna_values
from Bio import Entrez

pythonVersion = sys.version[0]
"""
# Ouverture de fichier sous Python2 #
if pythonVersion == str(2):
    import Tkinter as tk
    from tkFileDialog import askopenfilename
    from tkMessageBox import askokcancel, askyesno
"""
# Ouverture de fichier sous Python3 #
if pythonVersion == '3':
    import tkinter as tk
    from tkinter.filedialog import askopenfilename, askdirectory
    from tkinter.messagebox import askokcancel, askyesno


def main():
    global path, path_filelist_ab1, outputFolder, path_db, path_report, path_log, logout
    global filelist_ab1, filelist_nopos, filelist_nopos_F, filelist_nopos_R, pos, pos_F, pos_R, loop, loop_F, loop_R
    global dict_Patient_Family, dict_Family_LabelFamily
    global dict_FA, dict_GB, dict_CDS, dict_Transcripts, dict_StartStopCDSList
    global fileSeqObjsList, fileSeqObjsErrorList, variantsObjList, variantsObjListCDS

    ##########  Setup  ##########
    path_db = 'T:/BAnAnASpy - DB/RefSeq_Files'  # \\Jacob\PAILLASSES
    # path_db = 'D:/BioInformatique/Projet SeqIO/RefSeq_Files'
    # path_db = 'C:/BAnAnASpy - DB/RefSeq_Files'
    # path_db = 'C:/YourFolder'  # Path of your folder which contains reference files

    familyManager = 1  # <1> if you want to manage family members, <0> if you don't
    colIdName = "Prélèvement d'origine"  # Tag of the patient id column
    colFamilyName = "Libellé de la famille"  # Tag of the family id column

    Entrez.email = "christophe.oliveira@aphp.fr"
    # Entrez.email = "yourmail@domain.com"
    organism = 'Homo sapiens'

    proxy = 'http://proxym-inter.aphp.fr:8080'
    os.environ['http_proxy'] = proxy
    os.environ['HTTP_PROXY'] = proxy
    os.environ['https_proxy'] = proxy
    os.environ['HTTPS_PROXY'] = proxy


    ##########  Create Folders and Paths  ##########
    print('')
    # print('###########################################################     PROGRAM START     ###########################################################')
    print('#########################     PROGRAM START     #########################')
    print('')
    root = tk.Tk()
    root.withdraw()  # Bloque l'affichage de la deuxième fenêtre tkinter
    path = askdirectory(title='Please Select your work folder')
    currentFolder = os.path.basename(path)  # Recupere le dernier element du path (dossier courant)
    outputFolder = path + '/Output_Files'  # Windows/Linux
    date = datetime.today().strftime('%Y-%m-%d_%Hh%Mm%S')
    path_log = path + '/log_' + currentFolder + '_' + date + '.txt'
    logout = open(path_log, 'w')
    exportConsoleToLogStart()
    print(os.path.basename(__file__))
    print('START:', date)
    print('')
    exportConsoleToLogStop()
    try:  # Permet de vérifier si le dossier existe déjà. Crée le dossier si besoin.
        os.makedirs(path_db)
    except OSError:
        if not os.path.isdir(path_db):
            raise
    try:
        os.makedirs(outputFolder)
    except OSError:
        if not os.path.isdir(outputFolder):
            raise
    try:
        os.makedirs(path_db + '/CDS/')
    except OSError:
        if not os.path.isdir(path_db + '/CDS/'):
            raise
    try:
        os.makedirs(path_db + '/FA/')
    except OSError:
        if not os.path.isdir(path_db + '/FA/'):
            raise
    try:
        os.makedirs(path_db + '/GB/')
    except OSError:
        if not os.path.isdir(path_db + '/GB/'):
            raise
    try:
        os.makedirs(path_db + '/Transcripts/')
    except OSError:
        if not os.path.isdir(path_db + '/Transcripts/'):
            raise


    ##########  Lists Generation  ##########
    # FileLists generation #
    print('##########  Seq Lists Generation  ##########')
    exportConsoleToLogStart()
    print('##########  Seq Lists Generation  ##########')
    exportConsoleToLogStop()
    generateSeqList()
    test = askokcancel('Verification of Seq Lists', "The lists are correct?")
    if test == False:
        # print('###########################################################     STOPPED BY USER     ###########################################################')
        print('#########################     STOPPED BY USER     #########################')
        print('')
        os.system('pause')
        sys.exit()

    print('##########  Patients List Generation  ##########')
    idList = generateSplitList(filelist_nopos, 0)
    idListNoDouble = sortListNoNullNoDouble(idList)
    print('Patients:', idListNoDouble)
    print('')
    exportConsoleToLogStart()
    print('##########  Patients List Generation  ##########')
    print('Patients:', idListNoDouble)
    print('')
    exportConsoleToLogStop()

    # Family Lists and dicts #
    fdtList = []
    dict_Patient_Family = {}
    dict_Family_LabelFamily = {}
    if familyManager == 1:
        print('##########  Family Lists/Dicts Generation  ##########')
        exportConsoleToLogStart()
        print('##########  Family Lists/Dicts Generation  ##########')
        exportConsoleToLogStop()
        # Open files window #
        fdtList.append(askopenfilename(title='Choice of work files'))
        if fdtList[0] != '':
            test = True
        else:
            test = False
        while test == True:
            test = askyesno('Add a new file', 'Would you like to add a new work file?')
            if test == True:
                fdtList.append(askopenfilename(title='Choice of work files'))
            else:
                break

        # Open files and generate Patient/Family Lists #
        totalPatientList = []
        totalFamilyList = []
        if len(fdtList) != 0:
            for fdt in fdtList:
                extension = os.path.splitext(fdt)[1]
                if extension == '.xls':
                    bookRead = open_workbook(fdt)  # Ouverture du fichier xls avec xlrd
                    try:
                        mySheet = bookRead.sheet_by_name(u'Worksheet1')  # Choix de la feuille
                    except:
                        mySheet = bookRead.sheet_by_name(u'Feuil1')  # Choix de la feuille
                    try:
                        patientList, familyList = getPatientList_FamilyListXLS(mySheet, colIdName, colFamilyName)
                    except:
                        print("File '" + os.path.basename(fdt) + "' can't be readed")
                        print('Error:', e)
                        exportConsoleToLogStart()
                        print("File '" + os.path.basename(fdt) + "' can't be readed")
                        print('Error:', e)
                        exportConsoleToLogStop()
                        logger.error(e, exc_info=True)
                        patientList = []
                        familyList = []
                    totalPatientList = totalPatientList + patientList
                    totalFamilyList = totalFamilyList + familyList

                elif extension == '.xml':
                    fileXML = minidom.parse(fdt)  # Ouverture du fichier xml
                    try:
                        patientList, familyList = getPatientList_FamilyListXML(fileXML, colIdName, colFamilyName)
                    except Exception as e:
                        print("File '" + os.path.basename(fdt) + "' can't be readed")
                        print('Error:', e)
                        exportConsoleToLogStart()
                        print("File '" + os.path.basename(fdt) + "' can't be readed")
                        print('Error:', e)
                        exportConsoleToLogStop()
                        logger.error(e, exc_info=True)
                        patientList = []
                        familyList = []
                    totalPatientList = totalPatientList + patientList
                    totalFamilyList = totalFamilyList + familyList
        # print(totalPatientList)
        # print(totalFamilyList)

        # Family Dicts Generation #
        generateDictsFamily(totalPatientList, totalFamilyList)
        for key in dict_Patient_Family.keys():
            print(key, ':', getFamilyFromID(key))
        print('')
        exportConsoleToLogStart()
        for key in dict_Patient_Family.keys():
            print(key, ':', getFamilyFromID(key))
        print('')
        exportConsoleToLogStop()


    ##########  Sequences Analysis  ##########
    # Convert files ab1 > fastq #
    print('##########  Convert ab1 > fastq  ##########')
    convertAb1ToFastq()

    # Reverse + Complement of Seq_R #
    print('##########  Reverse Complement of Seq_R  ##########')
    reverseComplement()

    # Trimming + object lists creation #
    fileSeqObjsList = []
    fileSeqObjsErrorList = []
    variantsObjList = []
    variantsObjListCDS = []
    print('##########  Trimming + fileSeqObjsLists  ##########')
    exportConsoleToLogStart()
    print('##########  Trimming + fileSeqObjsLists  ##########')
    exportConsoleToLogStop()
    fileSeqObjsList, fileSeqObjsErrorList = trimmingSeq_FileSeqObjListCreation()

    # Sort objList by attribute 'patientID' #
    fileSeqObjsList.sort(key=lambda obj: obj.patientID)

    # DataBase analysis and 'Entrez' requests #
    print('##########  DataBase Analysis and Missing Files Generation  ##########')
    exportConsoleToLogStart()
    print('##########  DataBase Analysis and Missing Files Generation  ##########')
    exportConsoleToLogStop()
    dict_FA = generateRefSeqDict('FA')
    dict_GB = generateRefSeqDict('GB')
    dict_CDS = generateRefSeqDict('CDS')
    dict_Transcripts = generateRefSeqDict('Transcripts')

    genesList = []
    for fileObj in fileSeqObjsList:
        if fileObj.gene not in genesList:
            genesList.append(fileObj.gene)
    exportConsoleToLogStart()
    print('GenesList:', genesList)
    exportConsoleToLogStop()

    for gene in genesList:
        if gene in dict_FA.keys() and gene in dict_GB.keys():
            print('Gene ' + gene + ' found in local DB')
            exportConsoleToLogStart()
            print('Gene ' + gene + ' found in local DB')
            exportConsoleToLogStop()
        else:
            print('Gene ' + gene + ' not found in local DB => Searching on Entrez..........')
            exportConsoleToLogStart()
            print('Gene ' + gene + ' not found in local DB => Searching on Entrez..........')
            exportConsoleToLogStop()
            searchTerm = organism + '[Orgn] AND ' + gene + '[Gene] RefSeqGene'
            try:
                handle_XML = Entrez.esearch(db='nucleotide', term=searchTerm, retmax='30')
            except:
                print('')
                print('Pb config Proxy:')
                print('-Démarrer Internet Explorer')
                print('-Cliquer sur le menu "Outils"')
                print('-Cliquer sur "Options Internet"')
                print('-Cliquer sur l onglet "Connexions"')
                print('-Cliquer sur le bouton "Paramètres réseau"')
                print('-Décocher "Script de configuration automatique"')
                print('-Cocher Serveur Proxy(Utiliser un serveur proxy local)')
                print('-puis:\nAdresse: proxym-inter.aphp.fr\nPort: 8080')
                os.system('pause')

            searchResult = Entrez.read(handle_XML)
            if not (searchResult['IdList']):
                print(gene + ' Gene not found on Entrez')
                exportConsoleToLogStart()
                print(gene + ' Gene not found on Entrez')
                exportConsoleToLogStop()
                commentary = fileObj.commentary + gene + ' Gene not found on Entrez | '
                fileObj.commentary = commentary
                if fileObj not in fileSeqObjsErrorList:
                    fileSeqObjsErrorList.append(fileObj)
            else:
                # Get Fasta files #
                # TODO: Try&except cf DB-manager
                getSeqViaEntrez('fasta', 'NG', searchResult, gene)

                # Get GenBank files #
                # TODO: Try&except cf DB-manager
                getSeqViaEntrez('gb', 'NG', searchResult, gene)

            handle_XML.close()

        # CDS and Transcript files generation #
        if gene in dict_GB.keys():
            if gene not in dict_CDS.keys() or gene not in dict_Transcripts.keys():
                generateCDSFileAndTranscripts(gene)

    # Delete fileObj in error from main List #
    for fileObj in fileSeqObjsErrorList:
        if fileObj in fileSeqObjsList:
            fileSeqObjsList.remove(fileObj)
    print('')
    exportConsoleToLogStart()
    print('')
    exportConsoleToLogStop()

    print('##########  Sequences Alignment  ##########')
    exportConsoleToLogStart()
    print('##########  Sequences Alignment  ##########')
    exportConsoleToLogStop()
    aligner = alignerGlobalParameters()
    dict_StartStopCDSList = {}
    for fileObj in fileSeqObjsList:
        print('..... Analysis in progress: ' + fileObj.fileName + ' .....')
        exportConsoleToLogStart()
        print('..... Analysis in progress: ' + fileObj.fileName + ' .....')
        exportConsoleToLogStop()
        try:
            # Get 'TranscriptID' used #
            transcriptID = os.path.basename(dict_Transcripts[fileObj.gene]).split('-')[1]
            fileObj.transcriptID = transcriptID

            # Sequences Alignment on gene #
            targetSeqG = SeqIO.read(dict_FA[fileObj.gene], 'fasta').seq
            querySeqG = SeqIO.read(fileObj.path, 'fastq-sanger').seq
            print('Alignment on gene:')
            exportConsoleToLogStart()
            print('Alignment on gene:')
            exportConsoleToLogStop()
            alignmentG = aligner2Seq(targetSeqG, querySeqG, aligner)
            alignmentAnalyzerGene(alignmentG, targetSeqG, fileObj)

            # Coverage Verification #
            if fileObj.gene not in dict_StartStopCDSList:
                posCDSExtraction(fileObj.gene)
            coverageVerification(alignmentG, fileObj, shift=2)

            # Sequences Alignment on transcript #
            targetSeqC = SeqIO.read(dict_Transcripts[fileObj.gene], 'fasta').seq
            querySeqC = querySplitExon(alignmentG, fileObj.gene)
            print('Alignment on transcript:')
            exportConsoleToLogStart()
            print('Alignment on transcript:')
            exportConsoleToLogStop()
            alignmentC = aligner2Seq(targetSeqC, querySeqC, aligner)
            alignmentAnalyzerExon(alignmentC, targetSeqC, fileObj)
            print('')
            exportConsoleToLogStart()
            print('')
            exportConsoleToLogStop()

        except Exception as e:
            print('Error:', e)
            exportConsoleToLogStart()
            print('Error:', e)
            print('')
            exportConsoleToLogStop()
            logger.error(e, exc_info=True)
            commentary = fileObj.commentary + str(e) + ' | '
            fileObj.commentary = commentary
            if fileObj not in fileSeqObjsErrorList:
                fileSeqObjsErrorList.append(fileObj)
            print('')

    # Delete fileObj in error from main List #
    for fileObj in fileSeqObjsErrorList:
        if fileObj in fileSeqObjsList:
            fileSeqObjsList.remove(fileObj)

    # Analysis of CDS variants #
    print('##########  Analysis Exonic Variants  ##########')
    exportConsoleToLogStart()
    print('')
    print('##########  Analysis Exonic Variants  ##########')
    for variantObj in variantsObjListCDS:
        variantAnalyzer(variantObj)
    print('')
    exportConsoleToLogStop()
    print('done')
    print('')

    # Merge variantsLists and variantsCDS #
    try:
        print('##########  Merge Variants  ##########')
        exportConsoleToLogStart()
        print('##########  Merge Variants  ##########')
        print('done')
        exportConsoleToLogStop()
        variantsObjList = mergevariantsObjLists(variantsObjList, variantsObjListCDS)
        print('done')

    except Exception as e:
        print('done')
        print('Error:', e)
        exportConsoleToLogStart()
        print('Error:', e)
        exportConsoleToLogStop()
        logger.error(e, exc_info=True)
    print('')


    # Annotation of variants and 'allele Status' determination #
    print('##########  Variants Annotation  ##########')
    for variantObj in variantsObjList:
        variantAnnotation(variantObj)
        alleleStatusDetermination(variantObj)
    print('done')
    print('')
    exportConsoleToLogStart()
    print('')
    print('##########  Variants Annotation  ##########')
    print('done')
    print('')
    exportConsoleToLogStop()
    """
    print('variantsObjListFusion:')
    for variant in variantsObjList:
        print(variant.__dict__)
    print('')
    """


    ##########  Writting objLists in 'Report' file  ##########
    print('##########  Writting objLists in Report file  ##########')
    # Definition of report name #
    path_report = path + '/BAnAnASpy_Report_' + currentFolder
    if len(fdtList) != 0:
        for fdt in fdtList:
            if fdt != '':
                filename = os.path.splitext(os.path.basename(fdt))[0]  # Récupère le nom du fichier sans l'extension
                filename = filename.replace('-', ' ').replace('_', ' ')
                words = filename.split(' ')
                regex = re.compile(r'^[lL][tT]\d+')
                refLT = [word for word in words if regex.search(word)][0]
                path_report = path_report + '_' + refLT
    path_report = path_report + '.xlsx'

    # TODO: Problème de mémoire à l'ouverture du report quand un autre fichier xls est ouvert.
    # Convert object Lists to txt Lists #
    fileSeqList = objectsListToList(fileSeqObjsList)
    fileSeqErrorList = objectsListToList(fileSeqObjsErrorList)
    variantsList = objectsListToList(variantsObjList)
    variantsListCDS = objectsListToList(variantsObjListCDS)
    variantsListCDS = cleanList(variantsListCDS)  # Delete alignment object

    # Make Synthesis List + Merged F-R #
    synthesisList = generateSynthesisList(fileSeqList, fileSeqErrorList, variantsList)
    variantsListFRMerged = mergeForwardReverseVariants(variantsList)  # Merge forward/reverse

    # Create Results file #
    bookWrite = xlsxwriter.Workbook(path_report)

    style_header = bookWrite.add_format({'bold': True, 'align': 'center', 'valign': 'top', 'border': True, 'text_wrap': True, 'bg_color': 'gray'})
    style_header_noborder = bookWrite.add_format({'bold': True, 'align': 'center', 'valign': 'top', 'text_wrap': True})
    style_data = bookWrite.add_format({'align': 'center', 'valign': 'vcenter', 'border': True, 'text_wrap': True})
    style_data_noborder = bookWrite.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    style_fontRed = bookWrite.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': True, 'text_wrap': True, 'color': 'red'})
    #style_orange = bookWrite.add_format({'align': 'center', 'valign': 'vcenter', 'border': True, 'text_wrap': True, 'bg_color': 'orange'})
    style_unlocked = bookWrite.add_format({'align': 'center', 'valign': 'vcenter', 'border': True, 'locked': False})

    feuil1 = bookWrite.add_worksheet('Synthesis')
    feuil2 = bookWrite.add_worksheet('FilesList')
    feuil3 = bookWrite.add_worksheet('FilesList_Error')
    feuil4 = bookWrite.add_worksheet('VariantsList')
    feuil5 = bookWrite.add_worksheet('VariantsList_Error')

    createSynthesisSheet_xlsxwriter(feuil1, style_header)
    createFileListSheet_xlsxwriter(feuil2, style_header_noborder)
    createFileListErrorSheet_xlsxwriter(feuil3, style_header_noborder)
    createVariantListSheet_xlsxwriter(feuil4, style_header_noborder)
    createVariantListSheet_xlsxwriter(feuil5, style_header_noborder)

    # Write Lists on Results file #
    writeSheetByLine_xlsxwriter(feuil1, synthesisList, style_data, 1)
    writeSheetByLine_xlsxwriter(feuil2, fileSeqList, style_data_noborder, 0)
    writeSheetByLine_xlsxwriter(feuil3, fileSeqErrorList, style_data_noborder, 0)
    writeSheetByLine_xlsxwriter(feuil4, variantsListFRMerged, style_data_noborder, 0)
    writeSheetByLine_xlsxwriter(feuil5, variantsListCDS, style_data_noborder, 0)

    # Unlocked 'Checked' column #
    writeUnlockedEmptyCells(feuil1, synthesisList, 0, style_unlocked)

    # Cell Format Edition #
    cellFormatEdition(bookWrite, synthesisList)

    # Add comments on Results file #
    feuil1.write(len(synthesisList) + 3, 8, 'Report generated by script: ' + os.path.basename(__file__), style_data)

    if len(variantsObjListCDS) > 0:
        feuil1.write(len(synthesisList) + 5, 8, 'Merge pb of variantLists! Please check annotations for this file(s):', style_fontRed)
        fileListMergeFailed = getFileListMergeFailed(variantsListCDS)
        row = 0
        for file in fileListMergeFailed:
            feuil1.write(len(synthesisList) + 6 + row, 8, file, style_fontRed)
            row += 1

    # Add Autofilter #
    feuil1_nrows = len(synthesisList)
    feuil2_nrows = len(fileSeqList)
    feuil3_nrows = len(fileSeqErrorList)
    feuil4_nrows = len(variantsList)
    feuil5_nrows = len(variantsListCDS)
    feuil1.autofilter(0, 0, feuil1_nrows, 13)
    feuil2.autofilter(0, 0, feuil2_nrows, 14)
    feuil3.autofilter(0, 0, feuil3_nrows, 14)
    feuil4.autofilter(0, 0, feuil4_nrows, 19)
    feuil5.autofilter(0, 0, feuil5_nrows, 19)

    # Add Freeze panes #
    feuil1.freeze_panes(1, 2)
    feuil2.freeze_panes(1, 1)
    feuil3.freeze_panes(1, 1)
    feuil4.freeze_panes(1, 1)
    feuil5.freeze_panes(1, 1)

    # Add Sheets protection #
    # TODO: l'option 'sort' ne fonctionne pas
    feuil1.protect(password='', options={'format_cells': True, 'sort': True, 'autofilter': True})
    feuil2.protect(password='', options={'format_cells': True, 'sort': True, 'autofilter': True})
    feuil3.protect(password='', options={'format_cells': True, 'sort': True, 'autofilter': True})
    feuil4.protect(password='', options={'format_cells': True, 'sort': True, 'autofilter': True})
    feuil5.protect(password='', options={'format_cells': True, 'sort': True, 'autofilter': True})

    # Print format #
    feuil1.set_header(header=os.path.basename(path_report))
    feuil1.repeat_rows(0)
    feuil1.set_landscape()
    feuil1.set_margins(left=0.2, right=0.2, top=0.5, bottom=0.5)
    feuil1.fit_to_pages(1, 0)

    # Save and close #
    bookWrite.close()
    print('done')
    print('')
    exportConsoleToLogStart()
    print('##########  Writting objLists in Report file  ##########')
    print('done')
    print('')
    exportConsoleToLogStop()

    # print('###########################################################     END WITH SUCCESS     ###########################################################')
    print('#########################     END WITH SUCCESS     #########################')
    print('')
    exportConsoleToLogStart()
    print('END:', datetime.today().strftime('%Y-%m-%d_%Hh%Mm%S'))
    exportConsoleToLogStop()
    os.system('pause')
    sys.exit()


#####  Fonctions Gestion Listes  #####

def generateSeqList():
    global path, path_filelist_ab1, filelist_ab1, filelist_nopos, filelist_nopos_F, filelist_nopos_R, pos, pos_F, pos_R, loop, loop_F, loop_R

    files = []
    path_files = []
    for (repertoire, sousRepertoires, fichiers) in os.walk(path):
        files.extend(fichiers)
        for fichier in fichiers:
            path_files.append(os.path.join(repertoire, fichier))
    pattern = re.compile('.ab1')  # Definition de l'expression reguliere
    filelist_ab1 = [file for file in files if pattern.search(file)]  # Filtre selon le pattern defini
    path_filelist_ab1 = [path_file for path_file in path_files if pattern.search(path_file)]

    filelist = []
    filelist_nopos = []
    pos = []
    loop = len(filelist_ab1)
    n = 0
    while n < loop:
        filesplit = os.path.splitext(filelist_ab1[n])  # Split le nom du fichier et l'extension
        filelist.append(filesplit[0])
        filelist_nopos.append((filelist[n])[4:])
        pos.append((filelist[n])[0:3])
        n = n + 1

    # filelist_ab1_R = glob.glob('*-*-*[0-9]R*.ab1')  # Pour dossier courant uniquement
    pattern = re.compile(r'(\w+[-]){2}\w*[-_.]*([0-9]+[rR])[-]*\w*(.ab1)')  # Definition de l'expression reguliere
    # pattern = re.compile(r'(\w+[-]){2}[a-zA-Z]*[0-9]+[rR]''*.ab1')  # Definition de l'expression reguliere
    filelist_ab1_R = [file for file in files if pattern.search(file)]  # Filtre selon le pattern defini
    # print(filelist_ab1_R)

    filelist_R = []
    filelist_nopos_R = []
    pos_R = []
    loop_R = len(filelist_ab1_R)
    n = 0
    while n < loop_R:
        filesplit_R = os.path.splitext(filelist_ab1_R[n])
        filelist_R.append(filesplit_R[0])
        filelist_nopos_R.append((filelist_R[n])[4:])
        pos_R.append((filelist_R[n])[0:3])
        n = n + 1

    filelist_F = filelist[:]  # copie la liste de maniere indépendante
    filelist_nopos_F = []
    pos_F = []
    n = 0
    while n < loop_R:
        filelist_F.remove(filelist_R[n])
        n = n + 1
    loop_F = len(filelist_F)
    n = 0
    while n < loop_F:
        filelist_nopos_F.append((filelist_F[n])[4:])
        pos_F.append((filelist_F[n])[0:3])
        n = n + 1

    print('Seq_F:', filelist_nopos_F)
    print('Seq_R:', filelist_nopos_R)
    print(len(filelist_nopos_F), 'Seq_F |', len(filelist_nopos_R), 'Seq_R |', len(filelist_nopos), 'Seq au total')
    if loop_F != loop_R:
        print("Warning! Difference between number of seq_F and seq_R")
    print('')
    exportConsoleToLogStart()
    print('Seq_F:', filelist_nopos_F)
    print('Seq_R:', filelist_nopos_R)
    print(len(filelist_nopos_F), 'Seq_F |', len(filelist_nopos_R), 'Seq_R |', len(filelist_nopos), 'Seq total')
    print('')
    exportConsoleToLogStop()


def getPatientList_FamilyListXML(fileXML, colIdName, colFamilyName):

    """
        # Décomposition de tous les noeuds d'un arbre #
        xmldoc = minidom.parse('test.xml')
        rows = xmldoc.getElementsByTagName('Row')
        for row in rows:
            cells = row.getElementsByTagName('Cell')
            for cell in cells:
                dataNodes = cell.getElementsByTagName('Data')
                for dataNode in dataNodes:
                    if dataNode.hasChildNodes():
                        print(dataNode.childNodes[0].nodeValue)
    """

    rows = fileXML.getElementsByTagName('Row')

    # Make first line #
    firstRow = []
    cells = rows[0].getElementsByTagName('Cell')
    for cell in cells:
        dataNodes = cell.getElementsByTagName('Data')
        for dataNode in dataNodes:
            if dataNode.hasChildNodes():
                firstRow.append(dataNode.childNodes[0].nodeValue)
            else:
                firstRow.append('')

    # Determination of wished column positions #
    for i in range(len(firstRow)):
        if firstRow[i] == colIdName:
            numColPatient = i
        if firstRow[i] == colFamilyName:
            numColFamily = i

    # Make Patients list #
    patientList = []
    i = 1
    while i < len(rows):
        cells = rows[i].getElementsByTagName('Cell')
        dataNodes = cells[numColPatient].getElementsByTagName('Data')
        for dataNode in dataNodes:
            if dataNode.hasChildNodes():
                patientList.append(dataNode.childNodes[0].nodeValue)
            else:
                patientList.append('')
        i = i + 1

    # Make Family list #
    familyList = []
    i = 1
    while i < len(rows):
        cells = rows[i].getElementsByTagName('Cell')
        dataNodes = cells[numColFamily].getElementsByTagName('Data')
        for dataNode in dataNodes:
            if dataNode.hasChildNodes():
                familyList.append(dataNode.childNodes[0].nodeValue)
            else:
                familyList.append('')
        i = i + 1

    if len(patientList) != len(familyList):
        raise ValueError('Nb of values Patient/Family not equals')

    return patientList, familyList


def getPatientList_FamilyListXLS(bookSheet, colIdName, colFamilyName):

    line1 = bookSheet.row_values(0)  # Valeurs de la premiere ligne

    # Determination of wished column positions #
    for i in range(len(line1)):
        if line1[i] == colIdName:
            numColPatient = i
        if line1[i] == colFamilyName:
            numColFamily = i

    # Make Patients & Family lists #
    patientList = bookSheet.col_values(numColPatient)[1:]
    familyList = bookSheet.col_values(numColFamily)[1:]

    if len(patientList) != len(familyList):
        raise ValueError('Nb of values Patient/Family not equals')

    return patientList, familyList


def generateDictsFamily(patientList, familyList):
    global dict_Patient_Family, dict_Family_LabelFamily

    i = 0
    while i < len(patientList):
        if patientList[i] not in dict_Patient_Family.keys() and familyList[i] != '':
            dict_Patient_Family[patientList[i]] = familyList[i]
        i = i + 1

    n = 1
    for family in familyList:
        if family not in dict_Family_LabelFamily.keys() and family != '':
            dict_Family_LabelFamily[family] = 'Family ' + str(n)
            n = n + 1


def getFamilyFromID(id):
    global dict_Patient_Family, dict_Family_LabelFamily

    if id in dict_Patient_Family.keys():
        family = dict_Family_LabelFamily[dict_Patient_Family[id]]
    else:
        family = ''

    return family


def getExonFromFileName(fileName):

    # TODO: Erreur si chiffre seul
    exon = ''
    i = 2
    while i < len(fileName.split('-')):
        if re.match(r"[a-zA-Z]*[0-9]+[sfrSFR]$", fileName.split('-')[i]) is None:
            i = i + 1
        else:
            exon = fileName.split('-')[i]
            break
    if exon == '':  # Pour les PCRs duplex ou partielles
        i = 2
        while i < len(fileName.split('-')):
            if re.match(r"[a-zA-Z]*[0-9]+[_./][0-9]+[sfrSFR]$", fileName.split('-')[i]) is None:
                i = i + 1
            else:
                exon = fileName.split('-')[i]
                break
    exon = re.findall('\d+', exon)  # Extrait les nombres du string. Génère une liste!
    # print(exon)
    exon = exon[0]
    if exon[0] == '0':
        exon = exon[1]

    return exon


def matchListPairwise():
    global filelist_nopos_F, filelist_nopos_R, indexlist_dual_F, indexlist_dual_R, indexlist_single_F, indexlist_single_R
    # TODO: Revoir les var global indexlist > local

    #indexlist dual mode#
    i = 0
    j = 0
    char = 0
    indexlist_dual_F = []
    indexlist_dual_R = []
    print('Seq_F:', filelist_nopos_F)
    print('Seq_R:', filelist_nopos_R)
    print('')
    while i < loop_F:
        j = 0
        while j < loop_R:
            char = 0
            test = 0
            length = min(len(filelist_nopos_F[i]), len(filelist_nopos_R[j]))
            print(filelist_nopos_F[i])
            print(filelist_nopos_R[j])
            while char < length:
                if filelist_nopos_F[i][char] != filelist_nopos_R[j][char]:
                    test = test + 1
                if test > 1:
                    print("break")
                    print('')
                    break
                char = char + 1
            if test < 2:
                print("match")
                print('')
                indexlist_dual_F.append(i)
                indexlist_dual_R.append(j)
            j = j + 1
        i = i + 1

    # indexlist single mode#
    indexlist_single_F = []
    n = 0
    while n < loop_F:
        indexlist_single_F.append(n)
        n = n + 1
    n = 0
    # TODO: Attention bug si 2 fichiers du meme nom
    while n < len(indexlist_dual_F):
        indexlist_single_F.remove(indexlist_dual_F[n])
        n = n + 1

    indexlist_single_R = []
    n = 0
    while n < loop_R:
        indexlist_single_R.append(n)
        n = n + 1
    n = 0
    while n < len(indexlist_dual_R):
        indexlist_single_R.remove(indexlist_dual_R[n])
        n = n + 1

    print('indexlist_dual_F:  ', indexlist_dual_F)
    print('indexlist_dual_R:  ', indexlist_dual_R)
    print('indexlist_single_F:', indexlist_single_F)
    print('indexlist_single_R:', indexlist_single_R)
    print('')


def generateSplitList(list, index):
    # get the part separate by '-'(6) at 'index' position from list to new one

    splitList = []
    i = 0
    while i < len(list):
        splitList.append(list[i].split('-')[index])
        i = i + 1

    return splitList


def mergevariantsObjLists(listG, listC):
    global dict_StartStopCDSList

    nbMutIntron = 0
    i = 0
    # j = 0
    while i < len(listG):
        start = dict_StartStopCDSList[listG[i].gene][0]
        stop = dict_StartStopCDSList[listG[i].gene][1]

        k = 0
        test = False
        while k < len(start):
            if listG[i].posG > int(start[k]) and listG[i].posG < int(stop[k]):
                test = True  # test pour s'assurer qu'il soit exonic
                listG[i].exonic = 'Yes'
                # TODO: Faire un test avec une mut en prem ou der pos de l'exon
            k = k + 1

        if test == True:
            j = 0
            while j < len(listC):

                if listG[i].fileName == listC[j].fileName and listG[i].refG == listC[j].refC and listG[i].altG == listC[j].altC:
                    listG[i].posC = listC[j].posC
                    listG[i].refC = listC[j].refC
                    listG[i].altC = listC[j].altC
                    listG[i].posP = listC[j].posP
                    listG[i].refP = listC[j].refP
                    listG[i].altP = listC[j].altP
                    listG[i].annotation = listC[j].annotation
                    listG[i].commentary = listG[i].commentary + listC[j].commentary
                    listC.remove(listC[j])
                    j = len(listC)
                else:
                    j = j + 1

        else:
            nbMutIntron = nbMutIntron + 1
            k = 0
            while k < len(start):
                if listG[i].posG > (int(start[k]) - 2) and listG[i].posG < (int(stop[k]) + 2):
                    listG[i].exonic = 'Splice'
                    # TODO: Faire un test avec une mut sur site de splice
                k = k + 1
        i = i + 1

    if len(listC) > 0:
        raise ValueError('Merge pb of 2 lists! Please check annotations.')

    return listG


def mergeForwardReverseVariants(variantList):

    line = 0
    while line < len(variantList):
        tag = variantList[line][2] + variantList[line][4] + str(variantList[line][10]) + variantList[line][11] +\
              variantList[line][12]

        linebis = line + 1
        while linebis < len(variantList):
            tagbis = variantList[linebis][2]+variantList[linebis][4] + str(variantList[linebis][10]) +\
                     variantList[linebis][11] + variantList[linebis][12]

            if tag == tagbis:

                # Merge of values #
                list = variantList[linebis][0].split('-')
                list = list[2:]
                name = '-'.join(list)
                variantList[line][0] = variantList[line][0] + '_' + name
                variantList[line][1] = variantList[line][1] + '_' + variantList[linebis][1]
                variantList[line][9] = str(variantList[line][9]) + '_' + str(variantList[linebis][9])
                variantList[line][19] = variantList[line][19] + '_' + variantList[linebis][19]

                # Delete line #
                del variantList[linebis]

                break

            linebis += 1

        line += 1

    return variantList


def generateSynthesisList(fileList, fileErrorList, varList):

    synthesisList = []

    if len(fileErrorList) != 0:
        for file in fileErrorList:
            # Independent copy of file #
            fileCopy = file[:]

            # Delete column #
            del fileCopy[14:]

            # Line edition #
            fileCopy[7] = 'Not Analysed'
            fileCopy[8] = ''
            fileCopy[9] = ''
            fileCopy[10] = ''
            fileCopy[13] = ''

            synthesisList.append(fileCopy)

    if len(fileList) != 0:
        for file in fileList:
            # Independent copy of line #
            fileCopy = file[:]

            # Delete column #
            del fileCopy[14:]

            # Make variantType & annotations #
            variantType = []
            annotations = []
            for variant in varList:
                if fileCopy[0] == variant[0]:
                    variantType.append(variant[8])
                    if variant[6] != '':
                        annotations.append(variant[6] + ' :' + variant[7] + ': ' + str(variant[9]))

            # Make result #
            result = []
            if len(variantType) == 0:
                result.append('No Variants Found')
            else:
                result.append('Variant(s) Found')
            if 'coverage' in fileCopy[13]:
                result.append('Partial Coverage')
            if 'Homo Ins' in fileCopy[13]:
                result.append('Uninterpretable due to Insertion')

            # Make variants numbers #
            intronic = variantType.count('No')
            exonic = variantType.count('Yes')
            splice = variantType.count('Splice')

            # Line edition #
            fileCopy[7] = ' | '.join(result)
            fileCopy[8] = intronic
            fileCopy[9] = exonic
            fileCopy[10] = splice
            fileCopy[13] = ' // '.join(annotations)

            synthesisList.append(fileCopy)

    return synthesisList


def getFileListMergeFailed(varList):

    fileList = []
    for var in varList:
        fileList.append(var[0])

    fileList = sortListNoNullNoDouble(fileList)

    return fileList


def sortListNoNullNoDouble(listToSort):

    listNoNull = list(filter(None, listToSort))  # Filtre les val vides
    listNoNullNoDouble = list(set(listNoNull))  # Supprime les doublons

    return listNoNullNoDouble


def objectsListToList(objList):

    lines = []
    for obj in objList:
        line = list(obj.__dict__.values())
        lines.append(line)

    return lines


def cleanList(list):
    # Delete all values that is not 'string' nor 'int' from list #

    line = 0
    while line < len(list):
        col = 0
        while col < len(list[0]):
            if not isinstance(list[line][col], (str, int)):
                list[line][col] = ''
            col = col + 1
        line = line + 1

    return list


#####  Fonctions DataBase  #####

def generateRefSeqDict(refType):
    global path_db

    path_Ref = path_db + '/' + refType

    files = []
    dict = {}
    for (repertoire, sousRepertoires, fichiers) in os.walk(path_Ref):
        files.extend(fichiers)

    for file in files:
        key = file.split('-')[0]
        value = path_Ref + '/' + file
        dict[key] = value

    return dict


def getSeqViaEntrez(retType, refType, searchResult, gene):
    global path_db, dict_FA, dict_GB

    n = 0
    while n < len(searchResult['IdList']):
        handle_seq = Entrez.efetch(db="nucleotide", id=searchResult['IdList'][n], rettype=retType, retmode='text')
        seq_record = SeqIO.read(handle_seq, retType)
        # print(seq_record.id)
        if refType in seq_record.id and gene in seq_record.description:
            print('=> ' + retType + ' Found')
            exportConsoleToLogStart()
            print('=> ' + retType + ' Found')
            exportConsoleToLogStop()
            break
        else:
            pass
        n = n + 1

    handle_seq.close()

    if retType == 'fasta':
        SeqIO.write(seq_record, path_db + '/FA/' + gene + '-' + seq_record.id + '.fa', 'fasta')
        exportConsoleToLogStart()
        print(seq_record)
        exportConsoleToLogStop()
    elif retType == 'gb':
        SeqIO.write(seq_record, path_db + '/GB/' + gene + '-' + seq_record.id + '.gb', 'genbank')
        exportConsoleToLogStart()
        print('ID:', seq_record.id)
        print('Description:', seq_record.description)
        print('Seq:', seq_record.seq[0:80],'....')
        exportConsoleToLogStop()

    if retType == 'fasta' and refType == 'NG':
        dict_FA[gene] = path_db + '/FA/' + gene + '-' + seq_record.id + '.fa'
    if retType == 'gb' and refType == 'NG':
        dict_GB[gene] = path_db + '/GB/' + gene + '-' + seq_record.id + '.gb'


def generateCDSFileAndTranscripts(gene):
    global path_db, dict_GB, dict_CDS, dict_Transcripts

    print('CDS & Transcript File Generation..........')
    exportConsoleToLogStart()
    print('CDS & Transcript File Generation..........')
    exportConsoleToLogStop()
    for rec in SeqIO.parse(dict_GB[gene], 'genbank'):
        if rec.features:
            for feature in rec.features:
                if feature.type == 'mRNA':
                    transcript_id = feature.qualifiers['transcript_id'][0]
                if feature.type == 'CDS':
                    prot_id = feature.qualifiers['protein_id'][0]
                    cds_location = feature.location
                    rec_Transcript_Seq = feature.location.extract(rec).seq
                    break

    # CDS File #
    cds_location = str(cds_location)
    cds_location = cds_location.replace('join', '')
    cds_location = cds_location.replace('(+)', '').replace('(-)', '')
    cds_location = cds_location.replace('>', '').replace('<', '')
    cds_location = cds_location.replace(' ', '')
    cds_location = cds_location.replace('{', '').replace('}', '')
    cds_location = cds_location.replace('[', '').replace(']', '')
    cds_location = cds_location.split(',')

    path_cdsFile = path_db + '/CDS/' + gene + '-' + transcript_id + '-' + prot_id + '.txt'
    with open(path_cdsFile, 'w') as cdsFile:
        i = 0
        while i < len(cds_location):
            cdsFile.write(
                gene + '_' + str(i + 1) + '\t' + cds_location[i].split(':')[0] + '\t' + cds_location[i].split(':')[1] + '\n')
            i = i + 1
        cdsFile.close()
    dict_CDS[gene] = path_db + '/CDS/' + gene + '-' + transcript_id + '-' + prot_id + '.txt'

    # Transcript File #
    rec_Transcript = SeqRecord(rec_Transcript_Seq)
    rec_Transcript.id = gene + '-' + transcript_id + '-' + prot_id
    rec_Transcript.name = gene + '-' + transcript_id + '-' + prot_id
    rec_Transcript.description = 'GENE: ' + gene + ' | Transcript _ID: ' + transcript_id + ' | Protein_ID: ' + prot_id + ' | LENGTH: ' + str(len(rec_Transcript))
    SeqIO.write(rec_Transcript, path_db + '/Transcripts/' + rec_Transcript.name + '.fa', 'fasta')
    dict_Transcripts[gene] = path_db + '/Transcripts/' + rec_Transcript.name + '.fa'
    print(gene + ': CDS/Transcript files done')
    exportConsoleToLogStart()
    print('Transcript:')
    print(rec_Transcript)
    print('')
    exportConsoleToLogStop()


def posCDSExtraction(gene):
    global dict_CDS, dict_StartStopCDSList

    startList = []
    stopList = []
    startStopCDSList = ['', '']
    with open(dict_CDS[gene], 'r') as file:
        lines = file.read().splitlines()
        for line in lines:
            startList.append(line.split('\t')[1])
            stopList.append(line.split('\t')[2])
        file.close()
    startStopCDSList[0] = startList
    startStopCDSList[1] = stopList
    if gene not in dict_StartStopCDSList.keys():
        dict_StartStopCDSList[gene] = startStopCDSList


#####  Fonctions Excel  #####

def createSynthesisSheet_xlsxwriter(bookSheet, style):

    bookSheet.write(0, 0, 'Checked', style)
    bookSheet.write(0, 1, 'File', style)
    bookSheet.write(0, 2, 'Well', style)
    bookSheet.write(0, 3, 'Patient ID', style)
    bookSheet.write(0, 4, 'Famille', style)
    bookSheet.write(0, 5, 'Gene', style)
    bookSheet.write(0, 6, 'Exon', style)
    bookSheet.write(0, 7, 'Transcript ID', style)
    bookSheet.write(0, 8, 'Results', style)
    bookSheet.write(0, 9, 'Intronic Variants', style)
    bookSheet.write(0, 10, 'Exonic Variants', style)
    bookSheet.write(0, 11, 'Splice', style)
    bookSheet.write(0, 12, 'Avg QV', style)
    bookSheet.write(0, 13, 'Avg QV Trimmed', style)
    bookSheet.write(0, 14, 'Variant(s)', style)

    bookSheet.set_column(0, 0, 8)
    bookSheet.set_column(1, 1, 30)
    bookSheet.set_column(2, 2, 7)
    bookSheet.set_column(3, 3, 15)
    bookSheet.set_column(4, 4, 11)
    bookSheet.set_column(5, 5, 12)
    bookSheet.set_column(6, 6, 8)
    bookSheet.set_column(7, 7, 17)
    bookSheet.set_column(8, 8, 30)
    bookSheet.set_column(9, 13, 10)
    bookSheet.set_column(14, 14, 40)


def createFileListSheet_xlsxwriter(bookSheet, style):

    bookSheet.write(0, 0, 'File', style)
    bookSheet.write(0, 1, 'Well', style)
    bookSheet.write(0, 2, 'Patient ID', style)
    bookSheet.write(0, 3, 'Famille', style)
    bookSheet.write(0, 4, 'Gene', style)
    bookSheet.write(0, 5, 'Exon', style)
    bookSheet.write(0, 6, 'Transcript ID', style)
    bookSheet.write(0, 7, 'Trimming Start', style)
    bookSheet.write(0, 8, 'Trimming End', style)
    bookSheet.write(0, 9, 'Seq Length', style)
    bookSheet.write(0, 10, 'Seq Trimmed Length', style)
    bookSheet.write(0, 11, 'Avg QV', style)
    bookSheet.write(0, 12, 'Avg QV Trimmed', style)
    bookSheet.write(0, 13, 'Commentary', style)
    bookSheet.write(0, 14, 'Path_Analyzed_File', style)

    bookSheet.set_column(0, 0, 35)
    bookSheet.set_column(1, 1, 8)
    bookSheet.set_column(2, 4, 15)
    bookSheet.set_column(5, 5, 10)
    bookSheet.set_column(6, 6, 20)
    bookSheet.set_column(7, 10, 15)
    bookSheet.set_column(11, 12, 10)
    bookSheet.set_column(13, 13, 80)
    bookSheet.set_column(14, 14, 90)

    # TODO: col 7 et 8 Hidden
    # bookSheet.set_column(7, 7, None, {'hidden': True})
    # bookSheet.set_column(8, 8, None, {'hidden': True})
    # bookSheet.col(9).hidden = True
    # bookSheet.col(10).hidden = True
    # bookSheet.col(11).hidden = True
    # bookSheet.col(12).hidden = True
    # bookSheet.col(14).hidden = True


def createFileListErrorSheet_xlsxwriter(bookSheet, style):

    bookSheet.write(0, 0, 'File', style)
    bookSheet.write(0, 1, 'Well', style)
    bookSheet.write(0, 2, 'Patient ID', style)
    bookSheet.write(0, 3, 'Famille', style)
    bookSheet.write(0, 4, 'Gene', style)
    bookSheet.write(0, 5, 'Exon', style)
    bookSheet.write(0, 6, 'Transcript ID', style)
    bookSheet.write(0, 7, 'Trimming Start', style)
    bookSheet.write(0, 8, 'Trimming End', style)
    bookSheet.write(0, 9, 'Seq Length', style)
    bookSheet.write(0, 10, 'Seq Trimmed Length', style)
    bookSheet.write(0, 11, 'Avg QV', style)
    bookSheet.write(0, 12, 'Avg QV Trimmed', style)
    bookSheet.write(0, 13, 'Commentary', style)
    bookSheet.write(0, 14, 'Path_Analyzed_File', style)

    bookSheet.set_column(0, 0, 35)
    bookSheet.set_column(1, 1, 8)
    bookSheet.set_column(2, 4, 15)
    bookSheet.set_column(5, 5, 10)
    bookSheet.set_column(6, 6, 20)
    bookSheet.set_column(7, 10, 15)
    bookSheet.set_column(11, 12, 10)
    bookSheet.set_column(13, 13, 80)
    bookSheet.set_column(14, 14, 90)

    # TODO: col 6, 7, 8, 9, 10, 12 Hidden
    # bookSheet.set_column(7, 7, None, {'hidden': True})
    # bookSheet.set_column(8, 8, None, {'hidden': True})
    # bookSheet.col(9).hidden = True
    # bookSheet.col(10).hidden = True
    # bookSheet.col(11).hidden = True
    # bookSheet.col(12).hidden = True
    # bookSheet.col(14).hidden = True


def createVariantListSheet_xlsxwriter(bookSheet, style):

    bookSheet.write(0, 0, 'File', style)
    bookSheet.write(0, 1, 'Well', style)
    bookSheet.write(0, 2, 'Patient ID', style)
    bookSheet.write(0, 3, 'Famille', style)
    bookSheet.write(0, 4, 'Gene', style)
    bookSheet.write(0, 5, 'Exon', style)
    bookSheet.write(0, 6, 'Annotation', style)
    bookSheet.write(0, 7, 'Allele Status', style)
    bookSheet.write(0, 8, 'Exonic', style)
    bookSheet.write(0, 9, 'Seq Ab1 Pos', style)
    bookSheet.write(0, 10, 'Gene Pos', style)
    bookSheet.write(0, 11, 'Gene Ref', style)
    bookSheet.write(0, 12, 'Gene Alt', style)
    bookSheet.write(0, 13, 'Transcript\nPos', style)
    bookSheet.write(0, 14, 'Transcript\nRef', style)
    bookSheet.write(0, 15, 'Transcript\nAlt', style)
    bookSheet.write(0, 16, 'Protein\nPos', style)
    bookSheet.write(0, 17, 'Protein\nRef', style)
    bookSheet.write(0, 18, 'Protein\nAlt', style)
    bookSheet.write(0, 19, 'Commentary', style)

    bookSheet.set_column(0, 0, 35)
    bookSheet.set_column(1, 1, 8)
    bookSheet.set_column(2, 2, 15)
    bookSheet.set_column(5, 5, 10)
    bookSheet.set_column(6, 6, 35)
    bookSheet.set_column(7, 18, 10)
    bookSheet.set_column(19, 19, 80)


def writeSheetByLine_xlsxwriter(bookSheet, dataToWrite, style, colStart):

    line = 0
    while line < len(dataToWrite):
        col = 0
        while col < len(dataToWrite[0]):
            bookSheet.write(line + 1, col + colStart, dataToWrite[line][col], style)
            col = col + 1
        line = line + 1


def writeUnlockedEmptyCells(booksheet, data, col, style):

    line = 1
    while line <= len(data):
        booksheet.write(line, col, '', style)
        line += 1


def cellFormatEdition(workbook, data):

    resultColIndex = 7
    intronColIndex = 8
    exonColIndex = 9
    spliceColIndex = 10
    qvColIndex = 11
    qvTrimColIndex = 12

    style_green = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': True, 'text_wrap': True, 'bg_color': 'green'})
    style_yellow = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': True, 'text_wrap': True, 'bg_color': 'yellow'})
    style_orange = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': True, 'text_wrap': True, 'bg_color': 'orange'})
    style_red = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': True, 'text_wrap': True, 'bg_color': 'red', 'font_color': 'white'})

    worksheet = workbook.get_worksheet_by_name('Synthesis')

    # Results Edition #
    line = 0
    while line < len(data):

        if 'Insertion' in data[line][resultColIndex]:
            worksheet.write(line + 1, resultColIndex + 1, data[line][resultColIndex], style_red)
        elif 'Coverage' in data[line][resultColIndex]:
            worksheet.write(line + 1, resultColIndex + 1, data[line][resultColIndex], style_orange)
        elif 'Not Analysed' in data[line][resultColIndex]:
            worksheet.write(line + 1, resultColIndex + 1, data[line][resultColIndex], style_yellow)
        #elif 'No Variants' or 'Variant(s)' in data[line][resultColIndex]:
            #worksheet.write(line + 1, resultColIndex + 1, data[line][resultColIndex], style_green)

        line += 1

    # Intronic Edition #
    line = 0
    while line < len(data):

        if data[line][intronColIndex] != '':
            if int(data[line][intronColIndex]) > 10:
                worksheet.write(line + 1, intronColIndex + 1, data[line][intronColIndex], style_red)
            elif int(data[line][intronColIndex]) > 5:
                worksheet.write(line + 1, intronColIndex + 1, data[line][intronColIndex], style_orange)

        line += 1

    # Exonic Edition #
    line = 0
    while line < len(data):
        if data[line][exonColIndex] != '':
            if int(data[line][exonColIndex]) > 5:
                worksheet.write(line + 1, exonColIndex + 1, data[line][exonColIndex], style_red)
            elif int(data[line][exonColIndex]) > 2:
                worksheet.write(line + 1, exonColIndex + 1, data[line][exonColIndex], style_orange)

        line += 1

    # Splice Edition #
    line = 0
    while line < len(data):

        if data[line][spliceColIndex] != '':
            if int(data[line][spliceColIndex]) > 0:
                worksheet.write(line + 1, spliceColIndex + 1, data[line][spliceColIndex], style_yellow)

        line += 1

    # QV Edition #
    line = 0
    while line < len(data):

        if int(data[line][qvColIndex]) < 20:
            worksheet.write(line + 1, qvColIndex + 1, data[line][qvColIndex], style_red)
        elif int(data[line][qvColIndex]) < 30:
            worksheet.write(line + 1, qvColIndex + 1, data[line][qvColIndex], style_orange)

        line += 1

    # QV Trimm Edition #
    line = 0
    while line < len(data):

        if int(data[line][qvTrimColIndex]) < 20:
            worksheet.write(line + 1, qvTrimColIndex + 1, data[line][qvTrimColIndex], style_red)
        elif int(data[line][qvTrimColIndex]) < 30:
            worksheet.write(line + 1, qvTrimColIndex + 1, data[line][qvTrimColIndex], style_orange)

        line += 1


#####  Fonctions Log  #####

def exportConsoleToLogStart():
    global path_log, logout
    # Export console to log file
    # mode : 'r' = read, 'w' = write, 'a'= append

    logout = open(path_log, 'a')
    sys.stdout = logout
    sys.stderr = logout


def exportConsoleToLogStop():
    global logout
    # Restore initial values and save

    sys.stdout = sys.__stdout__
    sys.stderr = sys.__stderr__

    logout.close()


#####  Fonctions Analyse et Manipulation des Sequences  #####

def convertAb1ToFastq():
    global outputFolder, path_filelist_ab1, filelist_nopos, pos, loop

    n = 0
    while n < loop:
        SeqIO.convert(path_filelist_ab1[n], 'abi', outputFolder+'/'+filelist_nopos[n]+'_'+pos[n]+'.fastq', 'fastq-sanger')
        n = n+1

    print("Converted " + str(n) + "/" + str(len(path_filelist_ab1)) + " ab1 to fastq files")
    print('')
    exportConsoleToLogStart()
    print('##########  Convert ab1 > fastq  ##########')
    print("Converted " + str(n) + "/" + str(len(path_filelist_ab1)) + " ab1 to fastq files")
    print('')
    exportConsoleToLogStop()


def reverseComplement():
    global outputFolder, filelist_nopos_R, pos_R, loop_R

    n = 0
    while n < loop_R:
        records = (rec.reverse_complement(id=rec.id + "_RC", description="Reverse complement")
        for rec in SeqIO.parse(outputFolder+'/'+filelist_nopos_R[n] + '_' + pos_R[n] + '.fastq', 'fastq-sanger') if len(rec) < 700)
        SeqIO.write(records, outputFolder+'/'+filelist_nopos_R[n] + '_' + pos_R[n] + '_Rev_C' + '.fastq', 'fastq-sanger')
        n = n + 1

    print("Converted " + str(n) + "/" + str(len(filelist_nopos_R)) + " Seq_R to Reverse complement")
    print('')
    exportConsoleToLogStart()
    print('##########  Reverse Complement of Seq_R  ##########')
    print("Converted " + str(n) + "/" + str(len(filelist_nopos_R)) + " Seq_R to Reverse complement")
    print('')
    exportConsoleToLogStop()


def trimmingSeq_FileSeqObjListCreation():
    global outputFolder, filelist_nopos_F, filelist_nopos_R, pos_F, pos_R, loop_F, loop_R, fileSeqObjsList, fileSeqObjsErrorList

    nbBaseTrim = 10  # nb de bases pour le calcul du seuil
    # TODO: Gérer les seg de mauvaise qualité => pb de trimming. Etre plus restrictif
    seuilTrim = nbBaseTrim * 30  # valeur/base '25' modifiable
    n = 0
    # fileSeqObjsList = []
    scorePhred = 0
    while n < loop_F:
        for record in SeqIO.parse(outputFolder+'/'+filelist_nopos_F[n] + '_' + pos_F[n] + '.fastq', 'fastq-sanger'):
            try:
                # calculation of start position #
                j = 0
                scorePhred = 0
                while scorePhred < seuilTrim:
                    i = 0
                    scorePhred = 0
                    while i < nbBaseTrim:
                        scorePhred = scorePhred + record.letter_annotations["phred_quality"][i + j]
                        i = i + 1
                    j = j + 1
                startSeq = j + nbBaseTrim

                # calculation of end position #
                j = len(record.seq) - nbBaseTrim - 1
                scorePhred = 0
                while scorePhred < seuilTrim:
                    i = 0
                    scorePhred = 0
                    while i < nbBaseTrim:
                        scorePhred = scorePhred + record.letter_annotations["phred_quality"][i + j]
                        i = i + 1
                    j = j - 1
                endSeq = j

                trimSeq = record[startSeq:endSeq]
                trimSeq.id = record.id + "_Trimmed"
                trimSeq.description = record.description + "_Trimmed"
                SeqIO.write(trimSeq, outputFolder + '/' + filelist_nopos_F[n] + '_' + pos_F[n] + '_Trimmed' + '.fastq','fastq-sanger')

                #######################################################################################################
                # FileSeqObject Attributes #
                fileName = filelist_nopos_F[n]
                well = pos_F[n]
                id = filelist_nopos_F[n].split('-')[0]
                family = getFamilyFromID(id)
                gene = filelist_nopos_F[n].split('-')[1]
                exon = getExonFromFileName(fileName)
                lengthSeq = len(record)
                lengthSeqTrimm = len(trimSeq)

                # Average QV #
                qvRecord = 0
                for qv in record.letter_annotations["phred_quality"]:
                    qvRecord = qvRecord + qv
                qvRecord = round(qvRecord / len(record.seq), 1)
                qvTrimSeq = 0
                for qv in trimSeq.letter_annotations["phred_quality"]:
                    qvTrimSeq = qvTrimSeq + qv
                qvTrimSeq = round(qvTrimSeq / len(trimSeq.seq), 1)

                fileSeqObjsList.append(FileSeqObj(fileName=fileName, well=well, patientID=id, family=family, gene=gene,
                                                  exon=exon, transcriptID='', trimmStart=startSeq,
                                                  trimmStop=endSeq, lengthSeq=lengthSeq, lengthSeqTrimm=lengthSeqTrimm,
                                                  avgQV=qvRecord, avgQVTrimm=qvTrimSeq, commentary='',
                                                  path=outputFolder + '/' + filelist_nopos_F[n] + '_' + pos_F[n] + '_Trimmed' + '.fastq'))
                #######################################################################################################

                exportConsoleToLogStart()
                print(record.id)
                print('length:', len(record.seq))
                print('length trimmed:', len(trimSeq.seq))
                print('')
                exportConsoleToLogStop()

            except:
                #######################################################################################################
                # FileSeqObject Attributes #
                fileName = filelist_nopos_F[n]
                well = pos_F[n]
                id = filelist_nopos_F[n].split('-')[0]
                family = getFamilyFromID(id)
                try:
                    gene = filelist_nopos_F[n].split('-')[1]
                    exon = getExonFromFileName(fileName)
                except:
                    lengthSeq = len(record)
                    qvRecord = 0
                    for qv in record.letter_annotations["phred_quality"]:
                        qvRecord = qvRecord + qv
                    qvRecord = round(qvRecord / len(record.seq), 1)
                    commentary = 'Wrong Name of file. Please refer to the beginning of python file.'
                    fileSeqObjsErrorList.append(FileSeqObj(fileName=fileName, well=well, patientID=id, family=family,
                                                           gene=gene, exon=exon, transcriptID='', trimmStart='',
                                                           trimmStop='', lengthSeq=lengthSeq,
                                                           lengthSeqTrimm=0,
                                                           avgQV=qvRecord, avgQVTrimm=0, commentary=commentary,
                                                           path=''))

                lengthSeq = len(record)
                lengthSeqTrimm = 0
                commentary = 'Trimming error probably due to low quality sequence | '

                # Average QV #
                qvRecord = 0
                for qv in record.letter_annotations["phred_quality"]:
                    qvRecord = qvRecord + qv
                qvRecord = round(qvRecord / len(record.seq), 1)

                fileSeqObjsErrorList.append(FileSeqObj(fileName=fileName, well=well, patientID=id, family=family,
                                                       gene=gene, exon=exon, transcriptID='', trimmStart='',
                                                       trimmStop='', lengthSeq=lengthSeq, lengthSeqTrimm=lengthSeqTrimm,
                                                       avgQV=qvRecord, avgQVTrimm=0, commentary=commentary, path=''))
                #######################################################################################################

                exportConsoleToLogStart()
                print(record.id)
                print('length:', len(record.seq))
                print('length trimmed:', 0)
                print('')
                exportConsoleToLogStop()

        n = n + 1

    n = 0
    while n < loop_R:
        for record in SeqIO.parse(outputFolder+'/'+filelist_nopos_R[n] + '_' + pos_R[n] + '_Rev_C' + '.fastq', 'fastq-sanger'):
            try:
                # calculation of start position #
                j = 0
                scorePhred = 0
                while scorePhred < seuilTrim:
                    i = 0
                    scorePhred = 0
                    while i < nbBaseTrim:
                        scorePhred = scorePhred + record.letter_annotations["phred_quality"][i + j]
                        i = i + 1
                    j = j + 1
                startSeq = j + nbBaseTrim

                # calculation of end position #
                j = len(record.seq) - nbBaseTrim - 1
                scorePhred = 0
                while scorePhred < seuilTrim:
                    i = 0
                    scorePhred = 0
                    while i < nbBaseTrim:
                        scorePhred = scorePhred + record.letter_annotations["phred_quality"][i + j]
                        i = i + 1
                    j = j - 1
                endSeq = j

                trimSeq = record[startSeq:endSeq]
                trimSeq.id = record.id + "_Trimmed"
                trimSeq.description = record.description + "_Trimmed"
                SeqIO.write(trimSeq, outputFolder + '/' + filelist_nopos_R[n] + '_' + pos_R[n] + '_Rev_C_Trimmed' + '.fastq', 'fastq-sanger')

                #######################################################################################################
                # FileSeqObject Attributes #
                fileName = filelist_nopos_R[n]
                well = pos_R[n]
                id = filelist_nopos_R[n].split('-')[0]
                family = getFamilyFromID(id)
                gene = filelist_nopos_R[n].split('-')[1]
                exon = getExonFromFileName(fileName)
                lengthSeq = len(record)
                lengthSeqTrimm = len(trimSeq)

                # Average QV #
                qvRecord = 0
                for qv in record.letter_annotations["phred_quality"]:
                    qvRecord = qvRecord + qv
                qvRecord = round(qvRecord / len(record.seq), 1)
                qvTrimSeq = 0
                for qv in trimSeq.letter_annotations["phred_quality"]:
                    qvTrimSeq = qvTrimSeq + qv
                qvTrimSeq = round(qvTrimSeq / len(trimSeq.seq), 1)

                fileSeqObjsList.append(FileSeqObj(fileName=fileName, well=well, patientID=id, family=family, gene=gene,
                                                  exon=exon, transcriptID='', trimmStart=startSeq,
                                                  trimmStop=endSeq, lengthSeq=lengthSeq, lengthSeqTrimm=lengthSeqTrimm,
                                                  avgQV=qvRecord, avgQVTrimm=qvTrimSeq, commentary='',
                                                  path=outputFolder + '/' + filelist_nopos_R[n] + '_' + pos_R[n] + '_Rev_C_Trimmed' + '.fastq'))

                #######################################################################################################

                exportConsoleToLogStart()
                print(record.id)
                print('length:', len(record.seq))
                print('length trimmed:', len(trimSeq.seq))
                print('')
                exportConsoleToLogStop()

            except:
                #######################################################################################################
                # FileSeqObject Attributes #
                fileName = filelist_nopos_R[n]
                well = pos_R[n]
                id = filelist_nopos_R[n].split('-')[0]
                family = getFamilyFromID(id)
                try:
                    gene = filelist_nopos_R[n].split('-')[1]
                    exon = getExonFromFileName(fileName)
                except:
                    lengthSeq = len(record)
                    qvRecord = 0
                    for qv in record.letter_annotations["phred_quality"]:
                        qvRecord = qvRecord + qv
                    qvRecord = round(qvRecord / len(record.seq), 1)
                    commentary = 'Wrong Name of file. Please refer to the beginning of python file.'
                    fileSeqObjsErrorList.append(FileSeqObj(fileName=fileName, well=well, patientID=id, family=family,
                                                           gene=gene, exon=exon, transcriptID='', trimmStart='',
                                                           trimmStop='', lengthSeq=lengthSeq,
                                                           lengthSeqTrimm=0,
                                                           avgQV=qvRecord, avgQVTrimm=0, commentary=commentary,
                                                           path=''))

                lengthSeq = len(record)
                lengthSeqTrimm = 0
                commentary = 'Trimming error probably due to low quality sequence | '

                # Average QV #
                qvRecord = 0
                for qv in record.letter_annotations["phred_quality"]:
                    qvRecord = qvRecord + qv
                qvRecord = round(qvRecord / len(record.seq), 1)

                fileSeqObjsErrorList.append(FileSeqObj(fileName=fileName, well=well, patientID=id, family=family,
                                                       gene=gene, exon=exon, transcriptID='', trimmStart='',
                                                       trimmStop='', lengthSeq=lengthSeq, lengthSeqTrimm=lengthSeqTrimm,
                                                       avgQV=qvRecord, avgQVTrimm=0, commentary=commentary, path=''))

                #######################################################################################################

                exportConsoleToLogStart()
                print(record.id)
                print('length:', len(record.seq))
                print('length trimmed:', 0)
                print('')
                exportConsoleToLogStop()

        n = n + 1

    nbTrimmed = glob.glob(outputFolder + '/*Trimmed.fastq')
    nbTrimmed = len(nbTrimmed)
    print("Trimmed " + str(nbTrimmed) + "/" + str(len(filelist_ab1)) + " sequences")
    print('')
    exportConsoleToLogStart()
    print("Trimmed " + str(nbTrimmed) + "/" + str(len(filelist_ab1)) + " sequences")
    print('')
    exportConsoleToLogStop()

    return fileSeqObjsList, fileSeqObjsErrorList


def alignerGlobalParameters():

    aligner = Align.PairwiseAligner()
    aligner.mode = 'global'
    aligner.match_score = 2
    aligner.mismatch_score = -1
    aligner.open_gap_score = -10
    aligner.extend_gap_score = -0.5
    aligner.target_end_gap_score = 0.0
    aligner.query_end_gap_score = 0.0

    return aligner


def aligner2Seq(targetSeq, querySeq, aligner):
    # Pairwise alignment of 2 sequences

    score = aligner.score(targetSeq, querySeq)
    alignments = aligner.align(targetSeq, querySeq)
    alignment = alignments[0]
    print('Score:', score)
    exportConsoleToLogStart()
    print('Score:', score)
    exportConsoleToLogStop()

    return alignment


def alignmentAnalyzerGene(alignment, targetSeq, fileObj):
    global variantsObjList, filelist_nopos_F, filelist_nopos_R

    # Split alignment in target, query and result #
    alignmentStr = str(alignment)
    sautDeLigne = []
    i = 0
    while i < len(alignmentStr):
        if alignmentStr[i] == '\n':
            sautDeLigne.append(i)
        i = i + 1

    alignTarget = alignmentStr[:sautDeLigne[0]]
    alignQuery = alignmentStr[sautDeLigne[1]+1:sautDeLigne[2]]
    alignResult = alignmentStr[sautDeLigne[0]+1:sautDeLigne[1]]
    alignTarget_Length = len(alignTarget)

    print('Analyze Gene:')
    exportConsoleToLogStart()
    print('Analyze Gene:')

    # Cut target, query and result by aligned positions #
    posAlignment = str(alignment.aligned).replace(('('), '')
    posAlignment = posAlignment.replace(')', '')
    posAlignment = posAlignment.replace(' ', '')
    posAlignment = posAlignment.split(',')
    while '' in posAlignment:
        posAlignment.remove('')
    posAlignmentCorrected = []
    for pos in posAlignment:
        posAlignmentCorrected.append(int(pos) + 1)
    print('Split:', posAlignmentCorrected)

    if len(posAlignment) == 4:
        start = int(posAlignment[0])
        stop = start + (int(posAlignment[3]) - int(posAlignment[2]))
        alignTarget = alignTarget[start:stop]
        alignQuery = alignQuery[start:stop]
        alignResult = alignResult[start:stop]

        print('Reference:', alignTarget)
        print('Query:    ', alignQuery)
        print('Result:   ', alignResult)

    else:
        start = 0
        i = 0
        while i < len(alignResult):
            if alignResult[i] != '-':
                start = i
                break
            i = i + 1
        stop = 0
        i = len(alignResult) - 1
        while i >= 0:
            if alignResult[i] != '-':
                stop = i + 1
                break
            i = i - 1
        alignTarget = alignTarget[start:stop]
        alignQuery = alignQuery[start:stop]
        alignResult = alignResult[start:stop]
        print('Reference:', alignTarget)
        print('Query:    ', alignQuery)
        print('Result:   ', alignResult)

    # FileSeqObject Attributes #
    startTrimm = fileObj.trimmStart

    # Search variants #
    i = 0
    pos = start
    while i < len(alignResult):
        if alignResult[i] != '|':
            ref = alignTarget[i]
            alt = alignQuery[i]

            posG = pos + 1
            if fileObj.fileName in filelist_nopos_F:
                posAb1 = posG - start + startTrimm
            else:
                posAb1 = fileObj.lengthSeq - (posG - start + startTrimm + 1)

            variantsObjList.append(Variant(fileName=fileObj.fileName, well=fileObj.well,
                                           patientID=fileObj.patientID, family=fileObj.family, gene=fileObj.gene,
                                           exon=fileObj.exon, annotation='', alleleStatus='', exonic='No',
                                           posAb1=posAb1, posG=posG, refG=ref, altG=alt, posC='', refC='', altC='',
                                           posP='', refP='', altP='', commentary='', alignC=''))
            print('Pos Gene:', posG)
            print('Ref:', ref)
            print('Alt:', alt)

        pos = pos + 1
        i = i + 1
    exportConsoleToLogStop()
    print('done')

    # Test INDELS #
    if '-' in alignQuery:
        fileObj.commentary = fileObj.commentary + 'Homo Del to check on gene | '
    if '-' in alignTarget:
        fileObj.commentary = fileObj.commentary + 'Homo Ins to check on gene | '


def alignmentAnalyzerExon(alignment, targetSeq, fileObj):
    global variantsObjListCDS

    # Split alignment in target, query and result #
    alignmentStr = str(alignment)
    sautDeLigne = []
    i = 0
    while i < len(alignmentStr):
        if alignmentStr[i] == '\n':
            sautDeLigne.append(i)
        i = i + 1

    alignTarget = alignmentStr[:sautDeLigne[0]]
    alignQuery = alignmentStr[sautDeLigne[1]+1:sautDeLigne[2]]
    alignResult = alignmentStr[sautDeLigne[0]+1:sautDeLigne[1]]
    alignTarget_Length = len(alignTarget)

    print('Analyze Transcript:')
    exportConsoleToLogStart()
    print('Analyze Transcript:')

    # Cut target, query and result by aligned positions #
    posAlignment = str(alignment.aligned).replace('(', '')
    posAlignment = posAlignment.replace(')', '')
    posAlignment = posAlignment.replace(' ', '')
    posAlignment = posAlignment.split(',')
    while '' in posAlignment:
        posAlignment.remove('')
    posAlignmentCorrected = []
    for pos in posAlignment:
        posAlignmentCorrected.append(int(pos) + 1)
    print('Split:', posAlignmentCorrected)

    if len(posAlignment) == 4:
        start = int(posAlignment[0])
        stop = start + (int(posAlignment[3]) - int(posAlignment[2]))
        alignTarget = alignTarget[start:stop]
        alignQuery = alignQuery[start:stop]
        alignResult = alignResult[start:stop]
        print('Reference:', alignTarget)
        print('Query:    ', alignQuery)
        print('Result:   ', alignResult)

    else:
        start = 0
        i = 0
        while i < len(alignResult):
            if alignResult[i] != '-':
                start = i
                break
            i = i + 1
        stop = 0
        i = len(alignResult) - 1
        while i >= 0:
            if alignResult[i] != '-':
                stop = i + 1
                break
            i = i - 1
        alignTarget = alignTarget[start:stop]
        alignQuery = alignQuery[start:stop]
        alignResult = alignResult[start:stop]
        print('Reference:', alignTarget)
        print('Query:    ', alignQuery)
        print('Result:   ', alignResult)

    # Search variants #
    i = 0
    pos = start
    # variantsObjList = []
    while i < len(alignResult):
        if alignResult[i] != '|':
            ref = alignTarget[i]
            alt = alignQuery[i]

            posC = pos + 1
            variantsObjListCDS.append(Variant(fileName=fileObj.fileName, well=fileObj.well,
                                              patientID=fileObj.patientID, family=fileObj.family, gene=fileObj.gene,
                                              exon=fileObj.exon, annotation='', alleleStatus='', exonic='Yes',
                                              posAb1='', posG='', refG='', altG='', posC=posC, refC=ref, altC=alt,
                                              posP='', refP='', altP='', commentary='', alignC=alignment))
            print('Pos Transcript:', posC)
            print('Ref:', ref)
            print('Alt:', alt)

        pos = pos + 1
        i = i + 1
    exportConsoleToLogStop()
    print('done')

    # Test INDELS #
    if '-' in alignQuery:
        fileObj.commentary = fileObj.commentary + 'Homo Del to check on transcript | '
    if '-' in alignTarget:
        fileObj.commentary = fileObj.commentary + 'Homo Ins to check on transcript | '

"""
def alignmentAnalyzerGene(alignment, targetSeq, fileObj):
    global variantsObjList, filelist_nopos_F, filelist_nopos_R

    # Split alignment in target, query and result #
    alignmentStr = str(alignment)
    sautDeLigne = []
    i = 0
    while i < len(alignmentStr):
        if alignmentStr[i] == '\n':
            sautDeLigne.append(i)
        i = i + 1

    alignTarget = alignmentStr[:sautDeLigne[0]]
    alignQuery = alignmentStr[sautDeLigne[1]+1:sautDeLigne[2]]
    alignResult = alignmentStr[sautDeLigne[0]+1:sautDeLigne[1]]
    alignTarget_Length = len(alignTarget)

    print('Analyze Gene:')
    exportConsoleToLogStart()
    print('Analyze Gene:')

    # Cut target, query and result by aligned positions #
    posAlignment = str(alignment.aligned).replace(('('), '')
    posAlignment = posAlignment.replace(')', '')
    posAlignment = posAlignment.replace(' ', '')
    posAlignment = posAlignment.split(',')
    while '' in posAlignment:
        posAlignment.remove('')
    posAlignmentCorrected = []
    for pos in posAlignment:
        posAlignmentCorrected.append(int(pos) + 1)
    print('Split:', posAlignmentCorrected)

    if len(posAlignment) == 4:
        start = int(posAlignment[0])
        stop = start + (int(posAlignment[3]) - int(posAlignment[2]))
        alignTarget = alignTarget[start:stop]
        alignQuery = alignQuery[start:stop]
        alignResult = alignResult[start:stop]

        print('Reference:', alignTarget)
        print('Query:    ', alignQuery)
        print('Result:   ', alignResult)

    else:
        start = 0
        i = 0
        while i < len(alignResult):
            if alignResult[i] != '-':
                start = i
                break
            i = i + 1
        stop = 0
        i = len(alignResult) - 1
        while i >= 0:
            if alignResult[i] != '-':
                stop = i + 1
                break
            i = i - 1
        alignTarget = alignTarget[start:stop]
        alignQuery = alignQuery[start:stop]
        alignResult = alignResult[start:stop]
        print('Reference:', alignTarget)
        print('Query:    ', alignQuery)
        print('Result:   ', alignResult)

    # FileSeqObject Attributes #
    startTrimm = fileObj.trimmStart

    # Search variants #
    i = 0
    pos = start
    while i < len(alignResult):
        if alignResult[i] != '|':
            ref = alignTarget[i]
            alt = alignQuery[i]
            if alignTarget_Length == len(targetSeq):
                posG = pos + 1
                if fileObj.fileName in filelist_nopos_F:
                    posAb1 = posG - start + startTrimm
                else:
                    posAb1 = fileObj.lengthSeq - (posG - start + startTrimm + 1)

                variantsObjList.append(Variant(fileName=fileObj.fileName, well=fileObj.well,
                                               patientID=fileObj.patientID, family=fileObj.family, gene=fileObj.gene,
                                               exon=fileObj.exon, annotation='', alleleStatus='', exonic='No',
                                               posAb1=posAb1, posG=posG, refG=ref, altG=alt, posC='', refC='', altC='',
                                               posP='', refP='', altP='', commentary='', alignC=''))
                print('Pos Gene:', posG)
                print('Ref:', ref)
                print('Alt:', alt)
        pos = pos + 1
        i = i + 1
    exportConsoleToLogStop()
    print('done')

    # Test INDELS #
    if '-' in alignQuery:
        fileObj.commentary = fileObj.commentary + 'Homo Del to check on gene | '
    if '-' in alignTarget:
        fileObj.commentary = fileObj.commentary + 'Homo Ins to check on gene | '


def alignmentAnalyzerExon(alignment, targetSeq, fileObj):
    global variantsObjListCDS

    # Split alignment in target, query and result #
    alignmentStr = str(alignment)
    sautDeLigne = []
    i = 0
    while i < len(alignmentStr):
        if alignmentStr[i] == '\n':
            sautDeLigne.append(i)
        i = i + 1

    alignTarget = alignmentStr[:sautDeLigne[0]]
    alignQuery = alignmentStr[sautDeLigne[1]+1:sautDeLigne[2]]
    alignResult = alignmentStr[sautDeLigne[0]+1:sautDeLigne[1]]
    alignTarget_Length = len(alignTarget)

    print('Analyze Transcript:')
    exportConsoleToLogStart()
    print('Analyze Transcript:')

    # Cut target, query and result by aligned positions #
    posAlignment = str(alignment.aligned).replace('(', '')
    posAlignment = posAlignment.replace(')', '')
    posAlignment = posAlignment.replace(' ', '')
    posAlignment = posAlignment.split(',')
    while '' in posAlignment:
        posAlignment.remove('')
    posAlignmentCorrected = []
    for pos in posAlignment:
        posAlignmentCorrected.append(int(pos) + 1)
    print('Split:', posAlignmentCorrected)

    if len(posAlignment) == 4:
        start = int(posAlignment[0])
        stop = start + (int(posAlignment[3]) - int(posAlignment[2]))
        alignTarget = alignTarget[start:stop]
        alignQuery = alignQuery[start:stop]
        alignResult = alignResult[start:stop]
        print('Reference:', alignTarget)
        print('Query:    ', alignQuery)
        print('Result:   ', alignResult)

    else:
        start = 0
        i = 0
        while i < len(alignResult):
            if alignResult[i] != '-':
                start = i
                break
            i = i + 1
        stop = 0
        i = len(alignResult) - 1
        while i >= 0:
            if alignResult[i] != '-':
                stop = i + 1
                break
            i = i - 1
        alignTarget = alignTarget[start:stop]
        alignQuery = alignQuery[start:stop]
        alignResult = alignResult[start:stop]
        print('Reference:', alignTarget)
        print('Query:    ', alignQuery)
        print('Result:   ', alignResult)

    # Search variants #
    i = 0
    pos = start
    # variantsObjList = []
    while i < len(alignResult):
        if alignResult[i] != '|':
            ref = alignTarget[i]
            alt = alignQuery[i]
            if alignTarget_Length == len(targetSeq):
                posC = pos + 1
                variantsObjListCDS.append(Variant(fileName=fileObj.fileName, well=fileObj.well,
                                                  patientID=fileObj.patientID, family=fileObj.family, gene=fileObj.gene,
                                                  exon=fileObj.exon, annotation='', alleleStatus='', exonic='Yes',
                                                  posAb1='', posG='', refG='', altG='', posC=posC, refC=ref, altC=alt,
                                                  posP='', refP='', altP='', commentary='', alignC=alignment))
                print('Pos Transcript:', posC)
                print('Ref:', ref)
                print('Alt:', alt)
        pos = pos + 1
        i = i + 1
    exportConsoleToLogStop()
    print('done')

    # Test INDELS #
    if '-' in alignQuery:
        fileObj.commentary = fileObj.commentary + 'Homo Del to check on transcript | '
    if '-' in alignTarget:
        fileObj.commentary = fileObj.commentary + 'Homo Ins to check on transcript | '
"""

def coverageVerification(alignment, fileObj, shift):
    global dict_StartStopCDSList
    # shift=0: exon uniquement
    # shift=2: exon+splice

    # TODO: revoir l'utilisation de incompleteCoverage(inutile?). Basculer sur consensus si pas total
    # Get aligned position on gene #
    posAlignment = str(alignment.aligned).replace('(', '').replace(')', '')
    posAlignment = posAlignment.replace(' ', '')
    posAlignment = posAlignment.split(',')
    while '' in posAlignment:
        posAlignment.remove('')

    # Coverage test #
    start = dict_StartStopCDSList[fileObj.gene][0]
    stop = dict_StartStopCDSList[fileObj.gene][1]
    totalCoverage = 0  # Couverture totale de l'exon
    partialCoverage = 0  # Couverture incomplète sur une extrémité de l'exon
    partialCoveragePos = []
    incompleteCoverage = 0  # Couverture incomplète sur les 2 extrémités de l'exon
    incompleteCoveragePos = []
    if len(posAlignment) == 4:
        i = 0
        while i < len(start):
            if int(posAlignment[0]) < (int(start[i]) - shift) and int(posAlignment[1]) > (int(stop[i]) + shift):
                totalCoverage = totalCoverage + 1
            elif int(posAlignment[0]) > (int(start[i]) - shift) and int(posAlignment[0]) < (int(stop[i]) + shift):
                partialCoverage = partialCoverage + 1
                partialCoveragePos.append(start[i] + ':' + stop[i])
            elif int(posAlignment[1]) > (int(start[i]) - shift) and int(posAlignment[1]) < (int(stop[i]) + shift):
                partialCoverage = partialCoverage + 1
                partialCoveragePos.append(start[i] + ':' + stop[i])
            elif int(posAlignment[0]) > (int(start[i]) - shift) and int(posAlignment[1]) < (int(stop[i]) + shift):
                incompleteCoverage = incompleteCoverage + 1
                incompleteCoveragePos.append(start[i] + ':' + stop[i])
            i = i + 1
    else:
        # TODO: A revoir en cas de indels pb de décalage
        i = 0
        while i < len(start):
            if int(posAlignment[0]) < (int(start[i]) - shift) and (int(posAlignment[0]) + int(posAlignment[-1])) > (int(stop[i]) + shift):
                totalCoverage = totalCoverage + 1
            elif int(posAlignment[0]) > (int(start[i]) - shift) and int(posAlignment[0]) < (int(stop[i]) + shift):
                partialCoverage = partialCoverage + 1
                partialCoveragePos.append(start[i] + ':' + stop[i])
            elif (int(posAlignment[0]) + int(posAlignment[-1])) > (int(start[i]) - shift) and (int(posAlignment[0]) + int(posAlignment[-1])) < (int(stop[i]) + shift):
                partialCoverage = partialCoverage + 1
                partialCoveragePos.append(start[i] + ':' + stop[i])
            elif int(posAlignment[0]) > (int(start[i]) - shift) and (int(posAlignment[0]) + int(posAlignment[-1])) < (int(stop[i]) + shift):
                incompleteCoverage = incompleteCoverage + 1
                incompleteCoveragePos.append(start[i] + ':' + stop[i])
            i = i + 1

    # Delete doubles from 'partialCoveragePos' generated by 'incompleteCoverage' #
    for elt in incompleteCoveragePos:
        partialCoveragePos.remove(elt)

    # Make comment #
    commentary = ''
    for pos in partialCoveragePos:
        commentary = commentary + 'Partial coverage for region [' + pos + '] | '
    for pos in incompleteCoveragePos:
        commentary = commentary + 'Incomplete coverage for region [' + pos + '] | '
    fileObj.commentary = fileObj.commentary + commentary

    exportConsoleToLogStart()
    print('Exon_totalCoverage:', totalCoverage)
    print('Exon_partialCoverage:', partialCoverage + incompleteCoverage)
    # print('incompleteCoverage:', incompleteCoverage)
    if commentary != '':
        print(commentary)
    exportConsoleToLogStop()
    print('Coverage Test:')
    print('done')


def querySplitExon(alignment, gene):

    # Split alignment in target, query and result #
    alignmentStr = str(alignment)
    sautDeLigne = []
    i = 0
    while i < len(alignmentStr):
        if alignmentStr[i] == '\n':
            sautDeLigne.append(i)
        i = i + 1

    alignQuery = alignmentStr[sautDeLigne[1]+1:sautDeLigne[2]]

    # Cut query by exon position #
    start = dict_StartStopCDSList[gene][0]
    stop = dict_StartStopCDSList[gene][1]
    alignQueryExon = ''
    i = 0
    while i < len(start):
        exon = alignQuery[int(start[i]):int(stop[i])]
        if 'A' in exon or 'T' in exon or 'G' in exon or 'C' in exon:
            alignQueryExon = alignQueryExon + exon
            # print(alignQueryExon)
        i = i + 1
    queryExon = alignQueryExon.replace('-', '')

    return queryExon


def variantAnalyzer(variantObj):
    global variantsObjListCDS

    # Split alignment in target, query and result #
    alignmentStr = str(variantObj.alignC)
    sautDeLigne = []
    i = 0
    while i < len(alignmentStr):
        if alignmentStr[i] == '\n':
            sautDeLigne.append(i)
        i = i + 1

    alignTarget = alignmentStr[:sautDeLigne[0]]
    alignQuery = alignmentStr[sautDeLigne[1] + 1:sautDeLigne[2]]

    # Variant analysis and amino acid determination #
    codonRef = ''
    codonAlt = ''
    if variantObj.posC != '':
        translationFrame = int(variantObj.posC) % 3
        if translationFrame == 0:
            variantObj.posP = floor(int(variantObj.posC) / 3)  # Arrondi à l'entier inf
        else:
            variantObj.posP = floor(int(variantObj.posC) / 3 + 1)  # Arrondi à l'entier inf

        if translationFrame == 1:
            # print('Translation Frame: +1')
            pos = -1
            while len(codonRef) < 3:
                if alignTarget[variantObj.posC] != '-':
                    codonRef = codonRef + alignTarget[variantObj.posC + pos]
                pos = pos + 1
            pos = -1
            while len(codonAlt) < 3:
                codonAlt = codonAlt + alignQuery[variantObj.posC + pos]
                pos = pos + 1
            print('File:', variantObj.fileName)
            print('Codon Ref:            ', codonRef)
            print('Codon Alt Ambiguous:  ', codonAlt)
        elif translationFrame == 2:
            # print('Translation Frame: +2')
            pos = -2
            while len(codonRef) < 3:
                if alignTarget[variantObj.posC] != '-':
                    codonRef = codonRef + alignTarget[variantObj.posC + pos]
                pos = pos + 1
            pos = -2
            while len(codonAlt) < 3:
                codonAlt = codonAlt + alignQuery[variantObj.posC + pos]
                pos = pos + 1
            print('File:', variantObj.fileName)
            print('Codon Ref:            ', codonRef)
            print('Codon Alt Ambiguous:  ', codonAlt)
        elif translationFrame == 0:
            # print('Translation Frame: +3')
            pos = -3
            while len(codonRef) < 3:
                if alignTarget[variantObj.posC] != '-':
                    codonRef = codonRef + alignTarget[variantObj.posC + pos]
                pos = pos + 1
            pos = -3
            while len(codonAlt) < 3:
                codonAlt = codonAlt + alignQuery[variantObj.posC + pos]
                pos = pos + 1
            print('File:', variantObj.fileName)
            print('Codon Ref:            ', codonRef)
            print('Codon Alt Ambiguous:  ', codonAlt)

        # TODO: Transformer les bases IUPAC Ambiguous en Unambiguous. Méthode fonctionnelle.
        # TODO: Essayer de trouver une fonction plus simple.
        codonRefList = list(codonRef)
        codonAltList = list(codonAlt)
        i = 0
        while i < len(codonAltList):
            if codonAltList[i] != '-':
                ambiguousBase = ambiguous_dna_values[codonAltList[i]]
                if len(ambiguousBase) > 1:
                    unambiguousBase = ambiguousBase.replace(codonRefList[i], '')
                    codonAltList[i] = unambiguousBase
            i = i + 1
        #print(codonAltList)
        codonAltUnambiguous = ''.join(codonAltList)
        print('Codon Alt Unambiguous:', codonAltUnambiguous)

        if '-' not in codonRef:
            aaRef = str(Seq(codonRef).translate())
            print('AA Ref:', aaRef)
            variantObj.refP = aaRef
        else:
            variantObj.refP = '?'

        if '-' not in codonAltUnambiguous:
            aaAlt = str(Seq(codonAltUnambiguous).translate())
            print('AA Alt:', aaAlt)
            variantObj.altP = aaAlt
        else:
            variantObj.altP = '?'
        """
        m = motifs.create([Seq(codonRef, IUPAC.ambiguous_dna), Seq(codonAlt, IUPAC.ambiguous_dna)])
        print(m.degenerate_consensus)
        """


def ambigousToUnambigousBase(ref, alt):

    # TODO: Revoir pour les cas ou ambiguous = 3 bases (1020BI001006-POLG-3R)
    unambiguousBase = ''
    if alt != '' and alt != '-':
        if alt != 'A' or alt != 'T' or alt != 'G' or alt != 'C':
            ambiguousBase = ambiguous_dna_values[alt]
            unambiguousBase = ambiguousBase.replace(ref, '')

    return unambiguousBase


def variantAnnotation(variantObj):

    unambigousBase = ambigousToUnambigousBase(variantObj.refC, variantObj.altC)
    if variantObj.posC != '':
        annotation = 'g.' + str(variantObj.posG) + variantObj.refG + '>' + variantObj.altG + ' | '
        annotation = annotation + 'c.' + str(variantObj.posC) + variantObj.refC + '>' + unambigousBase + ' | '
        if variantObj.refP == '?' or variantObj.altP == '?':
            annotation = annotation + 'p.?'
        else:
            annotation = annotation + 'p.' + variantObj.refP + str(variantObj.posP) + variantObj.altP
        # annotation = annotation + 'p.' + variantObj.refP + str(variantObj.posP) + variantObj.altP
        variantObj.annotation = annotation


def alleleStatusDetermination(variantObj):

    if variantObj.altG == '-' or variantObj.altG == 'A' or variantObj.altG == 'T' or variantObj.altG == 'G' or variantObj.altG == 'C':
        variantObj.alleleStatus = ' 1/1 '
    else:
        variantObj.alleleStatus = ' 0/1 '


def consensus():
    global outputFolder, filelist_nopos_F, filelist_nopos_R, pos_F, pos_R, indexlist_dual_F, indexlist_dual_R, indexlist_single_F, indexlist_single_R

    n = 0
    loop = len(indexlist_dual_F)
    aligner = Align.PairwiseAligner()
    aligner.mode = 'global'
    aligner.match_score = 2
    aligner.mismatch_score = -1
    aligner.open_gap_score = -10
    aligner.extend_gap_score = -0.5
    aligner.target_end_gap_score = 0.0
    aligner.query_end_gap_score = 0.0
    while n < loop:
        seq_F = SeqIO.read(
            outputFolder+'/'+filelist_nopos_F[indexlist_dual_F[n]] + '_' + pos_F[indexlist_dual_F[n]] + '_Trimmed' + '.fastq',
            'fastq-sanger')
        seq_R = SeqIO.read(
            outputFolder+'/'+filelist_nopos_R[indexlist_dual_R[n]] + '_' + pos_R[indexlist_dual_R[n]] + '_Rev_C_Trimmed' + '.fastq',
            'fastq-sanger')
        seq_F = seq_F.seq
        seq_R = seq_R.seq
        # seq_F = "AAACCCTTTGGGGG"
        # seq_R = "AATTCGGAA"
        print(filelist_nopos_F[indexlist_dual_F[n]])
        print(filelist_nopos_R[indexlist_dual_R[n]])
        score = aligner.score(seq_F, seq_R)
        alignments = aligner.align(seq_F, seq_R)
        alignment = alignments[0]
        print('Score:', score)
        print(alignment.aligned)
        print(alignment)

        alignment = str(alignment)
        sautDeLigne = []
        i = 0
        while i < len(alignment):
            if alignment[i] == '\n':
                sautDeLigne.append(i)
            i = i + 1
        # print(sautDeLigne)

        alignment_F = alignment[:sautDeLigne[0]]
        alignment_R = alignment[sautDeLigne[1] + 1:sautDeLigne[2]]
        alignment_result = alignment[sautDeLigne[0] + 1:sautDeLigne[1]]

        print('')
        print(alignment_F)
        print('')
        print(alignment_R)
        print('')
        print(alignment_result)
        print('')
        # print(len(alignment_F))
        # print(len(alignment_R))
        # print(len(alignment_result))

        alignment_consensus = ""
        j = 0
        while j < len(alignment_result):
            if alignment_result[j] == '|':
                alignment_consensus = alignment_consensus + str(alignment_F[j])
            if alignment_result[j] == '-' and alignment_F[j] == '-':
                alignment_consensus = alignment_consensus + str(alignment_R[j])
            if alignment_result[j] == '-' and alignment_R[j] == '-':
                alignment_consensus = alignment_consensus + str(alignment_F[j])
            if alignment_result[j] == 'X' or alignment_result[j] == '.':
                if alignment_F[j] != 'A' and alignment_F[j] != 'T' and alignment_F[j] != 'G' and alignment_F[j] != 'C':
                    alignment_consensus = alignment_consensus + str(alignment_F[j])
                elif alignment_R[j] != 'A' and alignment_R[j] != 'T' and alignment_R[j] != 'G' and alignment_R[
                    j] != 'C':
                    alignment_consensus = alignment_consensus + str(alignment_R[j])
                else:
                    alignment_consensus = alignment_consensus + 'X'
            j = j + 1
        print(alignment_consensus)
        print('length consensus:', len(alignment_consensus))
        print('')

        seq_consensus = Seq(alignment_consensus)  # déclaration de l'objet Seq
        record = SeqRecord(seq_consensus)  # déclaration de l'objet SeqRecord à partir de l'objet Seq
        record.id = filelist_nopos_F[indexlist_dual_F[n]] + 'R_Consensus'
        record.name = filelist_nopos_F[indexlist_dual_F[n]] + 'R_Consensus'
        record.description = filelist_nopos_F[indexlist_dual_F[n]] + 'R_Consensus'
        SeqIO.write(record, outputFolder+'/'+filelist_nopos_F[indexlist_dual_F[n]] + 'R_' + pos_F[indexlist_dual_F[n]]
                    + pos_R[indexlist_dual_R[n]] + '_Consensus' + '.fa', 'fasta')
        print(record)
        print('')
        n = n + 1

    filelist_fa = glob.glob('./Output_Files/*.fa')
    print('' + str(len(filelist_fa)) + ' consensus sequence(s) created and ' + str(len(indexlist_single_F) + len(indexlist_single_R))
          + ' sequence(s) in single mode. (' + str(len(filelist_fa) * 2 + len(indexlist_single_F) + len(indexlist_single_R)) + '/'
          + str(len(filelist_ab1)) + ')')
    print('')


#####  Classes  #####

class FileSeqObj:

    def __init__(self, fileName, well, patientID, family, gene, exon, transcriptID, trimmStart, trimmStop, lengthSeq, lengthSeqTrimm, avgQV, avgQVTrimm, commentary, path):

        self.fileName = fileName
        self.well = well
        self.patientID = patientID
        self.family = family
        self.gene = gene
        self.exon = exon

        self.transcriptID = transcriptID

        self.trimmStart = trimmStart
        self.trimmStop = trimmStop
        self.lengthSeq = lengthSeq
        self.lengthSeqTrimm = lengthSeqTrimm

        self.avgQV = avgQV
        self.avgQVTrimm = avgQVTrimm

        self.commentary = commentary
        self.path = path

    # TODO: revoir la représentation de l'objet


class Variant:
    # TODO: Rajout des codons Ref et unambiguous
    def __init__(self, fileName, well, patientID, family, gene, exon, annotation, alleleStatus, exonic, posAb1, posG, refG, altG, posC, refC, altC, posP, refP, altP, commentary, alignC):

        self.fileName = fileName
        self.well = well
        self.patientID = patientID
        self.family = family
        self.gene = gene
        self.exon = exon

        self.annotation = annotation
        self.alleleStatus = alleleStatus
        self.exonic = exonic

        self.posAb1 = posAb1

        self.posG = posG
        self.refG = refG
        self.altG = altG

        self.posC = posC
        self.refC = refC
        self.altC = altC

        self.posP = posP
        self.refP = refP
        self.altP = altP

        self.commentary = commentary
        self.alignC = alignC

    # TODO: revoir la représentation de l'objet
    def __repr__(self):
        return 'Variant: file: {}, posG: {}, refG: {}, altG: {}, posC: {}, refC: {}, altC: {}, posP: {}, refP: {}, altP: {}, posAb1 {}'.format(
            self.fileName, self.posG, self.refG, self.altG, self.posC, self.refC, self.altC, self.posP, self.refP, self.altP, self.posAb1)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:  # BaseException est un niveau au dessus mais catch les SystemExit exception
        print('###########################################################     STOPPED WITH ERROR(S)     ###########################################################')
        print('')
        logger.error(e, exc_info = True)
        exportConsoleToLogStart()
        print('END:', datetime.today().strftime('%Y-%m-%d_%Hh%Mm%S'))
        exportConsoleToLogStop()
        os.system('pause')
