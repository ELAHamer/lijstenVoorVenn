#-------------------------------------------------------------------------------
# Name:        
# Purpose:     Inlezen van alle xml bestanden
#              spectra eruit halen
#               in lijsten zetten              
#
# Author:      Ernst Hamer
#
# Created:     02-10-2017
# Copyright:   (c) Ernst Hamer 2016
#              Hogeschool van Arnhem en Nijmegen
#-------------------------------------------------------------------------------
import xlrd

def main():
    file = openfile()
    readfile(file)
    
def openfile():
    try:
        file = open("C:\\Users\\Ernst\\Desktop\\SUPPTAB_S03 Differentially regulated proteins in control (KR) 0-24 h after DEX treatment.xlsx","r")
        return file 
    except IOError:
        print("bestand niet gevonden")
        quit()
        
def readfile(file):
    print (file.nsheets)
    print (file.sheet_names())
    first_sheet = file.sheet_by_index(0) 
    print (first_sheet.row_values(0))
 

    cell = first_sheet.cell(0,0)
    print (cell)
    print (cell.value)
    print (first_sheet.row_slice(rowx=0,
                                start_colx=0,
                                end_colx=2))
main()