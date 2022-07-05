# Python 2.7
# read csv file information and write into GDB feature field.
# create on: 2022-07-05 14:36:23

import csv
import arcpy

GDB = "D:\\ArcMap test\\GGITtest.gdb\\test_again"  # GDB path

GGITfile = open("D:\\ArcMap test\\GFIT_allitem_test.csv")  # csv file path

reader = csv.reader(GGITfile)  # read csv file
for column in reader:
    if 'Text' in column:  # Type: text
        texttype = column[0]  # field name
        stringlong = column[2]  # field length
        arcpy.AddField_management(GDB, texttype, "TEXT", "", "", stringlong, "", "NULLABLE")
        print (texttype + " field is added.")

    elif 'Double' in column:  # Type: double
        doubletype = column[0]  # field name
        arcpy.AddField_management(GDB, doubletype, "DOUBLE", "", "", "", "", "")
        print (doubletype + " field is added.")

    else:
        print ("Something wrong! Please check csv file.")

print ("Done!")
