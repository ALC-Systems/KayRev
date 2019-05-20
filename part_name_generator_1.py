#!/usr/bin/env python3


# address: python D:\01PythonScripts\kayrev_part_name_generator\part_name_generator_1.py
# test: python D:\01PythonScripts\kayrev_part_name_generator\part_name_generator_1.py 1 2 5 H 3

# author: Alex Liu, KayRev LLC
# alex.liu.carter@gmail.com
# 5/19/2019

# This script generates a list of part numbers. Inputs = Company Serial No. + Serial No. of Product + Total Number of Parts
# + Version (Latin letter) + Revision Number.

# the output should be in this format:
# 00X-00X-00X VX RX
# where 'X' stands for the variables.

import sys, xlwt

def construct_name(assemblyinfo):

    partnames = []
    companynum = assemblyinfo[0]
    productnum = assemblyinfo[1]
    numparts = assemblyinfo[2]     # this defines how many times the part number generation loop runs
    version = assemblyinfo[3]
    revision = assemblyinfo[4]

    companynum = addzero(companynum)
    productnum = addzero(productnum)

    for i in range (1,numparts+1):

        i = addzero(i)    # add zero to part number if applicable and turn it into string

        newpartname = str(companynum) + '-' + str(productnum) + '-' + i + ' ' + 'V' + str(version) + ' ' + 'R' + str(revision)
        partnames.append(newpartname)

    return partnames

# this function adds a zero for part numbers [1,9] and one zero for part numbers [10,99]

def addzero(num):

    if num < 10:
        num = '00' + str(num)
    elif num >= 10 and num < 100:
        num = '0' + str(num)
    else:
        num = str(num)

    return num

# this function writes each element in the list to a row in an excel sheet. They will be in column 1

def writetoexcel(list):

    wb = xlwt.Workbook()

    sheet1 = wb.add_sheet('Sheet 1')

    sheet1.write(0,0,'Part Number:')

    row = 2

    for i in list:
        sheet1.write(row, 0, str(i))
        row += 1

    wb.save('Part Numbers.xls')


if __name__ == '__main__':

# acquire variables from cmd

    companynum = int(sys.argv[1])
    productnum = int(sys.argv[2])
    numparts = int(sys.argv[3])
    version = sys.argv[4]
    revision = int(sys.argv[5])

# construct assembly info list

    assembly = [companynum, productnum, numparts, version, revision]

    finalnames = construct_name(assembly)

    writetoexcel(finalnames)
