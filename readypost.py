#! python3
import time
import os
import sys
import openpyxl
import re
import hashlib
from pathlib import Path


def imgStart(val):
    img_start_regex = re.compile(r'<img.*src=\"images')
    img_start_mo = img_start_regex.search(val)
    if img_start_mo != None:
        img_start_mo = str(img_start_mo.group())
        val = val.replace(
            img_start_mo, '<img src="http://www.thenewamerican.com/images')
    return val


def imgEnd(val):
    img_end_regex = re.compile(r'((.jpg|png)\" .*)/>')
    img_end_mo = img_end_regex.search(val)
    if img_end_mo != None:
        img_end_mo = str(img_end_mo.group())
        if img_end_mo.find("jpg") != -1:
            val = val.replace(img_end_mo, '.jpg">')
        elif img_end_mo.find("png") != -1:
            val = val.replace(img_end_mo, '.png">')
    return val


def getHash(postid):
    the_hash = hashlib.md5(("Image" + str(postid)).encode('utf-8')).hexdigest()
    return the_hash


def imageCheck(val):
    check = ''
    img_check_regex = re.compile(r'<img.*(jpg|png)">')
    img_check_mo = img_check_regex.search(val)
    if img_check_mo != None:
        check = img_check_mo
    return check


#Vars
# cat = sys.argv[1]
cat = "154"
data_folder = Path("C:/Users/clemmer/Desktop/migAssets")
file_name = cat + "Content.xlsx"
file_to_open = data_folder / file_name
save_as = str(cat) + "Content2.xlsx"
save_path = data_folder / save_as

#Start
#---------
wb = openpyxl.load_workbook(file_to_open)
ws = wb.active


#get max row of excel sheet that contains data
maxRow = None
counter = 2  # start on row after header
while (counter < (ws.max_row + 1)):
    currentA = "A" + str(counter)
    valA = ws[currentA].value
    if valA == None:
        maxRow = counter
        break
    counter = counter + 1

if (maxRow == None):
    maxRow = counter - 1
    
print(maxRow)

# wb.save(save_path)
