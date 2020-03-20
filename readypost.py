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
cat = sys.argv[1]
#cat = 560
data_folder = Path("C:/Users/clemmer/Desktop/migAssets")
file_name = str(cat) + "Content.xlsx"
file_to_open = data_folder / file_name
save_as = str(cat) + "Content2.xlsx"
save_path = data_folder / save_as

#Start
#---------
wb = openpyxl.load_workbook(file_to_open)
ws = wb.active


#get max row of excel sheet
maxRow = None
counter = 2  # start on row after header
while (counter < (ws.max_row + 1)):
    currentA = "A" + str(counter)
    valA = ws[currentA].value
    if valA == None:
        maxRow = int(counter)
        break
    counter = counter + 1

if maxRow == None:
    maxRow = ws.max_row + 1
print(maxRow)


#format post_excerpt images urls/tags
#loop through E and
#   remove leading before first <P on excerpts
#   remove everything between <img ... src="https:// ... .jpg|png"
counter = 2 #start on row after header
while ( counter < maxRow ):   
    currentG = "G" + str(counter)
    valG = ws[currentG].value
    valG = str(valG)
    currentE = "E" + str(counter)
    valE = ws[currentE].value
    valE = str(valE)
  
    if ( valG.find("VIDEO -") ) != -1:
        #remove valF style tags
        remove_portion = valE.split('<p')[0]
        valE = valE.replace(remove_portion, "")

        #wrap video links in a tags
        remove_portion = "{youtube}"
        valE = valE.replace(remove_portion, '<a href="')
        remove_portion = "{/youtube}"
        valE = valE.replace(remove_portion, '">')

        ws[currentE].value = valE
    else:
        #remove everything until first 'p' tag
        remove_portion = valG.split('<p')[0]
        valG = valG.replace(remove_portion, "")

        #replace img start on E
        valG = imgStart(valG)

        #replace img end on E
        valG = imgEnd(valG)
        ws[currentG].value = valG

        #replace img start on F
        valE = imgStart(valE)

        #replace img end on F
        valE = imgEnd(valE)
        ws[currentE].value = valE

    #concat the excerpt to the content
    ws[currentG].value = str(valG) + str(valE)
    
    #check post_content for an image
    valE = ws[currentE].value
    image_check = imageCheck(valE)
    if image_check == '':
        currentA = "A" + str(counter)
        valA = ws[currentA].value
        image_hash = getHash(valA)
        the_image = '<img src="https://www.thenewamerican.com/media/k2/items/src/' + image_hash + '.jpg">'
        append_portion = valE.split('<p')[0]
        valE = valE.replace( append_portion, (the_image), 1 )
        ws[currentE].value = str(valE)

    counter = counter + 1


#Concat post_id to post_name
counter = 2 #start on row after header
while ( counter < maxRow ):   
    currentA = "A" + str(counter)
    valA = ws[currentA].value
    currentK = "K" + str(counter)
    valK = ws[currentK].value

    concatVal = str(valA) + "-" + str(valK)
    ws[currentK].value = concatVal

    counter = counter + 1

#Format metadescription
#If K is blank, paste L
counter = 2 #start on row after header
while ( counter < maxRow ):
    #seo_metadesc   
    currentN = "N" + str(counter)
    valN = ws[currentN].value

    if valN == None:
        #seo_title
        currentP = "P" + str(counter)
        valP = ws[currentP].value 
        ws[currentN].value = valP

    counter = counter + 1


#remove all but first keyword from metakey
counter = 2 #start on row after header
while ( counter < maxRow ):   
    currentO = "O" + str(counter)
    valO = ws[currentO].value

    # if valO != None:
    #     firstkey = valO.split(',')[0]
    #     ws[currentO].value = firstkey
    # else:
    #     ws[currentO].value = sys.argv[2]
        #ws[currentN].value = "Asia"
    
    if valO == None:
        ws[currentO].value = sys.argv[2]

    counter = counter + 1

wb.save(save_path)
