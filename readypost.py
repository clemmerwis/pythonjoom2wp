#! python3
import time
import os
import sys
import openpyxl
import re
import hashlib
import json
from pathlib import Path

# Script Process:
# Open {catnum}Content.xlsx
#   * If "Video" found in excerpt, write post_id to Dic - write to json file
#   * If find <!--IMAGE.*--> in excerpt, remove entirely (rotator images that don't exist)
#   * Regex imageCheck in excerpt, if found - format existing image_url || If no image found, get hash from post_id, create image_url 
#   * Capture post_id and image_url in a dictionary - write to json file
#   * Remove image_url from post_excerpt
#   * Concat post_excerpt onto post_content
#   * Remove '<p>{modulepos inner_text_ad}</p>'
#   * Concat post_id to post_name (post-slug)
#   * Remove "Video -" records from spreadsheet
#   * Save content




# Check for 1 argument
# if sys.argv[1] == None:
#     print('"Script requires argument: "categoryNumber"')
#     sys.exit()


def imgStart(val):
    img_start_regex = re.compile(r'<img.*src=\"images')
    img_start_mo = img_start_regex.search(val)
    if img_start_mo != None:
        img_start_mo = str(img_start_mo.group())
        val = val.replace(img_start_mo, '<img src="http://www.thenewamerican.com/images')
    return val


def imgEnd(val):
    img_end_regex = re.compile(r'((.jpg|png)\" .*)/>')
    img_end_mo = img_end_regex.search(val)
    if img_end_mo != None:
        img_end_mo = str(img_end_mo.group())
        if img_end_mo.find("jpg") != -1:
            val = val.replace(img_end_mo, '.jpg"/>')
        elif img_end_mo.find("png") != -1:
            val = val.replace(img_end_mo, '.png"/>')
    return val


def format_image(val):
    val = imgStart(val)
    val = imgEnd(val)
    return val


def format_modulepos(val):
    mod_check_regex = re.compile(r'<p.*?{modulepos inner_text_ad}.*?p>')
    mod_tuples = mod_check_regex.findall(val)
    if mod_tuples != None:
        for tup in mod_tuples:
            val = val.replace(tup, "")
    return val


def get_youtube_embeds(val):
    embeds_regex = re.compile(r'(<p>.*?){youtube}(.*?){/youtube}(.*?</p>)')
    embed_check = embeds_regex.findall(val)
    embeds = list()
    for tup in embed_check:
        embeds.append(tup[1])
    return embeds

def replace_embeds(val, embeds):
    for embed in embeds:
        yt_iframe = create_video_iframe(embed)
        embed_regex_tags = re.compile(r'<p>.*?{youtube}.*{/youtube}.*?</p>')
        embed_tuples = embed_regex_tags.findall(val)
        for tup in embed_tuples:
            if embed in tup:
                val = val.replace(tup, yt_iframe)
    return val

def getHash(postid):
    the_hash = hashlib.md5(("Image" + str(postid)).encode('utf-8')).hexdigest()
    return the_hash


def imageCheck(val):
    check = ''
    img_check_regex = re.compile(r'<img.*(jpg|png)".*/>')
    img_check_mo = img_check_regex.search(val)
    if img_check_mo != None:
        check = img_check_mo.group()
    return check


def get_image_url(val):
    image_tag = imageCheck(val)
    img_url_regex = re.compile(r'src="(.*(jpg|png))')
    img_url_mo = img_url_regex.search(image_tag)
    img_url = img_url_mo.group(1)
    return img_url


def set_current_cell(letter, number):
    currentCell = letter + str(counter)
    valCell = ws[currentCell].value
    valCell = str(valCell)
    return valCell


def insert_before_first_p_tag(val, to_insert):
    append_portion = val.split('<p')[0]
    val = val.replace( append_portion, to_insert, 1 )
    return val


def create_video_iframe(videoID):
    iframe = f'<iframe frameborder="0" allowfullscreen="allowfullscreen" src="https://www.youtube.com/embed/{videoID}?rel=0&amp;showinfo=0&amp;autoplay=0"></iframe>'
    return iframe

def search_colA(searching_for):
    counter = 2  # start on row after header
    while (counter < (ws.max_row + 1)):
        currentA = "A" + str(counter)
        valA = ws[currentA].value
        if valA == searching_for:
            found_on_counter = counter
            break
        counter = counter + 1
    return found_on_counter

#Vars
# change me
cat = sys.argv[1]

# files and folders
jsondata_folder = Path("C:/Users/chris/Desktop/migAssets/json")
data_folder = Path("C:/Users/chris/Desktop/migAssets")
file_name = cat + "Content.xlsx"
file_to_open = data_folder / file_name
save_as = cat + "Content-formatted.xlsx"
save_path = data_folder / save_as


#Start
#---------
wb = openpyxl.load_workbook(file_to_open)
ws = wb.active


#get max row of excel sheet that contains data
maxRow = None
counter = 2  # start on row after header
while (counter < (ws.max_row + 2)):
    currentA = "A" + str(counter)
    valA = ws[currentA].value
    if valA == None:
        maxRow = counter
        break
    counter = counter + 1

if (maxRow == None):
    maxRow = counter - 1
    
print(maxRow)

idsUrls_Dic = {}
idsSkipped_Dic = {}
# Format spreadsheet
counter = 2 #start on row after header
while ( counter < maxRow ):  
    # Get this row's values for: post_id, post_excerpt, and post_content, post_name
    valA = set_current_cell("A", counter)
    if valA == None:
        print("empty: " + str(counter))
        continue
    valG = set_current_cell("G", counter)
    valE = set_current_cell("E", counter)
    valK = set_current_cell("K", counter)

    # if is a Video post
    if (valG.find("VIDEO -") ) != -1:
        # for now just keep track of video posts
        print(valA)
        idsSkipped_Dic.update( {"row" + str(counter): str(valA)} )
        counter += 1
        continue

        # Get video ID 
        videoID_regex = re.search('{youtube}(.*){/youtube}', valE)
        videoID = videoID_regex.group(1)
        
        # remove anything before first p tag on post_content
        valE = insert_before_first_p_tag("valE", "")

        # create and insert iframe into post_excerpt
        iframe = create_video_iframe(videoID)
        valG = insert_before_first_p_tag(valG, iframe)
    # else is a normal post
    else:
        # remove anything before first p tag
        valG = insert_before_first_p_tag(valG, "")

        # If no image, get image from ID and hash
        if imageCheck(valG) == '':
            image_hash = getHash(valA)
            # the_image_url = '<img src="https://www.thenewamerican.com/media/k2/items/src/' + image_hash + '.jpg">'
            the_image_url = 'https://www.thenewamerican.com/media/k2/items/src/' + image_hash + '.jpg'

            # add to Dictionary of ids and urls
            idsUrls_Dic.update( {valA: the_image_url} )

            # add image url to column R
            ws["Q"+str(counter)].value = idsUrls_Dic[valA]
        else:
            # format image in column G 
            ws["G"+str(counter)].value = format_image(valG)
            # Reattain g cell
            valG = set_current_cell("G", counter)
            the_image_url = get_image_url(valG)

            # add to Dictionary of ids and urls
            idsUrls_Dic.update( {valA: the_image_url} )

            # add image url to column R
            ws["Q"+str(counter)].value = idsUrls_Dic[valA]

            # remove image from excerpt
            valG = valG.replace( imageCheck(valG), "", 1 )
            
    #update post_content to excerpt + content
    ws["E"+str(counter)].value = valG + valE
    valE = set_current_cell("E", counter)

    # remove revive ads
    valE = format_modulepos(valE)
    # update post_content value 
    ws["E"+str(counter)].value = valE

    # replace embed codes with youtube iframes
    embeds = get_youtube_embeds(valE)
    valE = replace_embeds(valE, embeds)
    # update post_content value 
    ws["E"+str(counter)].value = valE
        
    #concat id to post_name
    ws["K"+str(counter)].value = valA + '-' + valK
    counter += 1


for key, value in idsSkipped_Dic.items():
    row_to_delete = search_colA(value)
    ws.delete_rows(row_to_delete)

# Writing to 252Skipped.json 
jsonified = json.dumps(idsSkipped_Dic, indent = 4)
fname = cat + "skipped.json"
json_file = jsondata_folder / fname
with open(json_file, "w+") as outfile: 
    outfile.write(jsonified)
    outfile.close() 
print('created: ' + str(json_file))

# Writing to 252Feats.json 
jsonified = json.dumps(idsUrls_Dic, indent = 4)
fname = cat + "feats.json"
json_file = jsondata_folder / fname
with open(json_file, "w+") as outfile: 
    outfile.write(jsonified) 
    outfile.close()
print('created: ' + str(json_file))


wb.save(save_path)
print('created: ' + str(save_path))

print("Finished")
sys.exit()
    

    



