#! python3
from lxml import html, etree
from pathlib import Path
import time
import os
import sys
import openpyxl
import re
import hashlib
import json
import requests
import webbrowser

def filter_linksByUS(linklist):
    toRemove_list = []
    removeAfter_index = linklist.index("https://virgin.craigslist.org/")
    for link in linklist:
        if linklist.index(link) > removeAfter_index or "#" in link:
            toRemove_list.append(link)
    for rlink in toRemove_list:
        linklist.remove(rlink)
    return linklist

class Page:
    def __init__(self, request):
        self.request = request

    def get_request(self):
        return self.request

    def get_html_fromRequest(self):
        # convert the data received into searchable HTML
        extractedHtml = html.fromstring(self.request.content)
        return extractedHtml

# class HomePage(Page):
    def get_links(self):
        html = self.get_html_fromRequest()
        # use an XPath query to find the image link (the 'src' attribute of the 'img' tag).
        links = html.xpath("//a/@href")
        return links

# startnum = int(sys.argv[1])
startnum = 0
endnum = startnum + 10
webpageLink = "https://www.craigslist.org/about/sites"

# create page object
r = requests.get(webpageLink)
craigslist_home = Page(r)

# Loop through US links
links = craigslist_home.get_links()
filtered_links = filter_linksByUS(links)
for link in filtered_links:
    if filtered_links.index(link) < startnum:
        continue
    webbrowser.open(link + "d/web-html-info-design/search/web")
    print("opened: " + link)
    print("")
    if filtered_links.index(link) >= endnum:
        print("finished")
        sys.exit()


