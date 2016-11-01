# -*- coding: utf-8 -*-

import os.path
from sys import exit

import lxml.html as html
from lxml import etree
import json

import xlsxwriter

print("Starting the Million Live Audio Room Parser...")

# Check for input file

if os.path.exists("ml.html") != True:
    print("Error: Cannot find ml.html.")
    print("To get started properly, save the webpage of")
    print("URL:http://imas.gree-apps.net/app/index.php/audio_room")
    print("as ml.html and put it in same directory as the script.")
    print("Exiting the program ...")
    exit()

# Parse the HTML

parser = etree.HTMLParser(encoding='utf-8')
print("Reading the file ...")
tree = etree.parse('ml.html', parser)

# The parser will return a list of nodes with 25 elements
# The album info is stored in the last script block (24)

scriptNodes = tree.xpath('//script')
scriptNode = scriptNodes[len(scriptNodes)-1]

# Cut out "var albums = " and ";" at the two ends to get json string

scriptText = scriptNode.text
startIdentifier = 'var albums = '
startIndex = scriptText.find(startIdentifier) + len(startIdentifier)
endIndex = (len(scriptText) - scriptText.find(";")) * -1
scriptText = scriptText[startIndex:endIndex]

# Parse the json string

parsedJson = json.loads(scriptText)

for thisItem in parsedJson:
    # Check if is live setlist (which we are not parsing)
    if thisItem['is_live'] != 0:
        continue

    else: # Actual album!

        print("Found %s" %thisItem['album_title'])

# Create output file and basic formatting

workbook = xlsxwriter.Workbook('MLMusicInfo.xlsx')
worksheet = workbook.add_worksheet()
defaultformat = workbook.add_format()
defaultformat.set_border(1)
worksheet.set_column(0, 1, 30)
worksheet.set_column(2, 2, 40)
worksheet.set_column(3, 3, 10)
worksheet.set_column(4, 4, 13)
worksheet.set_column(5, 5, 40)
worksheet.set_column(6, 7, 45)
worksheet.set_column(8, 8, 10)
worksheet.set_column(9, 9, 30)

worksheet.write(0, 0, 'Million Live! In game Audio Room info sheet')

headerformat = workbook.add_format()
headerformat.set_align('center')
headerformat.set_bold
worksheet.write(1, 0, 'Album Composer', headerformat)
worksheet.write(1, 1, 'Album Lyricist', headerformat)
worksheet.write(1, 2, 'Album Artist', headerformat)
worksheet.write(1, 3, 'Album Cover', headerformat)
worksheet.write(1, 4, 'Album Release', headerformat)
worksheet.write(1, 5, 'Album Title', headerformat)
worksheet.write(1, 6, 'Song Title', headerformat)
worksheet.write(1, 7, 'Song Artist', headerformat)
worksheet.write(1, 8, 'Song URL', headerformat)
worksheet.write(1, 9, 'Song Credit', headerformat)


workbook.close()

