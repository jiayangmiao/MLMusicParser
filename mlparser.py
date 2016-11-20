# -*- coding: utf-8 -*-

import json
import os
import sys
from sys import exit
from time import strftime

import xlsxwriter
from lxml import etree

reload(sys)
sys.setdefaultencoding('utf8')

print("Starting the Million Live Audio Room Parser...")

fileName = ('ml.html')

# Todo: make pyinstaller work with etree
# determine if application is a script file or frozen exe
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

rootPath = os.path.abspath(application_path)
#filePath = os.path.join(application_path, fileName)
filePath = rootPath + "/" + fileName
print(filePath)

# Check for input file
if os.path.exists(filePath) != True:
    print("Error: Cannot find ml.html.")
    print("To get started properly, save the webpage of")
    print("URL:http://imas.gree-apps.net/app/index.php/audio_room")
    print("as ml.html and put it in same directory as the script.")
    print("Exiting the program ...")
    exit()

# Parse the HTML
# The album info is stored in the last script block (24)

parser = etree.HTMLParser(encoding='utf-8')
print("Reading the file ...")
tree = etree.parse('ml.html', parser)
print("Passed HTML parsing!")
scriptNodes = tree.xpath('//script')
scriptNode = scriptNodes[len(scriptNodes) - 1]

# Cut out "var albums = " and ";" at the two ends to get json string

scriptText = scriptNode.text
startIdentifier = 'var albums = '
startIndex = scriptText.find(startIdentifier) + len(startIdentifier)
endIndex = (len(scriptText) - scriptText.find(";")) * -1
scriptText = scriptText[startIndex:endIndex]

parsedJson = json.loads(scriptText)

# Create output file and basic formatting
dateString = strftime("%Y-%m-%d%H%M")
workbook = xlsxwriter.Workbook('MLMusicInfo{}.xlsx'.format(dateString))
worksheet = workbook.add_worksheet()

worksheet.set_column(0, 1, 30)
worksheet.set_column(2, 2, 40)
worksheet.set_column(3, 3, 12)
worksheet.set_column(4, 4, 13)
worksheet.set_column(5, 5, 40)
worksheet.set_column(6, 7, 45)
worksheet.set_column(8, 8, 10)
worksheet.set_column(9, 9, 30)

headerFormat = workbook.add_format()
headerFormat.set_align('center')
headerFormat.set_align('vcenter')
headerFormat.set_border(1)
headerFormat.set_bold()

defaultFormat = workbook.add_format()
defaultFormat.set_align('center')
defaultFormat.set_align('vcenter')
defaultFormat.set_text_wrap()
defaultFormat.set_border(1)

urlFormat = workbook.add_format()
urlFormat.set_align('vcenter')
urlFormat.set_border(1)

importantFormat = workbook.add_format()
importantFormat.set_align('center')
importantFormat.set_align('vcenter')
importantFormat.set_text_wrap()
importantFormat.set_border(1)
importantFormat.set_bold()

worksheet.merge_range(0, 0, 0, 9, 'Million Live! in game Audio Room info sheet', headerFormat)
worksheet.write(1, 0, 'Album Composer', headerFormat)
worksheet.write(1, 1, 'Album Lyricist', headerFormat)
worksheet.write(1, 2, 'Album Artist', headerFormat)
worksheet.write(1, 3, 'Album Cover', headerFormat)
worksheet.write(1, 4, 'Album Release', headerFormat)
worksheet.write(1, 5, 'Album Title', headerFormat)
worksheet.write(1, 6, 'Song Title', headerFormat)
worksheet.write(1, 7, 'Song Artist', headerFormat)
worksheet.write(1, 8, 'Song URL', headerFormat)
worksheet.write(1, 9, 'Song Credit', headerFormat)

# Counter recording which row to write to
count = 2

for thisItem in parsedJson:
    # Check if is live setlist (we are only parsing CDs)
    if thisItem['is_live'] == 1:
        continue

    else: # Actual album!
        endCount = count + len(thisItem['records'])-1

        # Writing the album info
        if len(thisItem['records']) == 1:
            worksheet.write(count, 0, thisItem['composer'], defaultFormat)
            worksheet.write(count, 1, thisItem['lyricist'], defaultFormat)
            worksheet.write(count, 2, thisItem['album_artist_name'].replace('<br />','、').replace('<br>','、'), importantFormat)
            worksheet.write(count, 3, thisItem['jacket'], urlFormat)
            worksheet.write(count, 4, thisItem['release_date'], defaultFormat)
            worksheet.write(count, 5, thisItem['album_title'].replace('<br/>','\n').replace('<br>','\n'), importantFormat)
        else:
            worksheet.merge_range(count, 0, endCount, 0, thisItem['composer'], defaultFormat)
            worksheet.merge_range(count, 1, endCount, 1, thisItem['lyricist'], defaultFormat)
            worksheet.merge_range(count, 2, endCount, 2, thisItem['album_artist_name'].replace('<br />','、').replace('<br>','、'), importantFormat)
            worksheet.merge_range(count, 3, endCount, 3, thisItem['jacket'], urlFormat)
            worksheet.merge_range(count, 4, endCount, 4, thisItem['release_date'], defaultFormat)
            worksheet.merge_range(count, 5, endCount, 5, thisItem['album_title'].replace('<br/>','\n').replace('<br>','\n'), importantFormat)

        # Writing the track info
        trackCount = count
        #for thisTrack in sorted(thisItem['records']):
        for thisTrack in thisItem['records']:
            thisTrackInfo = thisItem['records']['{}'.format(thisTrack)]
            worksheet.write(trackCount, 6, thisTrackInfo['music_title'].replace('<small>',' ').replace('</small>',' ').replace('<br>',' '), importantFormat)
            worksheet.write(trackCount, 7, thisTrackInfo['music_artist_name'], importantFormat)
            worksheet.write(trackCount, 8, thisTrackInfo['music_src'], urlFormat)
            worksheet.write(trackCount, 9, thisTrackInfo['credit'], defaultFormat)

            trackCount += 1

        count = endCount + 1

workbook.close()

