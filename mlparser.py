# -*- coding: utf-8 -*-

import lxml.html as html
from lxml import etree
import json

print("Starting the Million Live Audio Room Parser...")
print("Run this program in directory with html file of")
print("URL:http://imas.gree-apps.net/app/index.php/audio_room")
print("renamed to ml.html. ")

# Parsing the HTML
parser = etree.HTMLParser(encoding='utf-8')
tree = etree.parse('ml.html', parser)

# The parser will return a list of nodes with 25 elements
# The album info is stored in the last script block (24)

scriptNodes = tree.xpath('//script')
scriptNode = scriptNodes[len(scriptNodes)-1]

# In order to let json read this string
# need to cut out "var albums = " and ";" at the two ends

scriptText = scriptNode.text

startIdentifier = 'var albums = '
startIndex = scriptText.find(startIdentifier) + len(startIdentifier)
endIndex = (len(scriptText) - scriptText.find(";")) * -1

scriptText = scriptText[startIndex:endIndex]




def byteify(input):
    if isinstance(input, dict):
        return {byteify(key): byteify(value)
                for key, value in input.iteritems()}
    elif isinstance(input, list):
        return [byteify(element) for element in input]
    elif isinstance(input, unicode):
        return input.encode('utf-8')
    else:
        return input

parsedJson = json.loads(scriptText)
#parsedJsonUnicoded = byteify(parsedJson)
#parsedJson = json.loads(scriptTextUnicoded, strict=False)

with open('test.txt', encoding='utf-8', mode='w+') as f:
    #f.write(scriptTextUnicoded[:500])
#with open('test.txt', encoding='utf-8', mode='w+') as f:
    f.write(json.dumps(parsedJson, indent = 4, sort_keys = True))
