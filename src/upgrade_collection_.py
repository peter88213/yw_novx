"""novelyst collection upgrader

- Convert a novelyst .nvcx collection file to .nvcx format.
- Convert the collection's .yw7 project files to .novx format. 

Copyright (c) 2024 Peter Triesberger
For further information see https://github.com/peter88213/
License: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.en.html)
"""
import sys
import os
import xml.etree.ElementTree as ET

from novxlib.xml.xml_indent import indent
from yw_novx_ import yw_novx

XML_HEADER = '''<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE COLLECTION SYSTEM "nvcx_1_0.dtd">
<?xml-stylesheet href="collection.css" type="text/css"?>
'''


def postprocess_xml_file(filePath):
    """Postprocess an xml file created by ElementTree.
    
    Positional argument:
        filePath -- str: path to xml file.
    
    Read the xml file, put a header on top. Overwrite the .nvcx xml file.
    """
    with open(filePath, 'r', encoding='utf-8') as f:
        text = f.read()
    with open(filePath, 'w', encoding='utf-8') as f:
        f.write(f'{XML_HEADER}{text}')


def convert(sourcePath):

    def set_element(xmlElement, targetElement, prefix):
        elemId = xmlElement.attrib[xmlMap['id']]
        targetElement.set('id', f"{prefix}{elemId}")
        xmlTitle = xmlElement.find(xmlMap['title'])
        if xmlTitle is not None:
            title = xmlTitle.text
            if title:
                ET.SubElement(targetElement, 'Title').text = title
        xmlDesc = xmlElement.find(xmlMap['desc'])
        if xmlDesc is not None:
            desc = xmlDesc.text
            if desc:
                targetDesc = ET.SubElement(targetElement, 'Desc')
                for paragraph in desc.split('\n'):
                    ET.SubElement(targetDesc, 'p').text = paragraph.strip()
        xmlPath = xmlElement.find(xmlMap['path'])
        if xmlPath is not None:

            # xmlElement is BOOK.
            yw7Path = xmlPath.text
            if yw7Path and os.path.isfile(yw7Path):
                bookPath, bookExt = os.path.splitext(yw7Path)
                if bookExt == '.yw7':
                    novxPath = f'{bookPath}.novx'
                    if not os.path.isfile(novxPath):

                        # Convert book to .novx.
                        yw_novx(yw7Path)
                    ET.SubElement(targetElement, 'Path').text = novxPath

    pathRoot , extension = os.path.splitext(sourcePath)
    if extension != '.pwc':
        raise ValueError(f'File must be .pwc type, but is "{extension}".')
    targetPath = f'{pathRoot}.nvcx'

    # There are different collection file versions.
    v1Map = dict(
            collection='collection',
            series='series',
            book='book',
            id='id',
            path='path',
            title='title',
            desc='desc',
            )
    oldMap = dict(
            collection='COLLECTION',
            series='SERIES',
            book='BOOK',
            id='ID',
            path='Path',
            title='Title',
            desc='Desc',
            )
    xmlSourceTree = ET.parse(sourcePath)
    xmlRoot = xmlSourceTree.getroot()
    if xmlRoot.tag == v1Map['collection']:
        xmlMap = v1Map
    elif xmlRoot.tag == oldMap['collection']:
        xmlMap = oldMap
    else:
        raise Exception(f'No collection found in file: "{os.path.normpath(sourcePath)}".')

    try:
        majorVersionStr, minorVersionStr = xmlRoot.attrib['version'].split('.')
        majorVersion = int(majorVersionStr)
    except:
        raise Exception(f'No valid version found in file: "{os.path.normpath(sourcePath)}".')

    if majorVersion > 1:
        raise Exception('The collection was created with a newer plugin version.')

    targetRoot = ET.Element('COLLECTION')
    targetRoot.set('version', '1.0')
    for xmlElement in xmlRoot:
        if xmlElement.tag == xmlMap['book']:
            targetBook = ET.SubElement(xmlRoot, 'BOOK')
            set_element(xmlElement, targetBook, 'bk')
        elif xmlElement.tag == xmlMap['series']:
            targetSeries = ET.SubElement(targetRoot, 'SERIES')
            set_element(xmlElement, targetSeries, 'sr')
            for xmlBook in xmlElement.iter(xmlMap['book']):
                targetBook = ET.SubElement(targetSeries, 'BOOK')
                set_element(xmlBook, targetBook, 'bk')

    indent(targetRoot)
    xmlTree = ET.ElementTree(targetRoot)
    backedUp = False
    if os.path.isfile(targetPath):
        os.replace(targetPath, f'{targetPath}.bak')
        backedUp = True
    try:
        xmlTree.write(targetPath, encoding='utf-8')
    except:
        if backedUp:
            os.replace(f'{targetPath}.bak', targetPath)
        raise Exception(f'{_("Cannot write file")}: "{os.path.normpath(targetPath)}".')

    postprocess_xml_file(targetPath)


if __name__ == '__main__':
    convert(sys.argv[1])
    print('Done')

