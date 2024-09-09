#!/usr/bin/python3
"""Convert yw7 to novx.

Copyright (c) 2024 Peter Triesberger
For further information see https://github.com/peter88213/yw_novx
License: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.en.html)
"""
SUFFIX = ''

import sys
import os

from datetime import datetime
from html import unescape
import re

from abc import ABC
from urllib.parse import quote

from calendar import day_name
from calendar import month_name
from datetime import date
from datetime import time
import gettext
import locale

ROOT_PREFIX = 'rt'
CHAPTER_PREFIX = 'ch'
PLOT_LINE_PREFIX = 'ac'
SECTION_PREFIX = 'sc'
PLOT_POINT_PREFIX = 'ap'
CHARACTER_PREFIX = 'cr'
LOCATION_PREFIX = 'lc'
ITEM_PREFIX = 'it'
PRJ_NOTE_PREFIX = 'pn'
CH_ROOT = f'{ROOT_PREFIX}{CHAPTER_PREFIX}'
PL_ROOT = f'{ROOT_PREFIX}{PLOT_LINE_PREFIX}'
CR_ROOT = f'{ROOT_PREFIX}{CHARACTER_PREFIX}'
LC_ROOT = f'{ROOT_PREFIX}{LOCATION_PREFIX}'
IT_ROOT = f'{ROOT_PREFIX}{ITEM_PREFIX}'
PN_ROOT = f'{ROOT_PREFIX}{PRJ_NOTE_PREFIX}'

BRF_SYNOPSIS_SUFFIX = '_brf_synopsis'
CHAPTERS_SUFFIX = '_chapters_tmp'
CHARACTER_REPORT_SUFFIX = '_character_report'
CHARACTERS_SUFFIX = '_characters_tmp'
CHARLIST_SUFFIX = '_charlist_tmp'
DATA_SUFFIX = '_data'
GRID_SUFFIX = '_grid_tmp'
ITEM_REPORT_SUFFIX = '_item_report'
ITEMLIST_SUFFIX = '_itemlist_tmp'
ITEMS_SUFFIX = '_items_tmp'
LOCATION_REPORT_SUFFIX = '_location_report'
LOCATIONS_SUFFIX = '_locations_tmp'
LOCLIST_SUFFIX = '_loclist_tmp'
MANUSCRIPT_SUFFIX = '_manuscript_tmp'
PARTS_SUFFIX = '_parts_tmp'
PLOTLIST_SUFFIX = '_plotlist'
PLOTLINES_SUFFIX = '_plotlines_tmp'
PROJECTNOTES_SUFFIX = '_projectnote_report'
PROOF_SUFFIX = '_proof_tmp'
SECTIONLIST_SUFFIX = '_sectionlist'
SECTIONS_SUFFIX = '_sections_tmp'
STAGES_SUFFIX = '_structure_tmp'
XREF_SUFFIX = '_xref'


class Error(Exception):
    pass


try:
    LOCALE_PATH
except NameError:
    locale.setlocale(locale.LC_TIME, "")
    LOCALE_PATH = f'{os.path.dirname(sys.argv[0])}/locale/'
    try:
        CURRENT_LANGUAGE = locale.getlocale()[0][:2]
    except:
        CURRENT_LANGUAGE = locale.getdefaultlocale()[0][:2]
    try:
        t = gettext.translation('novelibre', LOCALE_PATH, languages=[CURRENT_LANGUAGE])
        _ = t.gettext
    except:

        def _(message):
            return message

WEEKDAYS = day_name
MONTHS = month_name


def norm_path(path):
    if path is None:
        path = ''
    return os.path.normpath(path)


def string_to_list(text, divider=';'):
    elements = []
    try:
        tempList = text.split(divider)
        for element in tempList:
            element = element.strip()
            if element and not element in elements:
                elements.append(element)
        return elements

    except:
        return []


def list_to_string(elements, divider=';'):
    try:
        text = divider.join(elements)
        return text

    except:
        return ''


def intersection(elemList, refList):
    return [elem for elem in elemList if elem in refList]


def verified_date(dateStr):
    if dateStr is not None:
        date.fromisoformat(dateStr)
    return dateStr


def verified_int_string(intStr):
    if intStr is not None:
        int(intStr)
    return intStr


def verified_time(timeStr):
    if  timeStr is not None:
        time.fromisoformat(timeStr)
        while timeStr.count(':') < 2:
            timeStr = f'{timeStr}:00'
    return timeStr



class File(ABC):
    DESCRIPTION = _('File')
    EXTENSION = None
    SUFFIX = None

    def __init__(self, filePath, **kwargs):
        self.novel = None
        self._filePath = None
        self.projectName = None
        self.projectPath = None
        self.sectionsSplit = False
        self.filePath = filePath

    @property
    def filePath(self):
        return self._filePath

    @filePath.setter
    def filePath(self, filePath: str):
        filePath = filePath.replace('\\', '/')
        if self.SUFFIX is not None:
            suffix = self.SUFFIX
        else:
            suffix = ''
        if filePath.lower().endswith(f'{suffix}{self.EXTENSION}'.lower()):
            self._filePath = filePath
            try:
                head, tail = os.path.split(os.path.realpath(filePath))
            except:
                head, tail = os.path.split(filePath)
            self.projectPath = quote(head.replace('\\', '/'), '/:')
            self.projectName = quote(tail.replace(f'{suffix}{self.EXTENSION}', ''))

    def is_locked(self):
        return False

    def read(self):
        raise NotImplementedError

    def write(self):
        raise NotImplementedError



def create_id(elements, prefix=''):
    i = 1
    while f'{prefix}{i}' in elements:
        i += 1
    return f'{prefix}{i}'

from xml import sax


class NovxToShortcode(sax.ContentHandler):
    NOTE_TYPES = {
        'footnote':'@fn',
        'endnote':'@en',
    }

    def __init__(self):
        super().__init__()
        self.textList = None
        self._paragraph = None
        self._span = None
        self._comment = None

    def feed(self, xmlString):
        self.textList = []
        if xmlString:
            self._comment = False
            self._paragraph = False
            self._span = []
            sax.parseString(xmlString, self)

    def characters(self, content):
        if self._paragraph:
            self.textList.append(content)

    def endElement(self, name):
        if name == 'p':
            while self._span:
                self.textList.append(self._span.pop())
            if self._comment:
                self.textList.append(' ')
            else:
                self.textList.append('\n')
            self._paragraph = False
            return

        if name == 'em':
            self.textList.append('[/i]')
            return

        if name == 'strong':
            self.textList.append('[/b]')
            return

        if name == 'span':
            if self._span:
                self.textList.append(self._span.pop())
            return

        if name in ('comment', 'note'):
            self._comment = False
            self.textList.append('*/')
            return

        if name in ('creator', 'date', 'note-citation'):
            self._paragraph = True

    def startElement(self, name, attrs):
        xmlAttributes = {}
        for attribute in attrs.items():
            attrKey, attrValue = attribute
            xmlAttributes[attrKey] = attrValue
        locale = xmlAttributes.get('xml:lang', None)

        if name == 'p':
            self._paragraph = True
            if xmlAttributes.get('style', None) == 'quotations':
                self.textList.append('> ')
            return

        if name == 'em':
            self.textList.append('[i]')
            return

        if name == 'strong':
            self.textList.append('[b]')
            return

        if name == 'span':
            if locale is None:
                return

            self._span.append(f'[/lang={locale}]')
            self.textList.append(f'[lang={locale}]')
            return

        if name in ('comment', 'note'):
            self._comment = True
            self.textList.append('/*')

            if name == 'note':
                noteClass = xmlAttributes.get('class', 'footnote')
                self.textList.append(f"{self.NOTE_TYPES.get(noteClass, '@fn')} ")
            return

        if name in ('creator', 'date', 'note-citation'):
            self._paragraph = False



def indent(elem, level=0):
    PARAGRAPH_LEVEL = 5

    i = f'\n{level * "  "}'
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = f'{i}  '
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        if level < PARAGRAPH_LEVEL:
            for elem in elem:
                indent(elem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


def strip_illegal_characters(text):
    return re.sub('[\x00-\x08|\x0b-\x0c|\x0e-\x1f]', '', text)


LOCALE_PATH = f'{os.path.dirname(sys.argv[0])}/locale/'
try:
    CURRENT_LANGUAGE = locale.getlocale()[0][:2]
except:
    CURRENT_LANGUAGE = locale.getdefaultlocale()[0][:2]
try:
    t = gettext.translation('nv_yw7', LOCALE_PATH, languages=[CURRENT_LANGUAGE])
    _ = t.gettext
except:

    def _(message):
        return message

import xml.etree.ElementTree as ET


class Yw7File(File):
    DESCRIPTION = _('yWriter 7 project')
    EXTENSION = '.yw7'

    PRJ_KWVAR_YW7 = [
        'Field_WorkPhase',
        'Field_RenumberChapters',
        'Field_RenumberParts',
        'Field_RenumberWithinParts',
        'Field_RomanChapterNumbers',
        'Field_RomanPartNumbers',
        'Field_ChapterHeadingPrefix',
        'Field_ChapterHeadingSuffix',
        'Field_PartHeadingPrefix',
        'Field_PartHeadingSuffix',
        'Field_CustomGoal',
        'Field_CustomConflict',
        'Field_CustomOutcome',
        'Field_CustomChrBio',
        'Field_CustomChrGoals',
        'Field_SaveWordCount',
        'Field_LanguageCode',
        'Field_CountryCode',
        ]

    CHP_KWVAR_YW7 = [
        'Field_NoNumber',
        'Field_ArcDefinition',
        'Field_Arc_Definition',
        ]

    SCN_KWVAR_YW7 = [
        'Field_SceneArcs',
        'Field_SceneAssoc',
        'Field_CustomAR',
        'Field_SceneMode',
        ]

    CRT_KWVAR_YW7 = [
        'Field_Link',
        'Field_BirthDate',
        'Field_DeathDate',
        ]

    LOC_KWVAR_YW7 = [
        'Field_Link',
        ]

    ITM_KWVAR_YW7 = [
        'Field_Link',
        ]

    _CDATA_TAGS = [
        'Title',
        'AuthorName',
        'Bio',
        'Desc',
        'FieldTitle1',
        'FieldTitle2',
        'FieldTitle3',
        'FieldTitle4',
        'LaTeXHeaderFile',
        'Tags',
        'AKA',
        'ImageFile',
        'FullName',
        'Goals',
        'Notes',
        'RTFFile',
        'SceneContent',
        'Outcome',
        'Goal',
        'Conflict'
        'Field_ChapterHeadingPrefix',
        'Field_ChapterHeadingSuffix',
        'Field_PartHeadingPrefix',
        'Field_PartHeadingSuffix',
        'Field_CustomGoal',
        'Field_CustomConflict',
        'Field_CustomOutcome',
        'Field_CustomChrBio',
        'Field_CustomChrGoals',
        'Field_ArcDefinition',
        'Field_SceneArcs',
        'Field_CustomAR',
        ]

    STAGE_MARKER = 'stage'

    def __init__(self, filePath, **kwargs):
        super().__init__(filePath)
        self._nvSvc = kwargs['nv_service']
        self.tree = None
        self.wcLog = {}
        self._ywApIds = None

    def is_locked(self):
        return os.path.isfile(f'{self.filePath}.lock')

    def read(self):
        if self.is_locked():
            raise Error(f'{_("yWriter seems to be open. Please close first")}.')

        self._noteCounter = 0
        self._noteNumber = 0
        try:
            try:
                with open(self.filePath, 'r', encoding='utf-8') as f:
                    xmlText = f.read()
            except:
                with open(self.filePath, 'r', encoding='utf-16') as f:
                    xmlText = f.read()
        except:
            try:
                self.tree = ET.parse(self.filePath)
            except Exception as ex:
                raise Error(f'{_("Can not process file")} - {str(ex)}')

        xmlText = strip_illegal_characters(xmlText)
        root = ET.fromstring(xmlText)
        del xmlText

        self._ywApIds = []
        self.wcLog = {}
        self._read_project(root)
        self._read_locations(root)
        self._read_items(root)
        self._read_characters(root)
        self._read_projectvars(root)
        self._read_chapters(root)
        self._read_scenes(root)
        self._read_project_notes(root)

        xmlWclog = root.find('WCLog')
        if xmlWclog is not None:
            for xmlWc in xmlWclog.iterfind('WC'):
                wcDate = xmlWc.find('Date').text
                wcCount = xmlWc.find('Count').text
                wcTotalCount = xmlWc.find('TotalCount').text
                self.wcLog[wcDate] = [wcCount, wcTotalCount]

        for scId in self.novel.sections:
            if self.novel.sections[scId].characters is None:
                self.novel.sections[scId].characters = []
            if self.novel.sections[scId].locations is None:
                self.novel.sections[scId].locations = []
            if self.novel.sections[scId].items is None:
                self.novel.sections[scId].items = []
            if self.novel.sections[scId].status is None:
                self.novel.sections[scId].status = 1

        self.novel.check_locale()

    def write(self):
        if self.is_locked():
            raise Error(f'{_("yWriter seems to be open. Please close first")}.')

        self._novxParser = NovxToShortcode()
        self._noteCounter = 0
        self._noteNumber = 0
        if self.novel.languages is None:
            self.novel.get_languages()
        self._build_element_tree()
        self._write_element_tree(self)
        self._postprocess_xml_file(self.filePath)

    def _build_element_tree(self):

        def isTrue(value):
            if value:
                return '1'

        def set_element(parent, tag, text, index):
            if text is not None:
                subelement = ET.Element(tag)
                parent.insert(index, subelement)
                subelement.text = text
                index += 1
            return index

        def build_scene_subtree(xmlScene, prjScn, plotPoint=False):
            i = 1
            i = set_element(xmlScene, 'Title', prjScn.title, i)
            if prjScn.desc is not None:
                ET.SubElement(xmlScene, 'Desc').text = prjScn.desc

            if not plotPoint:
                scTags = prjScn.tags


            scTypeEncoding = (
                (False, None),
                (True, '1'),
                (True, '2'),
                (True, '0'),
                )
            if plotPoint:
                scType = 2
            elif prjScn.scType in (0, None):
                scType = 0
            elif prjScn.scType > 1:
                scType = 2
                if not scTags:
                    scTags = [self.STAGE_MARKER]
                elif not self.STAGE_MARKER in scTags:
                    scTags.append(self.STAGE_MARKER)
            else:
                scType = 3
            yUnused, ySceneType = scTypeEncoding[scType]
            if yUnused:
                ET.SubElement(xmlScene, 'Unused').text = '-1'
            if ySceneType is not None:
                ET.SubElement(xmlSceneFields[scId], 'Field_SceneType').text = ySceneType
            if plotPoint:
                ET.SubElement(xmlScene, 'Status').text = '1'
            elif prjScn.status is not None:
                ET.SubElement(xmlScene, 'Status').text = str(prjScn.status)

            if plotPoint:
                ET.SubElement(xmlScene, 'SceneContent')
                return

            self._novxParser.feed(f'<Content>{prjScn.sectionContent}</Content>')
            ET.SubElement(xmlScene, 'SceneContent').text = ''.join(self._novxParser.textList)
            if prjScn.notes:
                ET.SubElement(xmlScene, 'Notes').text = prjScn.notes
            if scTags:
                ET.SubElement(xmlScene, 'Tags').text = list_to_string(scTags)
            if prjScn.appendToPrev:
                ET.SubElement(xmlScene, 'AppendToPrev').text = '-1'

            if (prjScn.date) and (prjScn.time):
                separator = ' '
                dateTime = f'{prjScn.date}{separator}{prjScn.time}'
                ET.SubElement(xmlScene, 'SpecificDateTime').text = dateTime
                ET.SubElement(xmlScene, 'SpecificDateMode').text = '-1'
            elif (prjScn.day) or (prjScn.time):
                if prjScn.day:
                    ET.SubElement(xmlScene, 'Day').text = prjScn.day
                if prjScn.time:
                    hours, minutes, __ = prjScn.time.split(':')
                    ET.SubElement(xmlScene, 'Hour').text = hours
                    ET.SubElement(xmlScene, 'Minute').text = minutes

            if prjScn.lastsDays:
                ET.SubElement(xmlScene, 'LastsDays').text = prjScn.lastsDays
            if prjScn.lastsHours:
                ET.SubElement(xmlScene, 'LastsHours').text = prjScn.lastsHours
            if prjScn.lastsMinutes:
                ET.SubElement(xmlScene, 'LastsMinutes').text = prjScn.lastsMinutes

            if prjScn.scene == 2:
                ET.SubElement(xmlScene, 'ReactionScene').text = '-1'
            if prjScn.goal:
                ET.SubElement(xmlScene, 'Goal').text = prjScn.goal
            if prjScn.conflict:
                ET.SubElement(xmlScene, 'Conflict').text = prjScn.conflict
            if prjScn.outcome:
                ET.SubElement(xmlScene, 'Outcome').text = prjScn.outcome

            if prjScn.characters:
                xmlCharacters = ET.SubElement(xmlScene, 'Characters')
                for crId in prjScn.characters:
                    ET.SubElement(xmlCharacters, 'CharID').text = crId[2:]
            if prjScn.locations:
                xmlLocations = ET.SubElement(xmlScene, 'Locations')
                for lcId in prjScn.locations:
                    ET.SubElement(xmlLocations, 'LocID').text = lcId[2:]
            if prjScn.items:
                xmlItems = ET.SubElement(xmlScene, 'Items')
                for itId in prjScn.items:
                    ET.SubElement(xmlItems, 'ItemID').text = itId[2:]

        def build_chapter_subtree(xmlChapter, prjChp, plId=None, chType=None):

            chTypeEncoding = (
                (False, '0', '0'),
                (True, '1', '1'),
                (True, '1', '2'),
                (True, '1', '0'),
                )
            if chType is None:
                if plId is not None:
                    chType = 2
                elif prjChp.chType in (0, None):
                    chType = 0
                else:
                    chType = 3
            yUnused, yType, yChapterType = chTypeEncoding[chType]

            i = 1
            i = set_element(xmlChapter, 'Title', prjChp.title, i)
            i = set_element(xmlChapter, 'Desc', prjChp.desc, i)

            if yUnused:
                elem = ET.Element('Unused')
                elem.text = '-1'
                xmlChapter.insert(i, elem)
                i += 1

            xmlChapterFields = ET.SubElement(xmlChapter, 'Fields')
            i += 1
            if plId is None and prjChp.isTrash:
                ET.SubElement(xmlChapterFields, 'Field_IsTrash').text = '1'

            if plId is None:
                fields = { 'Field_NoNumber': isTrue(prjChp.noNumber)}
            else:
                fields = {'Field_ArcDefinition': self.novel.plotLines[plId].shortName}
            for field in fields:
                if fields[field]:
                    ET.SubElement(xmlChapterFields, field).text = fields[field]
            if plId is None and prjChp.chLevel == 1:
                ET.SubElement(xmlChapter, 'SectionStart').text = '-1'
                i += 1
            i = set_element(xmlChapter, 'Type', yType, i)
            i = set_element(xmlChapter, 'ChapterType', yChapterType, i)

        def build_location_subtree(xmlLoc, prjLoc):
            if prjLoc.title:
                ET.SubElement(xmlLoc, 'Title').text = prjLoc.title
            if prjLoc.desc:
                ET.SubElement(xmlLoc, 'Desc').text = prjLoc.desc
            if prjLoc.aka:
                ET.SubElement(xmlLoc, 'AKA').text = prjLoc.aka
            if prjLoc.tags:
                ET.SubElement(xmlLoc, 'Tags').text = list_to_string(prjLoc.tags)

        def add_projectvariable(title, desc, tags):
            pvId = create_id(prjVars)
            prjVars.append(pvId)
            xmlProjectvar = ET.SubElement(xmlProjectvars, 'PROJECTVAR')
            ET.SubElement(xmlProjectvar, 'ID').text = pvId
            ET.SubElement(xmlProjectvar, 'Title').text = title
            ET.SubElement(xmlProjectvar, 'Desc').text = desc
            ET.SubElement(xmlProjectvar, 'Tags').text = tags

        def build_item_subtree(xmlItm, prjItm):
            if prjItm.title:
                ET.SubElement(xmlItm, 'Title').text = prjItm.title
            if prjItm.desc:
                ET.SubElement(xmlItm, 'Desc').text = prjItm.desc
            if prjItm.aka:
                ET.SubElement(xmlItm, 'AKA').text = prjItm.aka
            if prjItm.tags:
                ET.SubElement(xmlItm, 'Tags').text = list_to_string(prjItm.tags)

        def build_character_subtree(xmlCrt, prjCrt):
            if prjCrt.title:
                ET.SubElement(xmlCrt, 'Title').text = prjCrt.title
            if prjCrt.desc:
                ET.SubElement(xmlCrt, 'Desc').text = prjCrt.desc
            if prjCrt.notes:
                ET.SubElement(xmlCrt, 'Notes').text = prjCrt.notes
            if prjCrt.aka:
                ET.SubElement(xmlCrt, 'AKA').text = prjCrt.aka
            if prjCrt.tags:
                ET.SubElement(xmlCrt, 'Tags').text = list_to_string(prjCrt.tags)
            if prjCrt.bio:
                ET.SubElement(xmlCrt, 'Bio').text = prjCrt.bio
            if prjCrt.goals:
                ET.SubElement(xmlCrt, 'Goals').text = prjCrt.goals
            if prjCrt.fullName:
                ET.SubElement(xmlCrt, 'FullName').text = prjCrt.fullName
            if prjCrt.isMajor:
                ET.SubElement(xmlCrt, 'Major').text = '-1'
            fields = {
                'Field_BirthDate': prjCrt.birthDate,
                'Field_DeathDate': prjCrt.deathDate,
                }
            xmlCrtFields = None
            for field in fields:
                if fields[field]:
                    if xmlCrtFields is None:
                        xmlCrtFields = ET.SubElement(xmlCrt, 'Fields')
                    ET.SubElement(xmlCrtFields, field).text = fields[field]

        def build_project_subtree(xmlProject):
            ET.SubElement(xmlProject, 'Ver').text = '7'
            if self.novel.title:
                ET.SubElement(xmlProject, 'Title').text = self.novel.title
            if self.novel.desc:
                ET.SubElement(xmlProject, 'Desc').text = self.novel.desc
            if self.novel.authorName:
                ET.SubElement(xmlProject, 'AuthorName').text = self.novel.authorName
            if self.novel.wordCountStart is not None:
                ET.SubElement(xmlProject, 'WordCountStart').text = str(self.novel.wordCountStart)
            if self.novel.wordTarget is not None:
                ET.SubElement(xmlProject, 'WordTarget').text = str(self.novel.wordTarget)

            workPhase = self.novel.workPhase
            if workPhase is not None:
                workPhase = str(workPhase)
            fields = {
                'Field_WorkPhase': workPhase,
                'Field_RenumberChapters': isTrue(self.novel.renumberChapters),
                'Field_RenumberParts': isTrue(self.novel.renumberParts),
                'Field_RenumberWithinParts': isTrue(self.novel.renumberWithinParts),
                'Field_RomanChapterNumbers': isTrue(self.novel.romanChapterNumbers),
                'Field_RomanPartNumbers': isTrue(self.novel.romanPartNumbers),
                'Field_ChapterHeadingPrefix': self.novel.chapterHeadingPrefix,
                'Field_ChapterHeadingSuffix': self.novel.chapterHeadingSuffix,
                'Field_PartHeadingPrefix': self.novel.partHeadingPrefix,
                'Field_PartHeadingSuffix': self.novel.partHeadingSuffix,
                'Field_CustomGoal': self.novel.customGoal,
                'Field_CustomConflict': self.novel.customConflict,
                'Field_CustomOutcome': self.novel.customOutcome,
                'Field_CustomChrBio': self.novel.customChrBio,
                'Field_CustomChrGoals': self.novel.customChrGoals,
                'Field_SaveWordCount': isTrue(self.novel.saveWordCount),
                'Field_ReferenceDate': self.novel.referenceDate,
                }
            xmlProjectFields = ET.SubElement(xmlProject, 'Fields')
            for field in fields:
                if fields[field]:
                    ET.SubElement(xmlProjectFields, field).text = fields[field]

        def build_prjNote_subtree(xmlProjectnote, projectNote):
            if projectNote.title is not None:
                ET.SubElement(xmlProjectnote, 'Title').text = projectNote.title

            if projectNote.desc is not None:
                ET.SubElement(xmlProjectnote, 'Desc').text = projectNote.desc

        root = ET.Element('YWRITER7')
        xmlProject = ET.SubElement(root, 'PROJECT')
        xmlLocations = ET.SubElement(root, 'LOCATIONS')
        xmlItems = ET.SubElement(root, 'ITEMS')
        xmlCharacters = ET.SubElement(root, 'CHARACTERS')
        xmlProjectvars = ET.SubElement(root, 'PROJECTVARS')
        xmlScenes = ET.SubElement(root, 'SCENES')
        xmlChapters = ET.SubElement(root, 'CHAPTERS')

        build_project_subtree(xmlProject)

        for lcId in self.novel.tree.get_children(LC_ROOT):
            xmlLoc = ET.SubElement(xmlLocations, 'LOCATION')
            ET.SubElement(xmlLoc, 'ID').text = lcId[2:]
            build_location_subtree(xmlLoc, self.novel.locations[lcId])

        for itId in self.novel.tree.get_children(IT_ROOT):
            xmlItm = ET.SubElement(xmlItems, 'ITEM')
            ET.SubElement(xmlItm, 'ID').text = itId[2:]
            build_item_subtree(xmlItm, self.novel.items[itId])

        for crId in self.novel.tree.get_children(CR_ROOT):
            xmlCrt = ET.SubElement(xmlCharacters, 'CHARACTER')
            ET.SubElement(xmlCrt, 'ID').text = crId[2:]
            build_character_subtree(xmlCrt, self.novel.characters[crId])

        if self.novel.languages or self.novel.languageCode or self.novel.countryCode:
            self.novel.check_locale()
            prjVars = []

            add_projectvariable('Language',
                                self.novel.languageCode,
                                '0')

            add_projectvariable('Country',
                                self.novel.countryCode,
                                '0')

            for langCode in self.novel.languages:
                add_projectvariable(f'lang={langCode}',
                                    f'<HTM <SPAN LANG="{langCode}"> /HTM>',
                                    '0')
                add_projectvariable(f'/lang={langCode}',
                                    f'<HTM </SPAN> /HTM>',
                                    '0')

        xmlSceneFields = {}
        scIds = list(self.novel.sections)
        for scId in scIds:
            xmlScene = ET.SubElement(xmlScenes, 'SCENE')
            ET.SubElement(xmlScene, 'ID').text = scId[2:]
            xmlSceneFields[scId] = ET.SubElement(xmlScene, 'Fields')
            build_scene_subtree(xmlScene, self.novel.sections[scId])

        newScIds = {}
        for ppId in self.novel.plotPoints:
            scId = create_id(scIds, prefix=SECTION_PREFIX)
            scIds.append(scId)
            newScIds[ppId] = scId
            xmlScene = ET.SubElement(xmlScenes, 'SCENE')
            ET.SubElement(xmlScene, 'ID').text = scId[2:]
            xmlSceneFields[scId] = ET.SubElement(xmlScene, 'Fields')
            build_scene_subtree(xmlScene, self.novel.plotPoints[ppId], plotPoint=True)

        chIds = list(self.novel.tree.get_children(CH_ROOT))
        for chId in chIds:
            xmlChapter = ET.SubElement(xmlChapters, 'CHAPTER')
            ET.SubElement(xmlChapter, 'ID').text = chId[2:]
            build_chapter_subtree(xmlChapter, self.novel.chapters[chId])
            srtScenes = self.novel.tree.get_children(chId)
            if srtScenes:
                xmlScnList = ET.SubElement(xmlChapter, 'Scenes')
                for scId in self.novel.tree.get_children(chId):
                    ET.SubElement(xmlScnList, 'ScID').text = scId[2:]

        chId = create_id(chIds, prefix=CHAPTER_PREFIX)
        chIds.append(chId)
        xmlChapter = ET.SubElement(xmlChapters, 'CHAPTER')
        ET.SubElement(xmlChapter, 'ID').text = chId[2:]
        arcPart = self._nvSvc.make_chapter(title=_('Plot lines'), chLevel=1)
        build_chapter_subtree(xmlChapter, arcPart, chType=2)
        for plId in self.novel.tree.get_children(PL_ROOT):
            chId = create_id(chIds, prefix=CHAPTER_PREFIX)
            chIds.append(chId)
            xmlChapter = ET.SubElement(xmlChapters, 'CHAPTER')
            ET.SubElement(xmlChapter, 'ID').text = chId[2:]
            build_chapter_subtree(xmlChapter, self.novel.plotLines[plId], plId=plId)
            srtScenes = self.novel.tree.get_children(plId)
            if srtScenes:
                xmlScnList = ET.SubElement(xmlChapter, 'Scenes')
                for ppId in srtScenes:
                    ET.SubElement(xmlScnList, 'ScID').text = newScIds[ppId][2:]

        if self.novel.tree.get_children(PN_ROOT):
            xmlProjectnotes = ET.SubElement(root, 'PROJECTNOTES')
            for pnId in self.novel.tree.get_children(PN_ROOT):
                xmlProjectnote = ET.SubElement(xmlProjectnotes, 'PROJECTNOTE')
                ET.SubElement(xmlProjectnote, 'ID').text = pnId[2:]
                build_prjNote_subtree(xmlProjectnote, self.novel.projectNotes[pnId])

        scPlotLines = {}
        sectionAssoc = {}
        for scId in scIds:
            scPlotLines[scId] = []
            sectionAssoc[scId] = []
        for plId in self.novel.plotLines:
            for scId in self.novel.plotLines[plId].sections:
                scPlotLines[scId].append(self.novel.plotLines[plId].shortName)
            for ppId in self.novel.tree.get_children(plId):
                scPlotLines[newScIds[ppId]].append(self.novel.plotLines[plId].shortName)
        for ppId in self.novel.plotPoints:
            if self.novel.plotPoints[ppId].sectionAssoc:
                sectionAssoc[self.novel.plotPoints[ppId].sectionAssoc].append(newScIds[ppId][2:])
                sectionAssoc[newScIds[ppId]].append(self.novel.plotPoints[ppId].sectionAssoc[2:])
        for scId in scIds:
            fields = {
                'Field_SceneArcs': list_to_string(scPlotLines[scId]),
                'Field_SceneAssoc': list_to_string(sectionAssoc[scId]),
                }
            for field in fields:
                if fields[field]:
                    ET.SubElement(xmlSceneFields[scId], field).text = fields[field]

        if self.wcLog:
            xmlWcLog = ET.SubElement(root, 'WCLog')
            wcLastCount = None
            wcLastTotalCount = None
            for wc in self.wcLog:
                if self.novel.saveWordCount:
                    if self.wcLog[wc][0] == wcLastCount and self.wcLog[wc][1] == wcLastTotalCount:
                        continue

                    wcLastCount = self.wcLog[wc][0]
                    wcLastTotalCount = self.wcLog[wc][1]
                xmlWc = ET.SubElement(xmlWcLog, 'WC')
                ET.SubElement(xmlWc, 'Date').text = wc
                ET.SubElement(xmlWc, 'Count').text = self.wcLog[wc][0]
                ET.SubElement(xmlWc, 'TotalCount').text = self.wcLog[wc][1]

        indent(root)
        self.tree = ET.ElementTree(root)

    def _convert_to_novx(self, text):

        def replace_note(match):
            noteType = match.group(1)
            self._noteCounter += 1
            self._noteNumber += 1
            noteLabel = f'{self._noteNumber}'
            if noteType.startswith('fn'):
                noteClass = 'footnote'
                if noteType.endswith('*'):
                    self._noteNumber -= 1
                    noteLabel = '*'
            elif noteType.startswith('en'):
                noteClass = 'endnote'
            return (f'<note id="ftn{self._noteCounter}" '
                    f'class="{noteClass}"><note-citation>{noteLabel}</note-citation>'
                    f'<p>{match.group(2)}</p></note>')

        def replace_comment(match):
            if self.novel.authorName:
                creator = self.novel.authorName
            else:
                creator = _('unknown')
            return (f'<comment><creator>{creator}</creator>'
                    f'<date>{datetime.today().replace(microsecond=0).isoformat()}</date>'
                    f'<p>{match.group(1)}</p></comment>')

        if not text:
            text = ''
        else:
            text = text.replace('<RTFBRK>', '')
            text = re.sub(r'\[\/*[h|c|r|s|u]\d*\]', '', text)
            for specialCode in ('HTM', 'TEX', 'RTF', 'epub', 'mobi', 'rtfimg'):
                text = re.sub(fr'\<{specialCode} .+?\/{specialCode}\>', '', text)

            xmlReplacements = [
                ('&', '&amp;'),
                ('>', '&gt;'),
                ('<', '&lt;'),
                ("'", '&apos;'),
                ('"', '&quot;'),
                ('\n', '</p><p>'),
                ('[i]', '<em>'),
                ('[/i]', '</em>'),
                ('[b]', '<strong>'),
                ('[/b]', '</strong>'),
                ]
            tags = ['i', 'b']
            if self.novel.languages is None:
                self.novel.get_languages()
            for language in self.novel.languages:
                tags.append(f'lang={language}')
                xmlReplacements.append((f'[lang={language}]', f'<span xml:lang="{language}">'))
                xmlReplacements.append((f'[/lang={language}]', '</span>'))

            newlines = []
            lines = text.split('\n')
            isOpen = {}
            opening = {}
            closing = {}
            for tag in tags:
                isOpen[tag] = False
                opening[tag] = f'[{tag}]'
                closing[tag] = f'[/{tag}]'
            for line in lines:
                for tag in tags:
                    if isOpen[tag]:
                        if line.startswith('&gt; '):
                            line = f"&gt; {opening[tag]}{line.lstrip('&gt; ')}"
                        else:
                            line = f'{opening[tag]}{line}'
                        isOpen[tag] = False
                    while line.count(opening[tag]) > line.count(closing[tag]):
                        line = f'{line}{closing[tag]}'
                        isOpen[tag] = True
                    while line.count(closing[tag]) > line.count(opening[tag]):
                        line = f'{opening[tag]}{line}'
                    line = line.replace(f'{opening[tag]}{closing[tag]}', '')
                newlines.append(line)
            text = '\n'.join(newlines).rstrip()

            for nv, od in xmlReplacements:
                text = text.replace(nv, od)

            if text.find('/*') > 0:
                text = re.sub(r'\/\* *@([ef]n\**) (.*?)\*\/', replace_note, text)
                text = re.sub(r'\/\*(.*?)\*\/', replace_comment, text)

            text = f'<p>{text}</p>'
            text = re.sub(r'\<p\>\&gt\; (.*?)\<\/p\>', '<p style="quotations">\\1</p>', text)
        return text

    def _postprocess_xml_file(self, filePath):
        with open(filePath, 'r', encoding='utf-8') as f:
            text = f.read()
        lines = text.split('\n')
        newlines = ['<?xml version="1.0" encoding="utf-8"?>']
        for line in lines:
            for tag in self._CDATA_TAGS:
                line = re.sub(fr'\<{tag}\>', f'<{tag}><![CDATA[', line)
                line = re.sub(fr'\<\/{tag}\>', f']]></{tag}>', line)
            newlines.append(line)
        text = '\n'.join(newlines)
        text = text.replace('[CDATA[ \n', '[CDATA[')
        text = text.replace('\n]]', ']]')
        if not self.novel.chapters:
            text = text.replace('<CHAPTERS />', '<CHAPTERS></CHAPTERS>')
        text = unescape(text)
        try:
            with open(filePath, 'w', encoding='utf-8') as f:
                f.write(text)
        except:
            raise Error(f'{_("Cannot write file")}: "{norm_path(filePath)}".')

    def _read_locations(self, root):
        self.novel.tree.delete_children(LC_ROOT)
        for xmlLocation in root.find('LOCATIONS'):
            lcId = f"{LOCATION_PREFIX}{xmlLocation.find('ID').text}"
            self.novel.tree.append(LC_ROOT, lcId)
            self.novel.locations[lcId] = self._nvSvc.make_world_element()

            if xmlLocation.find('Title') is not None:
                self.novel.locations[lcId].title = xmlLocation.find('Title').text


            if xmlLocation.find('Desc') is not None:
                self.novel.locations[lcId].desc = xmlLocation.find('Desc').text

            if xmlLocation.find('AKA') is not None:
                self.novel.locations[lcId].aka = xmlLocation.find('AKA').text

            if xmlLocation.find('Tags') is not None:
                if xmlLocation.find('Tags').text is not None:
                    tags = string_to_list(xmlLocation.find('Tags').text)
                    self.novel.locations[lcId].tags = self._strip_spaces(tags)

    def _read_items(self, root):
        self.novel.tree.delete_children(IT_ROOT)
        for xmlItem in root.find('ITEMS'):
            itId = f"{ITEM_PREFIX}{xmlItem.find('ID').text}"
            self.novel.tree.append(IT_ROOT, itId)
            self.novel.items[itId] = self._nvSvc.make_world_element()

            if xmlItem.find('Title') is not None:
                self.novel.items[itId].title = xmlItem.find('Title').text


            if xmlItem.find('Desc') is not None:
                self.novel.items[itId].desc = xmlItem.find('Desc').text

            if xmlItem.find('AKA') is not None:
                self.novel.items[itId].aka = xmlItem.find('AKA').text

            if xmlItem.find('Tags') is not None:
                if xmlItem.find('Tags').text is not None:
                    tags = string_to_list(xmlItem.find('Tags').text)
                    self.novel.items[itId].tags = self._strip_spaces(tags)

    def _read_chapters(self, root):
        self.novel.tree.delete_children(CH_ROOT)
        self.novel.tree.delete_children(PL_ROOT)
        for xmlChapter in root.find('CHAPTERS'):
            prjChapter = self._nvSvc.make_chapter()

            if xmlChapter.find('Title') is not None:
                prjChapter.title = xmlChapter.find('Title').text

            if xmlChapter.find('Desc') is not None:
                prjChapter.desc = xmlChapter.find('Desc').text

            if xmlChapter.find('SectionStart') is not None:
                prjChapter.chLevel = 1
            else:
                prjChapter.chLevel = 2


            prjChapter.chType = 0
            if xmlChapter.find('Unused') is not None:
                yUnused = True
            else:
                yUnused = False
            if xmlChapter.find('ChapterType') is not None:
                yChapterType = xmlChapter.find('ChapterType').text
                if yChapterType == '2':
                    prjChapter.chType = 1
                elif yChapterType == '1':
                    prjChapter.chType = 1
                elif yUnused:
                    prjChapter.chType = 1
            else:
                if xmlChapter.find('Type') is not None:
                    yType = xmlChapter.find('Type').text
                    if yType == '1':
                        prjChapter.chType = 1
                    elif yUnused:
                        prjChapter.chType = 1

            kwVarYw7 = {}
            for xmlChapterFields in xmlChapter.iterfind('Fields'):
                prjChapter.isTrash = False
                if xmlChapterFields.find('Field_IsTrash') is not None:
                    if xmlChapterFields.find('Field_IsTrash').text == '1':
                        prjChapter.isTrash = True

                for fieldName in self.CHP_KWVAR_YW7:
                    xmlField = xmlChapterFields.find(fieldName)
                    if xmlField  is not None:
                        kwVarYw7[fieldName] = xmlField .text
            prjChapter.noNumber = kwVarYw7.get('Field_NoNumber', False)
            shortName = kwVarYw7.get('Field_ArcDefinition', '')

            field = kwVarYw7.get('Field_Arc_Definition', None)
            if field is not None:
                shortName = field

            scenes = []
            if xmlChapter.find('Scenes') is not None:
                for scn in xmlChapter.find('Scenes').iterfind('ScID'):
                    scId = scn.text
                    scenes.append(scId)

            if shortName:
                plId = f"{PLOT_LINE_PREFIX}{xmlChapter.find('ID').text}"
                self.novel.plotLines[plId] = self._nvSvc.make_plot_line()
                self.novel.plotLines[plId].title = prjChapter.title
                self.novel.plotLines[plId].desc = prjChapter.desc
                self.novel.plotLines[plId].shortName = shortName
                self.novel.tree.append(PL_ROOT, plId)
                for scId in scenes:
                    self.novel.tree.append(plId, f'{PLOT_POINT_PREFIX}{scId}')
                    self._ywApIds.append(scId)
            else:
                chId = f"{CHAPTER_PREFIX}{xmlChapter.find('ID').text}"
                self.novel.chapters[chId] = prjChapter
                self.novel.tree.append(CH_ROOT, chId)
                for scId in scenes:
                    self.novel.tree.append(chId, f'{SECTION_PREFIX}{scId}')

    def _read_characters(self, root):
        self.novel.tree.delete_children(CR_ROOT)
        for xmlCharacter in root.find('CHARACTERS'):
            crId = f"{CHARACTER_PREFIX}{xmlCharacter.find('ID').text}"
            self.novel.tree.append(CR_ROOT, crId)
            self.novel.characters[crId] = self._nvSvc.make_character()

            if xmlCharacter.find('Title') is not None:
                self.novel.characters[crId].title = xmlCharacter.find('Title').text


            if xmlCharacter.find('Desc') is not None:
                self.novel.characters[crId].desc = xmlCharacter.find('Desc').text

            if xmlCharacter.find('AKA') is not None:
                self.novel.characters[crId].aka = xmlCharacter.find('AKA').text

            if xmlCharacter.find('Tags') is not None:
                if xmlCharacter.find('Tags').text is not None:
                    tags = string_to_list(xmlCharacter.find('Tags').text)
                    self.novel.characters[crId].tags = self._strip_spaces(tags)

            if xmlCharacter.find('Notes') is not None:
                self.novel.characters[crId].notes = xmlCharacter.find('Notes').text

            if xmlCharacter.find('Bio') is not None:
                self.novel.characters[crId].bio = xmlCharacter.find('Bio').text

            if xmlCharacter.find('Goals') is not None:
                self.novel.characters[crId].goals = xmlCharacter.find('Goals').text

            if xmlCharacter.find('FullName') is not None:
                self.novel.characters[crId].fullName = xmlCharacter.find('FullName').text

            if xmlCharacter.find('Major') is not None:
                self.novel.characters[crId].isMajor = True
            else:
                self.novel.characters[crId].isMajor = False

            kwVarYw7 = {}
            xmlCharacterFields = xmlCharacter.find('Fields')
            if xmlCharacterFields is not None:
                for fieldName in self.CRT_KWVAR_YW7:
                    xmlField = xmlCharacterFields.find(fieldName)
                    if xmlField  is not None:
                        kwVarYw7[fieldName] = xmlField.text
            self.novel.characters[crId].birthDate = kwVarYw7.get('Field_BirthDate', '')
            self.novel.characters[crId].deathDate = kwVarYw7.get('Field_DeathDate', '')

    def _read_project(self, root):
        xmlProject = root.find('PROJECT')

        if xmlProject.find('Title') is not None:
            self.novel.title = xmlProject.find('Title').text

        if xmlProject.find('AuthorName') is not None:
            self.novel.authorName = xmlProject.find('AuthorName').text

        if xmlProject.find('Desc') is not None:
            self.novel.desc = xmlProject.find('Desc').text

        if xmlProject.find('WordCountStart') is not None:
            try:
                self.novel.wordCountStart = int(xmlProject.find('WordCountStart').text)
            except:
                pass
        if xmlProject.find('WordTarget') is not None:
            try:
                self.novel.wordTarget = int(xmlProject.find('WordTarget').text)
            except:
                pass

        kwVarYw7 = {}
        for xmlProjectFields in xmlProject.iterfind('Fields'):
            for fieldName in self.PRJ_KWVAR_YW7:
                xmlField = xmlProjectFields.find(fieldName)
                if xmlField  is not None:
                    kwVarYw7[fieldName] = xmlField .text
        self.novel.workPhase = kwVarYw7.get('Field_WorkPhase', None)
        self.novel.renumberChapters = kwVarYw7.get('Field_RenumberChapters', False)
        self.novel.renumberParts = kwVarYw7.get('Field_RenumberParts', False)
        self.novel.renumberWithinParts = kwVarYw7.get('Field_RenumberWithinParts', False)
        self.novel.romanChapterNumbers = kwVarYw7.get('Field_RomanChapterNumbers', False)
        self.novel.romanPartNumbers = kwVarYw7.get('Field_RomanPartNumbers', False)
        self.novel.chapterHeadingPrefix = kwVarYw7.get('Field_ChapterHeadingPrefix', '')
        self.novel.chapterHeadingSuffix = kwVarYw7.get('Field_ChapterHeadingSuffix', '')
        self.novel.partHeadingPrefix = kwVarYw7.get('Field_PartHeadingPrefix', '')
        self.novel.partHeadingSuffix = kwVarYw7.get('Field_PartHeadingSuffix', '')
        self.novel.customGoal = kwVarYw7.get('Field_CustomGoal', '')
        self.novel.customConflict = kwVarYw7.get('Field_CustomConflict', '')
        self.novel.customOutcome = kwVarYw7.get('Field_CustomOutcome', '')
        self.novel.customChrBio = kwVarYw7.get('Field_CustomChrBio', '')
        self.novel.customChrGoals = kwVarYw7.get('Field_CustomChrGoals', '')
        self.novel.saveWordCount = kwVarYw7.get('Field_SaveWordCount', False)

        field = kwVarYw7.get('Field_LanguageCode', None)
        if field is not None:
            self.novel.languageCode = field
        field = kwVarYw7.get('Field_CountryCode', None)
        if field is not None:
            self.novel.countryCode = field

    def _read_project_notes(self, root):
        if root.find('PROJECTNOTES') is not None:
            for xmlProjectnote in root.find('PROJECTNOTES'):
                if xmlProjectnote.find('ID') is not None:
                    pnId = f"{PRJ_NOTE_PREFIX}{xmlProjectnote.find('ID').text}"
                    self.novel.tree.append(PN_ROOT, pnId)
                    self.novel.projectNotes[pnId] = self._nvSvc.make_basic_element()
                    if xmlProjectnote.find('Title') is not None:
                        self.novel.projectNotes[pnId].title = xmlProjectnote.find('Title').text
                    if xmlProjectnote.find('Desc') is not None:
                        self.novel.projectNotes[pnId].desc = xmlProjectnote.find('Desc').text

    def _read_projectvars(self, root):
        try:
            for xmlProjectvar in root.find('PROJECTVARS'):
                if xmlProjectvar.find('Title') is not None:
                    title = xmlProjectvar.find('Title').text
                    if title == 'Language':
                        if xmlProjectvar.find('Desc') is not None:
                            self.novel.languageCode = xmlProjectvar.find('Desc').text

                    elif title == 'Country':
                        if xmlProjectvar.find('Desc') is not None:
                            self.novel.countryCode = xmlProjectvar.find('Desc').text

                    elif title.startswith('lang='):
                        try:
                            __, langCode = title.split('=')
                            if self.novel.languages is None:
                                self.novel.languages = []
                            self.novel.languages.append(langCode)
                        except:
                            pass
        except:
            pass

    def _read_scenes(self, root):
        for xmlScene in root.find('SCENES'):
            prjScn = self._nvSvc.make_section()

            if xmlScene.find('Title') is not None:
                prjScn.title = xmlScene.find('Title').text

            if xmlScene.find('Desc') is not None:
                prjScn.desc = xmlScene.find('Desc').text

            if xmlScene.find('SceneContent') is not None:
                sceneContent = xmlScene.find('SceneContent').text
                if sceneContent is not None:
                    prjScn.sectionContent = self._convert_to_novx(sceneContent)



            prjScn.scType = 0
            kwVarYw7 = {}
            for xmlSceneFields in xmlScene.iterfind('Fields'):
                if xmlSceneFields.find('Field_SceneType') is not None:
                    if xmlSceneFields.find('Field_SceneType').text == '1':
                        prjScn.scType = 1
                    elif xmlSceneFields.find('Field_SceneType').text == '2':
                        prjScn.scType = 1

                for fieldName in self.SCN_KWVAR_YW7:
                    xmlField = xmlSceneFields.find(fieldName)
                    if xmlField  is not None:
                        kwVarYw7[fieldName] = xmlField.text

            ywScnArcs = string_to_list(kwVarYw7.get('Field_SceneArcs', ''))
            for shortName in ywScnArcs:
                for plId in self.novel.plotLines:
                    if self.novel.plotLines[plId].shortName == shortName:
                        if prjScn.scType == 0:
                            arcSections = self.novel.plotLines[plId].sections
                            if not arcSections:
                                arcSections = [f"{SECTION_PREFIX}{xmlScene.find('ID').text}"]
                            else:
                                arcSections.append(f"{SECTION_PREFIX}{xmlScene.find('ID').text}")
                            self.novel.plotLines[plId].sections = arcSections
                        break

            ywScnAssocs = string_to_list(kwVarYw7.get('Field_SceneAssoc', ''))
            prjScn.plotPoints = [f'{PLOT_POINT_PREFIX}{plotPoint}' for plotPoint in ywScnAssocs]

            if xmlScene.find('Goal') is not None:
                prjScn.goal = xmlScene.find('Goal').text

            if xmlScene.find('Conflict') is not None:
                prjScn.conflict = xmlScene.find('Conflict').text

            if xmlScene.find('Outcome') is not None:
                prjScn.outcome = xmlScene.find('Outcome').text

            if kwVarYw7.get('Field_CustomAR', None) is not None:
                prjScn.scene = 3
            elif xmlScene.find('ReactionScene') is not None:
                prjScn.scene = 2
            elif prjScn.goal or prjScn.conflict or prjScn.outcome:
                prjScn.scene = 1
            else:
                prjScn.scene = 0

            if xmlScene.find('Unused') is not None:
                if prjScn.scType == 0:
                    prjScn.scType = 1

            if xmlScene.find('Status') is not None:
                prjScn.status = int(xmlScene.find('Status').text)

            if xmlScene.find('Notes') is not None:
                prjScn.notes = xmlScene.find('Notes').text

            if xmlScene.find('Tags') is not None:
                if xmlScene.find('Tags').text is not None:
                    tags = string_to_list(xmlScene.find('Tags').text)
                    prjScn.tags = self._strip_spaces(tags)

            if xmlScene.find('AppendToPrev') is not None:
                prjScn.appendToPrev = True
            else:
                prjScn.appendToPrev = False

            if xmlScene.find('SpecificDateTime') is not None:
                dateTimeStr = xmlScene.find('SpecificDateTime').text

                try:
                    dateTime = datetime.fromisoformat(dateTimeStr)
                except:
                    prjScn.date = ''
                    prjScn.time = ''
                else:
                    startDateTime = dateTime.isoformat().split('T')
                    prjScn.date = startDateTime[0]
                    prjScn.time = startDateTime[1]
            else:
                if xmlScene.find('Day') is not None:
                    day = xmlScene.find('Day').text

                    try:
                        int(day)
                    except ValueError:
                        day = ''
                    prjScn.day = day

                hasUnspecificTime = False
                if xmlScene.find('Hour') is not None:
                    hour = xmlScene.find('Hour').text.zfill(2)
                    hasUnspecificTime = True
                else:
                    hour = '00'
                if xmlScene.find('Minute') is not None:
                    minute = xmlScene.find('Minute').text.zfill(2)
                    hasUnspecificTime = True
                else:
                    minute = '00'
                if hasUnspecificTime:
                    prjScn.time = f'{hour}:{minute}:00'

            if xmlScene.find('LastsDays') is not None:
                prjScn.lastsDays = xmlScene.find('LastsDays').text

            if xmlScene.find('LastsHours') is not None:
                prjScn.lastsHours = xmlScene.find('LastsHours').text

            if xmlScene.find('LastsMinutes') is not None:
                prjScn.lastsMinutes = xmlScene.find('LastsMinutes').text


            scCharacters = []
            if xmlScene.find('Characters') is not None:
                for character in xmlScene.find('Characters').iter('CharID'):
                    crId = f"{CHARACTER_PREFIX}{character.text}"
                    if crId in self.novel.tree.get_children(CR_ROOT):
                        scCharacters.append(crId)
            prjScn.characters = scCharacters

            scLocations = []
            if xmlScene.find('Locations') is not None:
                for location in xmlScene.find('Locations').iter('LocID'):
                    lcId = f"{LOCATION_PREFIX}{location.text}"
                    if lcId in self.novel.tree.get_children(LC_ROOT):
                        scLocations.append(lcId)
            prjScn.locations = scLocations

            scItems = []
            if xmlScene.find('Items') is not None:
                for item in xmlScene.find('Items').iter('ItemID'):
                    itId = f"{ITEM_PREFIX}{item.text}"
                    if itId in self.novel.tree.get_children(IT_ROOT):
                        scItems.append(itId)
            prjScn.items = scItems

            ywScId = xmlScene.find('ID').text
            if ywScId in self._ywApIds:
                ppId = f"{PLOT_POINT_PREFIX}{ywScId}"
                self.novel.plotPoints[ppId] = self._nvSvc.make_plot_point(title=prjScn.title,
                                                      desc=prjScn.desc
                                                      )
                if ywScnAssocs:
                    self.novel.plotPoints[ppId].sectionAssoc = f'{SECTION_PREFIX}{ywScnAssocs[0]}'
            else:
                if prjScn.tags and self.STAGE_MARKER in prjScn.tags:
                    prjScn.scType = 3
                    prjScn.tags = prjScn.tags.remove(self.STAGE_MARKER)
                scId = f"{SECTION_PREFIX}{ywScId}"
                self.novel.sections[scId] = prjScn

    def _strip_spaces(self, lines):
        stripped = []
        for line in lines:
            stripped.append(line.strip())
        return stripped

    def _write_element_tree(self, ywProject):
        backedUp = False
        if os.path.isfile(ywProject.filePath):
            try:
                os.replace(ywProject.filePath, f'{ywProject.filePath}.bak')
            except:
                raise Error(f'{_("Cannot overwrite file")}: "{norm_path(ywProject.filePath)}".')
            else:
                backedUp = True
        try:
            ywProject.tree.write(ywProject.filePath, xml_declaration=False, encoding='utf-8')
        except:
            if backedUp:
                os.replace(f'{ywProject.filePath}.bak', ywProject.filePath)
            raise Error(f'{_("Cannot write file")}: "{norm_path(ywProject.filePath)}".')

from configparser import ConfigParser


class Configuration:

    def __init__(self, settings={}, options={}):
        self.settings = None
        self.options = None
        self._sLabel = 'SETTINGS'
        self._oLabel = 'OPTIONS'
        self.set(settings, options)

    def read(self, iniFile):
        config = ConfigParser()
        config.read(iniFile, encoding='utf-8')
        if config.has_section(self._sLabel):
            section = config[self._sLabel]
            for setting in self.settings:
                fallback = self.settings[setting]
                self.settings[setting] = section.get(setting, fallback)
        if config.has_section(self._oLabel):
            section = config[self._oLabel]
            for option in self.options:
                fallback = self.options[option]
                self.options[option] = section.getboolean(option, fallback)

    def set(self, settings=None, options=None):
        if settings is not None:
            self.settings = settings.copy()
        if options is not None:
            self.options = options.copy()

    def write(self, iniFile):
        config = ConfigParser()
        if self.settings:
            config.add_section(self._sLabel)
            for settingId in self.settings:
                config.set(self._sLabel, settingId, str(self.settings[settingId]))
        if self.options:
            config.add_section(self._oLabel)
            for settingId in self.options:
                if self.options[settingId]:
                    config.set(self._oLabel, settingId, 'Yes')
                else:
                    config.set(self._oLabel, settingId, 'No')
        with open(iniFile, 'w', encoding='utf-8') as f:
            config.write(f)
import math


def get_moon_phase_day(isoDate):
    try:
        y, m, d = isoDate.split('-')
        year = int(y)
        month = int(m)
        day = int(d)
        r = year % 100
        r %= 19
        if r > 9:
            r -= 19
        r = ((r * 11) % 30) + month + day
        if month < 3:
            r += 2
        if year < 2000:
            r -= 4
        else:
            r -= 8.3
        r = math.floor(r + 0.5) % 30
        if r < 0:
            r += 30
    except:
        r = None
    return r


def get_moon_phase_string(isoDate):
    moonViews = [
        '🌑',
        '🌑',
        '🌒',
        '🌒',
        '🌒',
        '🌒',
        '🌓',
        '🌓',
        '🌓',
        '🌓',
        '🌔',
        '🌔',
        '🌔',
        '🌔',
        '🌕',
        '🌕',
        '🌕',
        '🌖',
        '🌖',
        '🌖',
        '🌖',
        '🌗',
        '🌗',
        '🌗',
        '🌗',
        '🌘',
        '🌘',
        '🌘',
        '🌘',
        '🌑'
    ]
    moonFractions = '00¼¼¼¼½½½½¾¾¾¾111¾¾¾¾½½½½¼¼¼¼0'
    moonPhaseDay = get_moon_phase_day(isoDate)
    if moonPhaseDay is not None:
        display = f'{moonPhaseDay} {moonViews[moonPhaseDay]} {moonFractions[moonPhaseDay]}'
    else:
        display = ''
    return display

from datetime import date



class BasicElement:

    def __init__(self,
            on_element_change=None,
            title=None,
            desc=None,
            links=None):
        if on_element_change is None:
            self.on_element_change = self.do_nothing
        else:
            self.on_element_change = on_element_change
        self._title = title
        self._desc = desc
        self._links = links

    @property
    def title(self):
        return self._title

    @title.setter
    def title(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._title != newVal:
            self._title = newVal
            self.on_element_change()

    @property
    def desc(self):
        return self._desc

    @desc.setter
    def desc(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._desc != newVal:
            self._desc = newVal
            self.on_element_change()

    @property
    def links(self):
        try:
            return self._links.copy()
        except AttributeError:
            return None

    @links.setter
    def links(self, newVal):
        if newVal is not None:
            for elem in newVal:
                val = newVal[elem]
                if val is not None:
                    assert type(val) == str
        if self._links != newVal:
            self._links = newVal
            self.on_element_change()

    def do_nothing(self):
        pass

    def from_xml(self, xmlElement):
        self.title = self._get_element_text(xmlElement, 'Title')
        self.desc = self._xml_element_to_text(xmlElement.find('Desc'))
        self.links = self._get_link_dict(xmlElement)

    def to_xml(self, xmlElement):
        if self.title:
            ET.SubElement(xmlElement, 'Title').text = self.title
        if self.desc:
            xmlElement.append(self._text_to_xml_element('Desc', self.desc))
        if self.links:
            for path in self.links:
                xmlLink = ET.SubElement(xmlElement, 'Link')
                ET.SubElement(xmlLink, 'Path').text = path
                if self.links[path]:
                    ET.SubElement(xmlLink, 'FullPath').text = self.links[path]

    def _get_element_text(self, xmlElement, tag, default=None):
        if xmlElement.find(tag) is not None:
            return xmlElement.find(tag).text
        else:
            return default

    def _get_link_dict(self, xmlElement):
        links = {}
        for xmlLink in xmlElement.iterfind('Link'):
            xmlPath = xmlLink.find('Path')
            if xmlPath is not None:
                path = xmlPath.text
                xmlFullPath = xmlLink.find('FullPath')
                if xmlFullPath is not None:
                    fullPath = xmlFullPath.text
                else:
                    fullPath = None
            else:
                path = xmlLink.attrib.get('path', None)
                fullPath = xmlLink.attrib.get('fullPath', None)
            if path:
                links[path] = fullPath
        return links

    def _text_to_xml_element(self, tag, text):
        xmlElement = ET.Element(tag)
        if text:
            for line in text.split('\n'):
                ET.SubElement(xmlElement, 'p').text = line
        return xmlElement

    def _xml_element_to_text(self, xmlElement):
        lines = []
        if xmlElement is not None:
            for paragraph in xmlElement.iterfind('p'):
                lines.append(''.join(t for t in paragraph.itertext()))
        return '\n'.join(lines)


LANGUAGE_TAG = re.compile(r'\<span xml\:lang=\"(.*?)\"\>')


class Novel(BasicElement):

    def __init__(self,
            authorName=None,
            wordTarget=None,
            wordCountStart=None,
            languageCode=None,
            countryCode=None,
            renumberChapters=None,
            renumberParts=None,
            renumberWithinParts=None,
            romanChapterNumbers=None,
            romanPartNumbers=None,
            saveWordCount=None,
            workPhase=None,
            chapterHeadingPrefix=None,
            chapterHeadingSuffix=None,
            partHeadingPrefix=None,
            partHeadingSuffix=None,
            customPlotProgress=None,
            customCharacterization=None,
            customWorldBuilding=None,
            customGoal=None,
            customConflict=None,
            customOutcome=None,
            customChrBio=None,
            customChrGoals=None,
            referenceDate=None,
            tree=None,
            **kwargs):
        super().__init__(**kwargs)
        self._authorName = authorName
        self._wordTarget = wordTarget
        self._wordCountStart = wordCountStart
        self._languageCode = languageCode
        self._countryCode = countryCode
        self._renumberChapters = renumberChapters
        self._renumberParts = renumberParts
        self._renumberWithinParts = renumberWithinParts
        self._romanChapterNumbers = romanChapterNumbers
        self._romanPartNumbers = romanPartNumbers
        self._saveWordCount = saveWordCount
        self._workPhase = workPhase
        self._chapterHeadingPrefix = chapterHeadingPrefix
        self._chapterHeadingSuffix = chapterHeadingSuffix
        self._partHeadingPrefix = partHeadingPrefix
        self._partHeadingSuffix = partHeadingSuffix
        self._customPlotProgress = customPlotProgress
        self._customCharacterization = customCharacterization
        self._customWorldBuilding = customWorldBuilding
        self._customGoal = customGoal
        self._customConflict = customConflict
        self._customOutcome = customOutcome
        self._customChrBio = customChrBio
        self._customChrGoals = customChrGoals

        self.chapters = {}
        self.sections = {}
        self.plotPoints = {}
        self.languages = None
        self.plotLines = {}
        self.locations = {}
        self.items = {}
        self.characters = {}
        self.projectNotes = {}
        try:
            self.referenceWeekDay = date.fromisoformat(referenceDate).weekday()
            self._referenceDate = referenceDate
        except:
            self.referenceWeekDay = None
            self._referenceDate = None
        self.tree = tree

    @property
    def authorName(self):
        return self._authorName

    @authorName.setter
    def authorName(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._authorName != newVal:
            self._authorName = newVal
            self.on_element_change()

    @property
    def wordTarget(self):
        return self._wordTarget

    @wordTarget.setter
    def wordTarget(self, newVal):
        if newVal is not None:
            assert type(newVal) == int
        if self._wordTarget != newVal:
            self._wordTarget = newVal
            self.on_element_change()

    @property
    def wordCountStart(self):
        return self._wordCountStart

    @wordCountStart.setter
    def wordCountStart(self, newVal):
        if newVal is not None:
            assert type(newVal) == int
        if self._wordCountStart != newVal:
            self._wordCountStart = newVal
            self.on_element_change()

    @property
    def languageCode(self):
        return self._languageCode

    @languageCode.setter
    def languageCode(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._languageCode != newVal:
            self._languageCode = newVal
            self.on_element_change()

    @property
    def countryCode(self):
        return self._countryCode

    @countryCode.setter
    def countryCode(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._countryCode != newVal:
            self._countryCode = newVal
            self.on_element_change()

    @property
    def renumberChapters(self):
        return self._renumberChapters

    @renumberChapters.setter
    def renumberChapters(self, newVal):
        if newVal is not None:
            assert type(newVal) == bool
        if self._renumberChapters != newVal:
            self._renumberChapters = newVal
            self.on_element_change()

    @property
    def renumberParts(self):
        return self._renumberParts

    @renumberParts.setter
    def renumberParts(self, newVal):
        if newVal is not None:
            assert type(newVal) == bool
        if self._renumberParts != newVal:
            self._renumberParts = newVal
            self.on_element_change()

    @property
    def renumberWithinParts(self):
        return self._renumberWithinParts

    @renumberWithinParts.setter
    def renumberWithinParts(self, newVal):
        if newVal is not None:
            assert type(newVal) == bool
        if self._renumberWithinParts != newVal:
            self._renumberWithinParts = newVal
            self.on_element_change()

    @property
    def romanChapterNumbers(self):
        return self._romanChapterNumbers

    @romanChapterNumbers.setter
    def romanChapterNumbers(self, newVal):
        if newVal is not None:
            assert type(newVal) == bool
        if self._romanChapterNumbers != newVal:
            self._romanChapterNumbers = newVal
            self.on_element_change()

    @property
    def romanPartNumbers(self):
        return self._romanPartNumbers

    @romanPartNumbers.setter
    def romanPartNumbers(self, newVal):
        if newVal is not None:
            assert type(newVal) == bool
        if self._romanPartNumbers != newVal:
            self._romanPartNumbers = newVal
            self.on_element_change()

    @property
    def saveWordCount(self):
        return self._saveWordCount

    @saveWordCount.setter
    def saveWordCount(self, newVal):
        if newVal is not None:
            assert type(newVal) == bool
        if self._saveWordCount != newVal:
            self._saveWordCount = newVal
            self.on_element_change()

    @property
    def workPhase(self):
        return self._workPhase

    @workPhase.setter
    def workPhase(self, newVal):
        if newVal is not None:
            assert type(newVal) == int
        if self._workPhase != newVal:
            self._workPhase = newVal
            self.on_element_change()

    @property
    def chapterHeadingPrefix(self):
        return self._chapterHeadingPrefix

    @chapterHeadingPrefix.setter
    def chapterHeadingPrefix(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._chapterHeadingPrefix != newVal:
            self._chapterHeadingPrefix = newVal
            self.on_element_change()

    @property
    def chapterHeadingSuffix(self):
        return self._chapterHeadingSuffix

    @chapterHeadingSuffix.setter
    def chapterHeadingSuffix(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._chapterHeadingSuffix != newVal:
            self._chapterHeadingSuffix = newVal
            self.on_element_change()

    @property
    def partHeadingPrefix(self):
        return self._partHeadingPrefix

    @partHeadingPrefix.setter
    def partHeadingPrefix(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._partHeadingPrefix != newVal:
            self._partHeadingPrefix = newVal
            self.on_element_change()

    @property
    def partHeadingSuffix(self):
        return self._partHeadingSuffix

    @partHeadingSuffix.setter
    def partHeadingSuffix(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._partHeadingSuffix != newVal:
            self._partHeadingSuffix = newVal
            self.on_element_change()

    @property
    def customPlotProgress(self):
        return self._customPlotProgress

    @customPlotProgress.setter
    def customPlotProgress(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._customPlotProgress != newVal:
            self._customPlotProgress = newVal
            self.on_element_change()

    @property
    def customCharacterization(self):
        return self._customCharacterization

    @customCharacterization.setter
    def customCharacterization(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._customCharacterization != newVal:
            self._customCharacterization = newVal
            self.on_element_change()

    @property
    def customWorldBuilding(self):
        return self._customWorldBuilding

    @customWorldBuilding.setter
    def customWorldBuilding(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._customWorldBuilding != newVal:
            self._customWorldBuilding = newVal
            self.on_element_change()

    @property
    def customGoal(self):
        return self._customGoal

    @customGoal.setter
    def customGoal(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._customGoal != newVal:
            self._customGoal = newVal
            self.on_element_change()

    @property
    def customConflict(self):
        return self._customConflict

    @customConflict.setter
    def customConflict(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._customConflict != newVal:
            self._customConflict = newVal
            self.on_element_change()

    @property
    def customOutcome(self):
        return self._customOutcome

    @customOutcome.setter
    def customOutcome(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._customOutcome != newVal:
            self._customOutcome = newVal
            self.on_element_change()

    @property
    def customChrBio(self):
        return self._customChrBio

    @customChrBio.setter
    def customChrBio(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._customChrBio != newVal:
            self._customChrBio = newVal
            self.on_element_change()

    @property
    def customChrGoals(self):
        return self._customChrGoals

    @customChrGoals.setter
    def customChrGoals(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._customChrGoals != newVal:
            self._customChrGoals = newVal
            self.on_element_change()

    @property
    def referenceDate(self):
        return self._referenceDate

    @referenceDate.setter
    def referenceDate(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._referenceDate != newVal:
            if not newVal:
                self._referenceDate = None
                self.referenceWeekDay = None
                self.on_element_change()
            else:
                try:
                    self.referenceWeekDay = date.fromisoformat(newVal).weekday()
                except:
                    pass
                else:
                    self._referenceDate = newVal
                    self.on_element_change()

    def check_locale(self):
        if not self._languageCode or self._languageCode == 'None':
            try:
                sysLng, sysCtr = locale.getlocale()[0].split('_')
            except:
                sysLng, sysCtr = locale.getdefaultlocale()[0].split('_')
            self._languageCode = sysLng
            self._countryCode = sysCtr
            self.on_element_change()
            return

        try:
            if len(self._languageCode) == 2:
                if len(self._countryCode) == 2:
                    return
        except:
            pass
        self._languageCode = 'zxx'
        self._countryCode = 'none'
        self.on_element_change()

    def from_xml(self, xmlElement):
        super().from_xml(xmlElement)
        self.renumberChapters = xmlElement.get('renumberChapters', None) == '1'
        self.renumberParts = xmlElement.get('renumberParts', None) == '1'
        self.renumberWithinParts = xmlElement.get('renumberWithinParts', None) == '1'
        self.romanChapterNumbers = xmlElement.get('romanChapterNumbers', None) == '1'
        self.romanPartNumbers = xmlElement.get('romanPartNumbers', None) == '1'
        self.saveWordCount = xmlElement.get('saveWordCount', None) == '1'
        workPhase = xmlElement.get('workPhase', None)
        if workPhase in ('1', '2', '3', '4', '5'):
            self.workPhase = int(workPhase)
        else:
            self.workPhase = None

        self.authorName = self._get_element_text(xmlElement, 'Author')

        self.chapterHeadingPrefix = self._get_element_text(xmlElement, 'ChapterHeadingPrefix')
        self.chapterHeadingSuffix = self._get_element_text(xmlElement, 'ChapterHeadingSuffix')

        self.partHeadingPrefix = self._get_element_text(xmlElement, 'PartHeadingPrefix')
        self.partHeadingSuffix = self._get_element_text(xmlElement, 'PartHeadingSuffix')

        self.customPlotProgress = self._get_element_text(xmlElement, 'CustomPlotProgress')
        self.customCharacterization = self._get_element_text(xmlElement, 'CustomCharacterization')
        self.customWorldBuilding = self._get_element_text(xmlElement, 'CustomWorldBuilding')

        self.customGoal = self._get_element_text(xmlElement, 'CustomGoal')
        self.customConflict = self._get_element_text(xmlElement, 'CustomConflict')
        self.customOutcome = self._get_element_text(xmlElement, 'CustomOutcome')

        self.customChrBio = self._get_element_text(xmlElement, 'CustomChrBio')
        self.customChrGoals = self._get_element_text(xmlElement, 'CustomChrGoals')

        if xmlElement.find('WordCountStart') is not None:
            self.wordCountStart = int(xmlElement.find('WordCountStart').text)
        if xmlElement.find('WordTarget') is not None:
            self.wordTarget = int(xmlElement.find('WordTarget').text)

        self.referenceDate = verified_date(self._get_element_text(xmlElement, 'ReferenceDate'))

    def get_languages(self):

        def languages(text):
            if not text:
                return

            m = LANGUAGE_TAG.search(text)
            while m:
                text = text[m.span()[1]:]
                yield m.group(1)
                m = LANGUAGE_TAG.search(text)

        self.languages = []
        for scId in self.sections:
            text = self.sections[scId].sectionContent
            if not text:
                continue

            for language in languages(text):
                if not language in self.languages:
                    self.languages.append(language)

    def to_xml(self, xmlElement):
        super().to_xml(xmlElement)
        if self.renumberChapters:
            xmlElement.set('renumberChapters', '1')
        if self.renumberParts:
            xmlElement.set('renumberParts', '1')
        if self.renumberWithinParts:
            xmlElement.set('renumberWithinParts', '1')
        if self.romanChapterNumbers:
            xmlElement.set('romanChapterNumbers', '1')
        if self.romanPartNumbers:
            xmlElement.set('romanPartNumbers', '1')
        if self.saveWordCount:
            xmlElement.set('saveWordCount', '1')
        if self.workPhase is not None:
            xmlElement.set('workPhase', str(self.workPhase))

        if self.authorName:
            ET.SubElement(xmlElement, 'Author').text = self.authorName

        if self.chapterHeadingPrefix:
            ET.SubElement(xmlElement, 'ChapterHeadingPrefix').text = self.chapterHeadingPrefix
        if self.chapterHeadingSuffix:
            ET.SubElement(xmlElement, 'ChapterHeadingSuffix').text = self.chapterHeadingSuffix

        if self.partHeadingPrefix:
            ET.SubElement(xmlElement, 'PartHeadingPrefix').text = self.partHeadingPrefix
        if self.partHeadingSuffix:
            ET.SubElement(xmlElement, 'PartHeadingSuffix').text = self.partHeadingSuffix

        if self.customPlotProgress:
            ET.SubElement(xmlElement, 'CustomPlotProgress').text = self.customPlotProgress
        if self.customCharacterization:
            ET.SubElement(xmlElement, 'CustomCharacterization').text = self.customCharacterization
        if self.customWorldBuilding:
            ET.SubElement(xmlElement, 'CustomWorldBuilding').text = self.customWorldBuilding

        if self.customGoal:
            ET.SubElement(xmlElement, 'CustomGoal').text = self.customGoal
        if self.customConflict:
            ET.SubElement(xmlElement, 'CustomConflict').text = self.customConflict
        if self.customOutcome:
            ET.SubElement(xmlElement, 'CustomOutcome').text = self.customOutcome

        if self.customChrBio:
            ET.SubElement(xmlElement, 'CustomChrBio').text = self.customChrBio
        if self.customChrGoals:
            ET.SubElement(xmlElement, 'CustomChrGoals').text = self.customChrGoals

        if self.wordCountStart:
            ET.SubElement(xmlElement, 'WordCountStart').text = str(self.wordCountStart)
        if self.wordTarget:
            ET.SubElement(xmlElement, 'WordTarget').text = str(self.wordTarget)

        if self.referenceDate:
            ET.SubElement(xmlElement, 'ReferenceDate').text = self.referenceDate

    def update_plot_lines(self):
        for scId in self.sections:
            self.sections[scId].scPlotPoints = {}
            self.sections[scId].scPlotLines = []
            for plId in self.plotLines:
                if scId in self.plotLines[plId].sections:
                    self.sections[scId].scPlotLines.append(plId)
                    for ppId in self.tree.get_children(plId):
                        if self.plotPoints[ppId].sectionAssoc == scId:
                            self.sections[scId].scPlotPoints[ppId] = plId
                            break



class BasicElementNotes(BasicElement):

    def __init__(self,
            notes=None,
            **kwargs):
        super().__init__(**kwargs)
        self._notes = notes

    @property
    def notes(self):
        return self._notes

    @notes.setter
    def notes(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._notes != newVal:
            self._notes = newVal
            self.on_element_change()

    def from_xml(self, xmlElement):
        super().from_xml(xmlElement)
        self.notes = self._xml_element_to_text(xmlElement.find('Notes'))

    def to_xml(self, xmlElement):
        super().to_xml(xmlElement)
        if self.notes:
            xmlElement.append(self._text_to_xml_element('Notes', self.notes))



class Chapter(BasicElementNotes):

    def __init__(self,
            chLevel=None,
            chType=None,
            noNumber=None,
            isTrash=None,
            **kwargs):
        super().__init__(**kwargs)
        self._chLevel = chLevel
        self._chType = chType
        self._noNumber = noNumber
        self._isTrash = isTrash

    @property
    def chLevel(self):
        return self._chLevel

    @chLevel.setter
    def chLevel(self, newVal):
        if newVal is not None:
            assert type(newVal) == int
        if self._chLevel != newVal:
            self._chLevel = newVal
            self.on_element_change()

    @property
    def chType(self):
        return self._chType

    @chType.setter
    def chType(self, newVal):
        if newVal is not None:
            assert type(newVal) == int
        if self._chType != newVal:
            self._chType = newVal
            self.on_element_change()

    @property
    def noNumber(self):
        return self._noNumber

    @noNumber.setter
    def noNumber(self, newVal):
        if newVal is not None:
            assert type(newVal) == bool
        if self._noNumber != newVal:
            self._noNumber = newVal
            self.on_element_change()

    @property
    def isTrash(self):
        return self._isTrash

    @isTrash.setter
    def isTrash(self, newVal):
        if newVal is not None:
            assert type(newVal) == bool
        if self._isTrash != newVal:
            self._isTrash = newVal
            self.on_element_change()

    def from_xml(self, xmlElement):
        super().from_xml(xmlElement)
        typeStr = xmlElement.get('type', '0')
        if typeStr in ('0', '1'):
            self.chType = int(typeStr)
        else:
            self.chType = 1
        chLevel = xmlElement.get('level', None)
        if chLevel == '1':
            self.chLevel = 1
        else:
            self.chLevel = 2
        self.isTrash = xmlElement.get('isTrash', None) == '1'
        self.noNumber = xmlElement.get('noNumber', None) == '1'

    def to_xml(self, xmlElement):
        super().to_xml(xmlElement)
        if self.chType:
            xmlElement.set('type', str(self.chType))
        if self.chLevel == 1:
            xmlElement.set('level', '1')
        if self.isTrash:
            xmlElement.set('isTrash', '1')
        if self.noNumber:
            xmlElement.set('noNumber', '1')


class BasicElementTags(BasicElementNotes):

    def __init__(self,
            tags=None,
            **kwargs):
        super().__init__(**kwargs)
        self._tags = tags

    @property
    def tags(self):
        return self._tags

    @tags.setter
    def tags(self, newVal):
        if newVal is not None:
            for elem in newVal:
                if elem is not None:
                    assert type(elem) == str
        if self._tags != newVal:
            self._tags = newVal
            self.on_element_change()

    def from_xml(self, xmlElement):
        super().from_xml(xmlElement)
        tags = string_to_list(self._get_element_text(xmlElement, 'Tags'))
        strippedTags = []
        for tag in tags:
            strippedTags.append(tag.strip())
        self.tags = strippedTags

    def to_xml(self, xmlElement):
        super().to_xml(xmlElement)
        tagStr = list_to_string(self.tags)
        if tagStr:
            ET.SubElement(xmlElement, 'Tags').text = tagStr



class WorldElement(BasicElementTags):

    def __init__(self,
            aka=None,
            **kwargs):
        super().__init__(**kwargs)
        self._aka = aka

    @property
    def aka(self):
        return self._aka

    @aka.setter
    def aka(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._aka != newVal:
            self._aka = newVal
            self.on_element_change()

    def from_xml(self, xmlElement):
        super().from_xml(xmlElement)
        self.aka = self._get_element_text(xmlElement, 'Aka')

    def to_xml(self, xmlElement):
        super().to_xml(xmlElement)
        if self.aka:
            ET.SubElement(xmlElement, 'Aka').text = self.aka



class Character(WorldElement):
    MAJOR_MARKER = 'Major'
    MINOR_MARKER = 'Minor'

    def __init__(self,
            bio=None,
            goals=None,
            fullName=None,
            isMajor=None,
            birthDate=None,
            deathDate=None,
            **kwargs):
        super().__init__(**kwargs)
        self._bio = bio
        self._goals = goals
        self._fullName = fullName
        self._isMajor = isMajor
        self._birthDate = birthDate
        self._deathDate = deathDate

    @property
    def bio(self):
        return self._bio

    @bio.setter
    def bio(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._bio != newVal:
            self._bio = newVal
            self.on_element_change()

    @property
    def goals(self):
        return self._goals

    @goals.setter
    def goals(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._goals != newVal:
            self._goals = newVal
            self.on_element_change()

    @property
    def fullName(self):
        return self._fullName

    @fullName.setter
    def fullName(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._fullName != newVal:
            self._fullName = newVal
            self.on_element_change()

    @property
    def isMajor(self):
        return self._isMajor

    @isMajor.setter
    def isMajor(self, newVal):
        if newVal is not None:
            assert type(newVal) == bool
        if self._isMajor != newVal:
            self._isMajor = newVal
            self.on_element_change()

    @property
    def birthDate(self):
        return self._birthDate

    @birthDate.setter
    def birthDate(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._birthDate != newVal:
            self._birthDate = newVal
            self.on_element_change()

    @property
    def deathDate(self):
        return self._deathDate

    @deathDate.setter
    def deathDate(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._deathDate != newVal:
            self._deathDate = newVal
            self.on_element_change()

    def from_xml(self, xmlElement):
        super().from_xml(xmlElement)
        self.isMajor = xmlElement.get('major', None) == '1'
        self.fullName = self._get_element_text(xmlElement, 'FullName')
        self.bio = self._xml_element_to_text(xmlElement.find('Bio'))
        self.goals = self._xml_element_to_text(xmlElement.find('Goals'))
        self.birthDate = verified_date(self._get_element_text(xmlElement, 'BirthDate'))
        self.deathDate = verified_date(self._get_element_text(xmlElement, 'DeathDate'))

    def to_xml(self, xmlElement):
        super().to_xml(xmlElement)
        if self.isMajor:
            xmlElement.set('major', '1')
        if self.fullName:
            ET.SubElement(xmlElement, 'FullName').text = self.fullName
        if self.bio:
            xmlElement.append(self._text_to_xml_element('Bio', self.bio))
        if self.goals:
            xmlElement.append(self._text_to_xml_element('Goals', self.goals))
        if self.birthDate:
            ET.SubElement(xmlElement, 'BirthDate').text = self.birthDate
        if self.deathDate:
            ET.SubElement(xmlElement, 'DeathDate').text = self.deathDate



class NvTree:

    def __init__(self):
        self.roots = {
            CH_ROOT:[],
            CR_ROOT:[],
            LC_ROOT:[],
            IT_ROOT:[],
            PL_ROOT:[],
            PN_ROOT:[],
        }
        self.srtSections = {}
        self.srtTurningPoints = {}

    def append(self, parent, iid):
        if parent in self.roots:
            self.roots[parent].append(iid)
            if parent == CH_ROOT:
                self.srtSections[iid] = []
            elif parent == PL_ROOT:
                self.srtTurningPoints[iid] = []
            return

        if parent.startswith(CHAPTER_PREFIX):
            if parent in self.srtSections:
                self.srtSections[parent].append(iid)
            else:
                self.srtSections[parent] = [iid]
            return

        if parent.startswith(PLOT_LINE_PREFIX):
            if parent in self.srtTurningPoints:
                self.srtTurningPoints[parent].append(iid)
            else:
                self.srtTurningPoints[parent] = [iid]

    def delete(self, *items):
        raise NotImplementedError

    def delete_children(self, parent):
        if parent in self.roots:
            self.roots[parent] = []
            if parent == CH_ROOT:
                self.srtSections = {}
                return

            if parent == PL_ROOT:
                self.srtTurningPoints = {}
            return

        if parent.startswith(CHAPTER_PREFIX):
            self.srtSections[parent] = []
            return

        if parent.startswith(PLOT_LINE_PREFIX):
            self.srtTurningPoints[parent] = []

    def get_children(self, item):
        if item in self.roots:
            return self.roots[item]

        if item.startswith(CHAPTER_PREFIX):
            return self.srtSections.get(item, [])

        if item.startswith(PLOT_LINE_PREFIX):
            return self.srtTurningPoints.get(item, [])

    def index(self, item):
        raise NotImplementedError

    def insert(self, parent, index, iid):
        if parent in self.roots:
            self.roots[parent].insert(index, iid)
            if parent == CH_ROOT:
                self.srtSections[iid] = []
            elif parent == PL_ROOT:
                self.srtTurningPoints[iid] = []
            return

        if parent.startswith(CHAPTER_PREFIX):
            if parent in self.srtSections:
                self.srtSections[parent].insert(index, iid)
            else:
                self.srtSections[parent] = [iid]
            return

        if parent.startswith(PLOT_LINE_PREFIX):
            if parent in self.srtTurningPoints:
                self.srtTurningPoints[parent].insert(index, iid)
            else:
                self.srtTurningPoints[parent] = [iid]

    def move(self, item, parent, index):
        raise NotImplementedError

    def next(self, item):
        raise NotImplementedError

    def parent(self, item):
        raise NotImplementedError

    def prev(self, item):
        raise NotImplementedError

    def reset(self):
        for item in self.roots:
            self.roots[item] = []
        self.srtSections = {}
        self.srtTurningPoints = {}

    def set_children(self, item, newchildren):
        if item in self.roots:
            self.roots[item] = newchildren[:]
            if item == CH_ROOT:
                self.srtSections = {}
                return

            if item == PL_ROOT:
                self.srtTurningPoints = {}
            return

        if item.startswith(CHAPTER_PREFIX):
            self.srtSections[item] = newchildren[:]
            return

        if item.startswith(PLOT_LINE_PREFIX):
            self.srtTurningPoints[item] = newchildren[:]



class PlotLine(BasicElementNotes):

    def __init__(self,
            shortName=None,
            sections=None,
            **kwargs):
        super().__init__(**kwargs)

        self._shortName = shortName
        self._sections = sections

    @property
    def shortName(self):
        return self._shortName

    @shortName.setter
    def shortName(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._shortName != newVal:
            self._shortName = newVal
            self.on_element_change()

    @property
    def sections(self):
        try:
            return self._sections[:]
        except TypeError:
            return None

    @sections.setter
    def sections(self, newVal):
        if newVal is not None:
            for elem in newVal:
                if elem is not None:
                    assert type(elem) == str
        if self._sections != newVal:
            self._sections = newVal
            self.on_element_change()

    def from_xml(self, xmlElement):
        super().from_xml(xmlElement)
        self.shortName = self._get_element_text(xmlElement, 'ShortName')
        plSections = []
        xmlSections = xmlElement.find('Sections')
        if xmlSections is not None:
            scIds = xmlSections.get('ids', None)
            if scIds is not None:
                for scId in string_to_list(scIds, divider=' '):
                    plSections.append(scId)
        self.sections = plSections

    def to_xml(self, xmlElement):
        super().to_xml(xmlElement)
        if self.shortName:
            ET.SubElement(xmlElement, 'ShortName').text = self.shortName
        if self.sections:
            attrib = {'ids':' '.join(self.sections)}
            ET.SubElement(xmlElement, 'Sections', attrib=attrib)


class PlotPoint(BasicElementNotes):

    def __init__(self,
            sectionAssoc=None,
            **kwargs):
        super().__init__(**kwargs)

        self._sectionAssoc = sectionAssoc

    @property
    def sectionAssoc(self):
        return self._sectionAssoc

    @sectionAssoc.setter
    def sectionAssoc(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._sectionAssoc != newVal:
            self._sectionAssoc = newVal
            self.on_element_change()

    def from_xml(self, xmlElement):
        super().from_xml(xmlElement)
        xmlSectionAssoc = xmlElement.find('Section')
        if xmlSectionAssoc is not None:
            self.sectionAssoc = xmlSectionAssoc.get('id', None)

    def to_xml(self, xmlElement):
        super().to_xml(xmlElement)
        if self.sectionAssoc:
            ET.SubElement(xmlElement, 'Section', attrib={'id': self.sectionAssoc})
from datetime import date
from datetime import datetime
from datetime import time
from datetime import timedelta

from calendar import isleap
from datetime import date
from datetime import datetime
from datetime import timedelta


def difference_in_years(startDate, endDate):
    diffyears = endDate.year - startDate.year
    difference = endDate - startDate.replace(endDate.year)
    days_in_year = isleap(endDate.year) and 366 or 365
    years = diffyears + (difference.days + difference.seconds / 86400.0) / days_in_year
    return int(years)


def get_age(nowIso, birthDateIso, deathDateIso):
    now = datetime.fromisoformat(nowIso)
    if deathDateIso:
        deathDate = datetime.fromisoformat(deathDateIso)
        if now > deathDate:
            years = difference_in_years(deathDate, now)
            return -1 * years

    birthDate = datetime.fromisoformat(birthDateIso)
    years = difference_in_years(birthDate, now)
    return years


def get_specific_date(dayStr, refIso):
    refDate = date.fromisoformat(refIso)
    return date.isoformat(refDate + timedelta(days=int(dayStr)))


def get_unspecific_date(dateIso, refIso):
    refDate = date.fromisoformat(refIso)
    return str((date.fromisoformat(dateIso) - refDate).days)


ADDITIONAL_WORD_LIMITS = re.compile(r'--|—|–|\<\/p\>')

NO_WORD_LIMITS = re.compile(r'\<note\>.*?\<\/note\>|\<comment\>.*?\<\/comment\>|\<.+?\>')


class Section(BasicElementTags):

    SCENE = ['-', 'A', 'R', 'x']

    STATUS = [
        None,
        _('Outline'),
        _('Draft'),
        _('1st Edit'),
        _('2nd Edit'),
        _('Done')
    ]

    NULL_DATE = '0001-01-01'
    NULL_TIME = '00:00:00'

    def __init__(self,
            scType=None,
            scene=None,
            status=None,
            appendToPrev=None,
            goal=None,
            conflict=None,
            outcome=None,
            plotNotes=None,
            scDate=None,
            scTime=None,
            day=None,
            lastsMinutes=None,
            lastsHours=None,
            lastsDays=None,
            characters=None,
            locations=None,
            items=None,
            **kwargs):
        super().__init__(**kwargs)
        self._sectionContent = None
        self.wordCount = 0

        self._scType = scType
        self._scene = scene
        self._status = status
        self._appendToPrev = appendToPrev
        self._goal = goal
        self._conflict = conflict
        self._outcome = outcome
        self._plotlineNotes = plotNotes
        try:
            newDate = date.fromisoformat(scDate)
            self._weekDay = newDate.weekday()
            self._localeDate = newDate.strftime('%x')
            self._date = scDate
        except:
            self._weekDay = None
            self._localeDate = None
            self._date = None
        self._time = scTime
        self._day = day
        self._lastsMinutes = lastsMinutes
        self._lastsHours = lastsHours
        self._lastsDays = lastsDays
        self._characters = characters
        self._locations = locations
        self._items = items

        self.scPlotLines = []
        self.scPlotPoints = {}

    @property
    def sectionContent(self):
        return self._sectionContent

    @sectionContent.setter
    def sectionContent(self, text):
        if text is not None:
            assert type(text) == str
        if self._sectionContent != text:
            self._sectionContent = text
            if text is not None:
                text = ADDITIONAL_WORD_LIMITS.sub(' ', text)
                text = NO_WORD_LIMITS.sub('', text)
                wordList = text.split()
                self.wordCount = len(wordList)
            else:
                self.wordCount = 0
            self.on_element_change()

    @property
    def scType(self):
        return self._scType

    @scType.setter
    def scType(self, newVal):
        if newVal is not None:
            assert type(newVal) == int
        if self._scType != newVal:
            self._scType = newVal
            self.on_element_change()

    @property
    def scene(self):
        return self._scene

    @scene.setter
    def scene(self, newVal):
        if newVal is not None:
            assert type(newVal) == int
        if self._scene != newVal:
            self._scene = newVal
            self.on_element_change()

    @property
    def status(self):
        return self._status

    @status.setter
    def status(self, newVal):
        if newVal is not None:
            assert type(newVal) == int
        if self._status != newVal:
            self._status = newVal
            self.on_element_change()

    @property
    def appendToPrev(self):
        return self._appendToPrev

    @appendToPrev.setter
    def appendToPrev(self, newVal):
        if newVal is not None:
            assert type(newVal) == bool
        if self._appendToPrev != newVal:
            self._appendToPrev = newVal
            self.on_element_change()

    @property
    def goal(self):
        return self._goal

    @goal.setter
    def goal(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._goal != newVal:
            self._goal = newVal
            self.on_element_change()

    @property
    def conflict(self):
        return self._conflict

    @conflict.setter
    def conflict(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._conflict != newVal:
            self._conflict = newVal
            self.on_element_change()

    @property
    def outcome(self):
        return self._outcome

    @outcome.setter
    def outcome(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._outcome != newVal:
            self._outcome = newVal
            self.on_element_change()

    @property
    def plotlineNotes(self):
        try:
            return dict(self._plotlineNotes)
        except TypeError:
            return None

    @plotlineNotes.setter
    def plotlineNotes(self, newVal):
        if newVal is not None:
            for elem in newVal:
                val = newVal[elem]
                if val is not None:
                    assert type(val) == str
        if self._plotlineNotes != newVal:
            self._plotlineNotes = newVal
            self.on_element_change()

    @property
    def date(self):
        return self._date

    @date.setter
    def date(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._date != newVal:
            if not newVal:
                self._date = None
                self._weekDay = None
                self._localeDate = None
                self.on_element_change()
                return

            try:
                newDate = date.fromisoformat(newVal)
                self._weekDay = newDate.weekday()
            except:
                return

            try:
                self._localeDate = newDate.strftime('%x')
            except:
                self._localeDate = newVal
            self._date = newVal
            self.on_element_change()

    @property
    def weekDay(self):
        return self._weekDay

    @property
    def localeDate(self):
        return self._localeDate

    @property
    def time(self):
        return self._time

    @time.setter
    def time(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._time != newVal:
            self._time = newVal
            self.on_element_change()

    @property
    def day(self):
        return self._day

    @day.setter
    def day(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._day != newVal:
            self._day = newVal
            self.on_element_change()

    @property
    def lastsMinutes(self):
        return self._lastsMinutes

    @lastsMinutes.setter
    def lastsMinutes(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._lastsMinutes != newVal:
            self._lastsMinutes = newVal
            self.on_element_change()

    @property
    def lastsHours(self):
        return self._lastsHours

    @lastsHours.setter
    def lastsHours(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._lastsHours != newVal:
            self._lastsHours = newVal
            self.on_element_change()

    @property
    def lastsDays(self):
        return self._lastsDays

    @lastsDays.setter
    def lastsDays(self, newVal):
        if newVal is not None:
            assert type(newVal) == str
        if self._lastsDays != newVal:
            self._lastsDays = newVal
            self.on_element_change()

    @property
    def characters(self):
        try:
            return self._characters[:]
        except TypeError:
            return None

    @characters.setter
    def characters(self, newVal):
        if newVal is not None:
            for elem in newVal:
                if elem is not None:
                    assert type(elem) == str
        if self._characters != newVal:
            self._characters = newVal
            self.on_element_change()

    @property
    def locations(self):
        try:
            return self._locations[:]
        except TypeError:
            return None

    @locations.setter
    def locations(self, newVal):
        if newVal is not None:
            for elem in newVal:
                if elem is not None:
                    assert type(elem) == str
        if self._locations != newVal:
            self._locations = newVal
            self.on_element_change()

    @property
    def items(self):
        try:
            return self._items[:]
        except TypeError:
            return None

    @items.setter
    def items(self, newVal):
        if newVal is not None:
            for elem in newVal:
                if elem is not None:
                    assert type(elem) == str
        if self._items != newVal:
            self._items = newVal
            self.on_element_change()

    def day_to_date(self, referenceDate):
        if self._date:
            return True

        try:
            self.date = get_specific_date(self._day, referenceDate)
            self._day = None
            return True

        except:
            self.date = None
            return False

    def date_to_day(self, referenceDate):
        if self._day:
            return True

        try:
            self._day = get_unspecific_date(self._date, referenceDate)
            self.date = None
            return True

        except:
            self._day = None
            return False

    def from_xml(self, xmlElement):
        super().from_xml(xmlElement)
        typeStr = xmlElement.get('type', '0')
        if typeStr in ('0', '1', '2', '3'):
            self.scType = int(typeStr)
        else:
            self.scType = 1
        status = xmlElement.get('status', None)
        if status in ('2', '3', '4', '5'):
            self.status = int(status)
        else:
            self.status = 1
        scene = xmlElement.get('scene', 0)
        if scene in ('1', '2', '3'):
            self.scene = int(scene)
        else:
            self.scene = 0

        if not self.scene:
            sceneKind = xmlElement.get('pacing', None)
            if sceneKind in ('1', '2'):
                self.scene = int(sceneKind) + 1

        self.appendToPrev = xmlElement.get('append', None) == '1'

        self.goal = self._xml_element_to_text(xmlElement.find('Goal'))
        self.conflict = self._xml_element_to_text(xmlElement.find('Conflict'))
        self.outcome = self._xml_element_to_text(xmlElement.find('Outcome'))

        xmlPlotNotes = xmlElement.find('PlotNotes')
        if xmlPlotNotes is None:
            xmlPlotNotes = xmlElement
        plotNotes = {}
        for xmlPlotLineNote in xmlPlotNotes.iterfind('PlotlineNotes'):
            plId = xmlPlotLineNote.get('id', None)
            plotNotes[plId] = self._xml_element_to_text(xmlPlotLineNote)
        self.plotlineNotes = plotNotes

        if xmlElement.find('Date') is not None:
            self.date = verified_date(xmlElement.find('Date').text)
        elif xmlElement.find('Day') is not None:
            self.day = verified_int_string(xmlElement.find('Day').text)

        if xmlElement.find('Time') is not None:
            self.time = verified_time(xmlElement.find('Time').text)

        self.lastsDays = verified_int_string(self._get_element_text(xmlElement, 'LastsDays'))
        self.lastsHours = verified_int_string(self._get_element_text(xmlElement, 'LastsHours'))
        self.lastsMinutes = verified_int_string(self._get_element_text(xmlElement, 'LastsMinutes'))

        scCharacters = []
        xmlCharacters = xmlElement.find('Characters')
        if xmlCharacters is not None:
            crIds = xmlCharacters.get('ids', None)
            if crIds is not None:
                for crId in string_to_list(crIds, divider=' '):
                    scCharacters.append(crId)
        self.characters = scCharacters

        scLocations = []
        xmlLocations = xmlElement.find('Locations')
        if xmlLocations is not None:
            lcIds = xmlLocations.get('ids', None)
            if lcIds is not None:
                for lcId in string_to_list(lcIds, divider=' '):
                    scLocations.append(lcId)
        self.locations = scLocations

        scItems = []
        xmlItems = xmlElement.find('Items')
        if xmlItems is not None:
            itIds = xmlItems.get('ids', None)
            if itIds is not None:
                for itId in string_to_list(itIds, divider=' '):
                    scItems.append(itId)
        self.items = scItems

        xmlContent = xmlElement.find('Content')
        if xmlContent is not None:
            xmlStr = ET.tostring(
                xmlContent,
                encoding='utf-8',
                short_empty_elements=False
                ).decode('utf-8')
            xmlStr = xmlStr.replace('<Content>', '').replace('</Content>', '')

            lines = xmlStr.split('\n')
            newlines = []
            for line in lines:
                newlines.append(line.strip())
            xmlStr = ''.join(newlines)
            if xmlStr:
                self.sectionContent = xmlStr
            else:
                self.sectionContent = '<p></p>'
        else:
            self.sectionContent = '<p></p>'

    def get_end_date_time(self):
        endDate = None
        endTime = None
        endDay = None
        if self.lastsDays:
            lastsDays = int(self.lastsDays)
        else:
            lastsDays = 0
        if self.lastsHours:
            lastsSeconds = int(self.lastsHours) * 3600
        else:
            lastsSeconds = 0
        if self.lastsMinutes:
            lastsSeconds += int(self.lastsMinutes) * 60
        sectionDuration = timedelta(days=lastsDays, seconds=lastsSeconds)
        if self.time:
            if self.date:
                try:
                    sectionStart = datetime.fromisoformat(f'{self.date} {self.time}')
                    sectionEnd = sectionStart + sectionDuration
                    endDate, endTime = sectionEnd.isoformat().split('T')
                except:
                    pass
            else:
                try:
                    if self.day:
                        dayInt = int(self.day)
                    else:
                        dayInt = 0
                    startDate = (date.min + timedelta(days=dayInt)).isoformat()
                    sectionStart = datetime.fromisoformat(f'{startDate} {self.time}')
                    sectionEnd = sectionStart + sectionDuration
                    endDate, endTime = sectionEnd.isoformat().split('T')
                    endDay = str((date.fromisoformat(endDate) - date.min).days)
                    endDate = None
                except:
                    pass
        return endDate, endTime, endDay

    def to_xml(self, xmlElement):
        super().to_xml(xmlElement)
        if self.scType:
            xmlElement.set('type', str(self.scType))
        if self.status > 1:
            xmlElement.set('status', str(self.status))
        if self.scene > 0:
            xmlElement.set('scene', str(self.scene))
        if self.appendToPrev:
            xmlElement.set('append', '1')

        if self.goal:
            xmlElement.append(self._text_to_xml_element('Goal', self.goal))
        if self.conflict:
            xmlElement.append(self._text_to_xml_element('Conflict', self.conflict))
        if self.outcome:
            xmlElement.append(self._text_to_xml_element('Outcome', self.outcome))

        if self.plotlineNotes:
            for plId in self.plotlineNotes:
                if not plId in self.scPlotLines:
                    continue

                if not self.plotlineNotes[plId]:
                    continue

                xmlPlotlineNotes = self._text_to_xml_element('PlotlineNotes', self.plotlineNotes[plId])
                xmlPlotlineNotes.set('id', plId)
                xmlElement.append(xmlPlotlineNotes)

        if self.date:
            ET.SubElement(xmlElement, 'Date').text = self.date
        elif self.day:
            ET.SubElement(xmlElement, 'Day').text = self.day
        if self.time:
            ET.SubElement(xmlElement, 'Time').text = self.time

        if self.lastsDays and self.lastsDays != '0':
            ET.SubElement(xmlElement, 'LastsDays').text = self.lastsDays
        if self.lastsHours and self.lastsHours != '0':
            ET.SubElement(xmlElement, 'LastsHours').text = self.lastsHours
        if self.lastsMinutes and self.lastsMinutes != '0':
            ET.SubElement(xmlElement, 'LastsMinutes').text = self.lastsMinutes

        if self.characters:
            attrib = {'ids':' '.join(self.characters)}
            ET.SubElement(xmlElement, 'Characters', attrib=attrib)

        if self.locations:
            attrib = {'ids':' '.join(self.locations)}
            ET.SubElement(xmlElement, 'Locations', attrib=attrib)

        if self.items:
            attrib = {'ids':' '.join(self.items)}
            ET.SubElement(xmlElement, 'Items', attrib=attrib)

        sectionContent = self.sectionContent
        if sectionContent:
            if not sectionContent in ('<p></p>', '<p />'):
                xmlElement.append(ET.fromstring(f'<Content>{sectionContent}</Content>'))
from datetime import date



def get_xml_root(filePath):
    try:
        xmlTree = ET.parse(filePath)
    except Exception as ex:
        raise Error(f'{_("Cannot process file")}: "{norm_path(filePath)}" - {str(ex)}')

    return xmlTree.getroot()


class NovxFile(File):
    DESCRIPTION = _('novelibre project')
    EXTENSION = '.novx'

    MAJOR_VERSION = 1
    MINOR_VERSION = 4

    XML_HEADER = f'''<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE novx SYSTEM "novx_{MAJOR_VERSION}_{MINOR_VERSION}.dtd">
<?xml-stylesheet href="novx.css" type="text/css"?>
'''

    def __init__(self, filePath, **kwargs):
        super().__init__(filePath)
        self.on_element_change = None
        self.xmlTree = None

        self.wcLog = {}

        self.wcLogUpdate = {}

        self.timestamp = None

    def adjust_section_types(self):
        partType = 0
        for chId in self.novel.tree.get_children(CH_ROOT):
            if self.novel.chapters[chId].chLevel == 1:
                partType = self.novel.chapters[chId].chType
            elif partType != 0 and not self.novel.chapters[chId].isTrash:
                self.novel.chapters[chId].chType = partType
            for scId in self.novel.tree.get_children(chId):
                if self.novel.sections[scId].scType < self.novel.chapters[chId].chType:
                    self.novel.sections[scId].scType = self.novel.chapters[chId].chType

    def count_words(self):
        count = 0
        totalCount = 0
        for chId in self.novel.tree.get_children(CH_ROOT):
            if not self.novel.chapters[chId].isTrash:
                for scId in self.novel.tree.get_children(chId):
                    if self.novel.sections[scId].scType < 2:
                        totalCount += self.novel.sections[scId].wordCount
                        if self.novel.sections[scId].scType == 0:
                            count += self.novel.sections[scId].wordCount
        return count, totalCount

    def read(self):
        xmlRoot = get_xml_root(self.filePath)
        self._check_version(xmlRoot)
        try:
            locale = xmlRoot.attrib['{http://www.w3.org/XML/1998/namespace}lang']
            self.novel.languageCode, self.novel.countryCode = locale.split('-')
        except:
            pass
        self.novel.tree.reset()
        try:
            self._read_project(xmlRoot)
            self._read_locations(xmlRoot)
            self._read_items(xmlRoot)
            self._read_characters(xmlRoot)
            self._read_chapters_and_sections(xmlRoot)
            self._read_plot_lines_and_points(xmlRoot)
            self._read_project_notes(xmlRoot)
            self.adjust_section_types()
            self._read_word_count_log(xmlRoot)
        except Exception as ex:
            raise Error(f"{_('Corrupt project data')} ({str(ex)})")
        self._get_timestamp()
        self._keep_word_count()

    def write(self):
        self._update_word_count_log()
        self.adjust_section_types()
        self.novel.get_languages()

        attrib = {
            'version':f'{self.MAJOR_VERSION}.{self.MINOR_VERSION}',
            'xml:lang':f'{self.novel.languageCode}-{self.novel.countryCode}',
        }
        xmlRoot = ET.Element('novx', attrib=attrib)
        self._build_project(xmlRoot)
        self._build_chapters_and_sections(xmlRoot)
        self._build_characters(xmlRoot)
        self._build_locations(xmlRoot)
        self._build_items(xmlRoot)
        self._build_plot_lines_and_points(xmlRoot)
        self._build_project_notes(xmlRoot)
        self._build_word_count_log(xmlRoot)

        indent(xmlRoot)

        self.xmlTree = ET.ElementTree(xmlRoot)
        self._write_element_tree(self)
        self._postprocess_xml_file(self.filePath)
        self._get_timestamp()

    def _build_project(self, root):
        xmlProject = ET.SubElement(root, 'PROJECT')
        self.novel.to_xml(xmlProject)

    def _build_chapters_and_sections(self, root):
        xmlChapters = ET.SubElement(root, 'CHAPTERS')
        for chId in self.novel.tree.get_children(CH_ROOT):
            xmlChapter = ET.SubElement(xmlChapters, 'CHAPTER', attrib={'id':chId})
            self.novel.chapters[chId].to_xml(xmlChapter)
            for scId in self.novel.tree.get_children(chId):
                self.novel.sections[scId].to_xml(ET.SubElement(xmlChapter, 'SECTION', attrib={'id':scId}))

    def _build_characters(self, root):
        xmlCharacters = ET.SubElement(root, 'CHARACTERS')
        for crId in self.novel.tree.get_children(CR_ROOT):
            self.novel.characters[crId].to_xml(ET.SubElement(xmlCharacters, 'CHARACTER', attrib={'id':crId}))

    def _build_locations(self, root):
        xmlLocations = ET.SubElement(root, 'LOCATIONS')
        for lcId in self.novel.tree.get_children(LC_ROOT):
            self.novel.locations[lcId].to_xml(ET.SubElement(xmlLocations, 'LOCATION', attrib={'id':lcId}))

    def _build_items(self, root):
        xmlItems = ET.SubElement(root, 'ITEMS')
        for itId in self.novel.tree.get_children(IT_ROOT):
            self.novel.items[itId].to_xml(ET.SubElement(xmlItems, 'ITEM', attrib={'id':itId}))

    def _build_plot_lines_and_points(self, root):
        xmlPlotLines = ET.SubElement(root, 'ARCS')
        for plId in self.novel.tree.get_children(PL_ROOT):
            xmlPlotLine = ET.SubElement(xmlPlotLines, 'ARC', attrib={'id':plId})
            self.novel.plotLines[plId].to_xml(xmlPlotLine)
            for ppId in self.novel.tree.get_children(plId):
                self.novel.plotPoints[ppId].to_xml(ET.SubElement(xmlPlotLine, 'POINT', attrib={'id':ppId}))

    def _build_project_notes(self, root):
        xmlProjectNotes = ET.SubElement(root, 'PROJECTNOTES')
        for pnId in self.novel.tree.get_children(PN_ROOT):
            self.novel.projectNotes[pnId].to_xml(ET.SubElement(xmlProjectNotes, 'PROJECTNOTE', attrib={'id':pnId}))

    def _build_word_count_log(self, root):
        if not self.wcLog:
            return

        xmlWcLog = ET.SubElement(root, 'PROGRESS')
        wcLastCount = None
        wcLastTotalCount = None
        for wc in self.wcLog:
            if self.novel.saveWordCount:
                if self.wcLog[wc][0] == wcLastCount and self.wcLog[wc][1] == wcLastTotalCount:
                    continue

                wcLastCount = self.wcLog[wc][0]
                wcLastTotalCount = self.wcLog[wc][1]
            xmlWc = ET.SubElement(xmlWcLog, 'WC')
            ET.SubElement(xmlWc, 'Date').text = wc
            ET.SubElement(xmlWc, 'Count').text = self.wcLog[wc][0]
            ET.SubElement(xmlWc, 'WithUnused').text = self.wcLog[wc][1]

    def _check_id(self, elemId, elemPrefix):
        if not elemId.startswith(elemPrefix):
            raise Error(f"bad ID: '{elemId}'")

    def _check_version(self, xmlRoot):
        if xmlRoot.tag != 'novx':
            raise Error(f'{_("No valid xml root element found in file")}: "{norm_path(self.filePath)}".')
        try:
            majorVersionStr, minorVersionStr = xmlRoot.attrib['version'].split('.')
            majorVersion = int(majorVersionStr)
            minorVersion = int(minorVersionStr)
        except:
            raise Error(f'{_("No valid version found in file")}: "{norm_path(self.filePath)}".')
        if majorVersion > self.MAJOR_VERSION:
            raise Error(_('The project "{}" was created with a newer novelibre version.').format(norm_path(self.filePath)))
        elif majorVersion < self.MAJOR_VERSION:
            raise Error(_('The project "{}" was created with an outdated novelibre version.').format(norm_path(self.filePath)))
        elif minorVersion > self.MINOR_VERSION:
            raise Error(_('The project "{}" was created with a newer novelibre version.').format(norm_path(self.filePath)))

    def _get_timestamp(self):
        try:
            self.timestamp = os.path.getmtime(self.filePath)
        except:
            self.timestamp = None

    def _keep_word_count(self):
        if not self.wcLog:
            return

        actualCountInt, actualTotalCountInt = self.count_words()
        actualCount = str(actualCountInt)
        actualTotalCount = str(actualTotalCountInt)
        latestDate = list(self.wcLog)[-1]
        latestCount = self.wcLog[latestDate][0]
        latestTotalCount = self.wcLog[latestDate][1]
        if actualCount != latestCount or actualTotalCount != latestTotalCount:
            try:
                fileDateIso = date.fromtimestamp(self.timestamp).isoformat()
            except:
                fileDateIso = date.today().isoformat()
            self.wcLogUpdate[fileDateIso] = [actualCount, actualTotalCount]

    def _postprocess_xml_file(self, filePath):
        with open(filePath, 'r', encoding='utf-8') as f:
            text = f.read()
            text = strip_illegal_characters(text)
        try:
            with open(filePath, 'w', encoding='utf-8') as f:
                f.write(f'{self.XML_HEADER}{text}')
        except:
            raise Error(f'{_("Cannot write file")}: "{norm_path(filePath)}".')

    def _read_chapters_and_sections(self, root):
        xmlChapters = root.find('CHAPTERS')
        if xmlChapters is None:
            return

        for xmlChapter in xmlChapters.iterfind('CHAPTER'):
            chId = xmlChapter.attrib['id']
            self._check_id(chId, CHAPTER_PREFIX)
            self.novel.chapters[chId] = Chapter(on_element_change=self.on_element_change)
            self.novel.chapters[chId].from_xml(xmlChapter)
            self.novel.tree.append(CH_ROOT, chId)

            for xmlSection in xmlChapter.iterfind('SECTION'):
                scId = xmlSection.attrib['id']
                self._check_id(scId, SECTION_PREFIX)
                self._read_section(xmlSection, scId)
                self.novel.tree.append(chId, scId)

    def _read_characters(self, root):
        xmlCharacters = root.find('CHARACTERS')
        if xmlCharacters is None:
            return

        for xmlCharacter in xmlCharacters.iterfind('CHARACTER'):
            crId = xmlCharacter.attrib['id']
            self._check_id(crId, CHARACTER_PREFIX)
            self.novel.characters[crId] = Character(on_element_change=self.on_element_change)
            self.novel.characters[crId].from_xml(xmlCharacter)
            self.novel.tree.append(CR_ROOT, crId)

    def _read_items(self, root):
        xmlItems = root.find('ITEMS')
        if xmlItems is None:
            return

        for xmlItem in xmlItems.iterfind('ITEM'):
            itId = xmlItem.attrib['id']
            self._check_id(itId, ITEM_PREFIX)
            self.novel.items[itId] = WorldElement(on_element_change=self.on_element_change)
            self.novel.items[itId].from_xml(xmlItem)
            self.novel.tree.append(IT_ROOT, itId)

    def _read_locations(self, root):
        xmlLocations = root.find('LOCATIONS')
        if xmlLocations is None:
            return

        for xmlLocation in xmlLocations.iterfind('LOCATION'):
            lcId = xmlLocation.attrib['id']
            self._check_id(lcId, LOCATION_PREFIX)
            self.novel.locations[lcId] = WorldElement(on_element_change=self.on_element_change)
            self.novel.locations[lcId].from_xml(xmlLocation)
            self.novel.tree.append(LC_ROOT, lcId)

    def _read_plot_lines_and_points(self, root):
        xmlPlotLines = root.find('ARCS')
        if xmlPlotLines is None:
            return

        for xmlPlotLine in xmlPlotLines.iterfind('ARC'):
            plId = xmlPlotLine.attrib['id']
            self._check_id(plId, PLOT_LINE_PREFIX)
            self.novel.plotLines[plId] = PlotLine(on_element_change=self.on_element_change)
            self.novel.plotLines[plId].from_xml(xmlPlotLine)
            self.novel.tree.append(PL_ROOT, plId)

            self.novel.plotLines[plId].sections = intersection(self.novel.plotLines[plId].sections, self.novel.sections)

            for scId in self.novel.plotLines[plId].sections:
                self.novel.sections[scId].scPlotLines.append(plId)

            for xmlPlotPoint in xmlPlotLine.iterfind('POINT'):
                ppId = xmlPlotPoint.attrib['id']
                self._check_id(ppId, PLOT_POINT_PREFIX)
                self._read_plot_point(xmlPlotPoint, ppId, plId)
                self.novel.tree.append(plId, ppId)

    def _read_plot_point(self, xmlPlotPoint, ppId, plId):
        self.novel.plotPoints[ppId] = PlotPoint(on_element_change=self.on_element_change)
        self.novel.plotPoints[ppId].from_xml(xmlPlotPoint)

        scId = self.novel.plotPoints[ppId].sectionAssoc
        if scId in self.novel.sections:
            self.novel.sections[scId].scPlotPoints[ppId] = plId
        else:
            self.novel.plotPoints[ppId].sectionAssoc = None

    def _read_project(self, root):
        xmlProject = root.find('PROJECT')
        if xmlProject is None:
            return

        self.novel.from_xml(xmlProject)

    def _read_project_notes(self, root):
        xmlProjectNotes = root.find('PROJECTNOTES')
        if xmlProjectNotes is None:
            return

        for xmlProjectNote in xmlProjectNotes.iterfind('PROJECTNOTE'):
            pnId = xmlProjectNote.attrib['id']
            self._check_id(pnId, PRJ_NOTE_PREFIX)
            self.novel.projectNotes[pnId] = BasicElement()
            self.novel.projectNotes[pnId].from_xml(xmlProjectNote)
            self.novel.tree.append(PN_ROOT, pnId)

    def _read_section(self, xmlSection, scId):
        self.novel.sections[scId] = Section(on_element_change=self.on_element_change)
        self.novel.sections[scId].from_xml(xmlSection)

        self.novel.sections[scId].characters = intersection(self.novel.sections[scId].characters, self.novel.characters)
        self.novel.sections[scId].locations = intersection(self.novel.sections[scId].locations, self.novel.locations)
        self.novel.sections[scId].items = intersection(self.novel.sections[scId].items, self.novel.items)

    def _read_word_count_log(self, xmlRoot):
        xmlWclog = xmlRoot.find('PROGRESS')
        if xmlWclog is None:
            return

        for xmlWc in xmlWclog.iterfind('WC'):
            wcDate = verified_date(xmlWc.find('Date').text)
            wcCount = verified_int_string(xmlWc.find('Count').text)
            wcTotalCount = verified_int_string(xmlWc.find('WithUnused').text)
            if wcDate and wcCount and wcTotalCount:
                self.wcLog[wcDate] = [wcCount, wcTotalCount]

    def _update_word_count_log(self):
        if self.novel.saveWordCount:
            newCountInt, newTotalCountInt = self.count_words()
            newCount = str(newCountInt)
            newTotalCount = str(newTotalCountInt)
            todayIso = date.today().isoformat()
            self.wcLogUpdate[todayIso] = [newCount, newTotalCount]
            for wcDate in self.wcLogUpdate:
                self.wcLog[wcDate] = self.wcLogUpdate[wcDate]
        self.wcLogUpdate = {}

    def _write_element_tree(self, xmlProject):
        backedUp = False
        if os.path.isfile(xmlProject.filePath):
            try:
                os.replace(xmlProject.filePath, f'{xmlProject.filePath}.bak')
            except:
                raise Error(f'{_("Cannot overwrite file")}: "{norm_path(xmlProject.filePath)}".')
            else:
                backedUp = True
        try:
            xmlProject.xmlTree.write(xmlProject.filePath, xml_declaration=False, encoding='utf-8')
        except:
            if backedUp:
                os.replace(f'{xmlProject.filePath}.bak', xmlProject.filePath)
            raise Error(f'{_("Cannot write file")}: "{norm_path(xmlProject.filePath)}".')



class NovxService:

    def get_novx_file_extension(self):
        return NovxFile.EXTENSION

    def make_basic_element(self, **kwargs):
        return BasicElement(**kwargs)

    def make_chapter(self, **kwargs):
        return Chapter(**kwargs)

    def make_character(self, **kwargs):
        return Character(**kwargs)

    def make_novel(self, **kwargs):
        kwargs['tree'] = kwargs.get('tree', NvTree())
        return Novel(**kwargs)

    def make_nv_tree(self, **kwargs):
        return NvTree(**kwargs)

    def make_plot_line(self, **kwargs):
        return PlotLine(**kwargs)

    def make_plot_point(self, **kwargs):
        return PlotPoint(**kwargs)

    def make_section(self, **kwargs):
        return Section(**kwargs)

    def make_world_element(self, **kwargs):
        return WorldElement(**kwargs)

    def make_novx_file(self, filePath, **kwargs):
        return NovxFile(filePath, **kwargs)

from tkinter import ttk



class NvTreeview(ttk.Treeview):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.on_element_change = self.do_nothing

        self.append('', CH_ROOT)
        self.append('', CR_ROOT)
        self.append('', LC_ROOT)
        self.append('', IT_ROOT)
        self.append('', PL_ROOT)
        self.append('', PN_ROOT)

    def append(self, parent, iid, text=None):
        if text is None:
            text = iid
        self.insert(parent, 'end', iid, text=text)

    def delete(self, *items):
        super().delete(*items)
        self.on_element_change()

    def delete_children(self, parent):
        for child in self.get_children(parent):
            self.delete(child)

    def insert(self, parent, index, iid=None, **kw):
        super().insert(parent, index, iid, **kw)
        self.on_element_change()

    def move(self, item, parent, index):
        super().move(item, parent, index)
        self.on_element_change()

    def reset(self):
        self.on_element_change = self.do_nothing
        for rootElement in self.get_children(''):
            for child in self.get_children(rootElement):
                self.delete(child)

    def do_nothing(self):
        pass


class NvService(NovxService):

    def get_moon_phase_str(self, isoDate):
        return get_moon_phase_string(isoDate)

    def make_configuration(self, **kwargs):
        return Configuration(**kwargs)

    def make_novel(self, **kwargs):
        kwargs['tree'] = kwargs.get('tree', NvTreeview())
        return Novel(**kwargs)

    def make_nv_tree(self, **kwargs):
        return NvTreeview(**kwargs)


def yw_novx(sourcePath):
    path, extension = os.path.splitext(sourcePath)
    if extension != '.yw7':
        raise ValueError(f'File must be .yw7 type, but is "{extension}".')

    nvService = NvService()
    targetPath = f'{path}.novx'
    source = Yw7File(sourcePath, nv_service=nvService)
    target = nvService.make_novx_file(targetPath)
    source.novel = nvService.make_novel()
    source.read()
    target.novel = source.novel
    target.wcLog = source.wcLog
    target.write()


if __name__ == '__main__':
    yw_novx(sys.argv[1])
    print('Done')

