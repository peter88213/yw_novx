#!/usr/bin/python3
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



def indent(elem, level=0):
    i = f'\n{level * "  "}'
    if elem:
        if not elem.text or not elem.text.strip():
            elem.text = f'{i}  '
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i
#!/usr/bin/python3
SUFFIX = ''


from nvywlib.yw7_file import Yw7File
from datetime import date
from datetime import time

from abc import ABC
from urllib.parse import quote

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
ITEM_REPORT_SUFFIX = '_item_report'
ITEMLIST_SUFFIX = '_itemlist_tmp'
ITEMS_SUFFIX = '_items_tmp'
LOCATION_REPORT_SUFFIX = '_location_report'
LOCATIONS_SUFFIX = '_locations_tmp'
LOCLIST_SUFFIX = '_loclist_tmp'
MANUSCRIPT_SUFFIX = '_manuscript_tmp'
PARTS_SUFFIX = '_parts_tmp'
PLOTLIST_SUFFIX = '_plotlist'
PLOT_SUFFIX = '_plot'
PROJECTNOTES_SUFFIX = '_projectnote_report'
PROOF_SUFFIX = '_proof_tmp'
SECTIONLIST_SUFFIX = '_sectionlist'
GRID_SUFFIX = '_grid_tmp'
SECTIONS_SUFFIX = '_sections_tmp'
XREF_SUFFIX = '_xref'


class Error(Exception):
    pass


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

WEEKDAYS = [
    _('Monday'),
    _('Tuesday'),
    _('Wednesday'),
    _('Thursday'),
    _('Friday'),
    _('Saturday'),
    _('Sunday')
    ]


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



class File(ABC):
    DESCRIPTION = _('File')
    EXTENSION = None
    SUFFIX = None

    PRJ_KWVAR = [
        'Field_ChapterHeadingPrefix',
        'Field_ChapterHeadingSuffix',
        'Field_PartHeadingPrefix',
        'Field_PartHeadingSuffix',
        'Field_CustomGoal',
        'Field_CustomConflict',
        'Field_CustomOutcome',
        'Field_CustomChrBio',
        'Field_CustomChrGoals',
        ]

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



class BasicElement:

    def __init__(self,
            on_element_change=None,
            title=None,
            desc=None):
        if on_element_change is None:
            self.on_element_change = self.do_nothing
        else:
            self.on_element_change = on_element_change
        self._title = title
        self._desc = desc

    @property
    def title(self):
        return self._title

    @title.setter
    def title(self, newVal):
        if self._title != newVal:
            self._title = newVal
            self.on_element_change()

    @property
    def desc(self):
        return self._desc

    @desc.setter
    def desc(self, newVal):
        if self._desc != newVal:
            self._desc = newVal
            self.on_element_change()

    def do_nothing(self):
        pass



class PlotLine(BasicElement):

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
        if self._sections != newVal:
            self._sections = newVal
            self.on_element_change()



class Chapter(BasicElement):

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
        if self._chLevel != newVal:
            self._chLevel = newVal
            self.on_element_change()

    @property
    def chType(self):
        return self._chType

    @chType.setter
    def chType(self, newVal):
        if self._chType != newVal:
            self._chType = newVal
            self.on_element_change()

    @property
    def noNumber(self):
        return self._noNumber

    @noNumber.setter
    def noNumber(self, newVal):
        if self._noNumber != newVal:
            self._noNumber = newVal
            self.on_element_change()

    @property
    def isTrash(self):
        return self._isTrash

    @isTrash.setter
    def isTrash(self, newVal):
        if self._isTrash != newVal:
            self._isTrash = newVal
            self.on_element_change()


class WorldElement(BasicElement):

    def __init__(self,
            aka=None,
            tags=None,
            links=None,
            **kwargs):
        super().__init__(**kwargs)
        self._aka = aka
        self._tags = tags
        self._links = links

    @property
    def aka(self):
        return self._aka

    @aka.setter
    def aka(self, newVal):
        if self._aka != newVal:
            self._aka = newVal
            self.on_element_change()

    @property
    def tags(self):
        return self._tags

    @tags.setter
    def tags(self, newVal):
        if self._tags != newVal:
            self._tags = newVal
            self.on_element_change()

    @property
    def links(self):
        try:
            return self._links.copy()
        except AttributeError:
            return None

    @links.setter
    def links(self, newVal):
        if self._links != newVal:
            self._links = newVal
            for linkPath in newVal:
                self._links[linkPath] = os.path.split(linkPath)[1]
            self.on_element_change()



class Character(WorldElement):
    MAJOR_MARKER = 'Major'
    MINOR_MARKER = 'Minor'

    def __init__(self,
            notes=None,
            bio=None,
            goals=None,
            fullName=None,
            isMajor=None,
            birthDate=None,
            deathDate=None,
            **kwargs):
        super().__init__(**kwargs)
        self._notes = notes
        self._bio = bio
        self._goals = goals
        self._fullName = fullName
        self._isMajor = isMajor
        self._birthDate = birthDate
        self._deathDate = deathDate

    @property
    def notes(self):
        return self._notes

    @notes.setter
    def notes(self, newVal):
        if self._notes != newVal:
            self._notes = newVal
            self.on_element_change()

    @property
    def bio(self):
        return self._bio

    @bio.setter
    def bio(self, newVal):
        if self._bio != newVal:
            self._bio = newVal
            self.on_element_change()

    @property
    def goals(self):
        return self._goals

    @goals.setter
    def goals(self, newVal):
        if self._goals != newVal:
            self._goals = newVal
            self.on_element_change()

    @property
    def fullName(self):
        return self._fullName

    @fullName.setter
    def fullName(self, newVal):
        if self._fullName != newVal:
            self._fullName = newVal
            self.on_element_change()

    @property
    def isMajor(self):
        return self._isMajor

    @isMajor.setter
    def isMajor(self, newVal):
        if self._isMajor != newVal:
            self._isMajor = newVal
            self.on_element_change()

    @property
    def birthDate(self):
        return self._birthDate

    @birthDate.setter
    def birthDate(self, newVal):
        if self._birthDate != newVal:
            self._birthDate = newVal
            self.on_element_change()

    @property
    def deathDate(self):
        return self._deathDate

    @deathDate.setter
    def deathDate(self, newVal):
        if self._deathDate != newVal:
            self._deathDate = newVal
            self.on_element_change()

from datetime import datetime, date, timedelta
import re


ADDITIONAL_WORD_LIMITS = re.compile('--|—|–|\<\/p\>')

NO_WORD_LIMITS = re.compile('\<note\>.*?\<\/note\>|\<comment\>.*?\<\/comment\>|\<.+?\>')


class Section(BasicElement):
    PACING = ['A', 'R', 'C']

    NULL_DATE = '0001-01-01'
    NULL_TIME = '00:00:00'

    def __init__(self,
            scType=None,
            scPacing=None,
            status=None,
            notes=None,
            tags=None,
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
        self._scPacing = scPacing
        self._status = status
        self._notes = notes
        self._tags = tags
        self._appendToPrev = appendToPrev
        self._goal = goal
        self._conflict = conflict
        self._outcome = outcome
        self._plotNotes = plotNotes
        try:
            self.weekDay = date.fromisoformat(scDate).weekday()
            self._date = scDate
        except:
            self.weekDay = None
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
        if self._scType != newVal:
            self._scType = newVal
            self.on_element_change()

    @property
    def scPacing(self):
        return self._scPacing

    @scPacing.setter
    def scPacing(self, newVal):
        if self._scPacing != newVal:
            self._scPacing = newVal
            self.on_element_change()

    @property
    def status(self):
        return self._status

    @status.setter
    def status(self, newVal):
        if self._status != newVal:
            self._status = newVal
            self.on_element_change()

    @property
    def notes(self):
        return self._notes

    @notes.setter
    def notes(self, newVal):
        if self._notes != newVal:
            self._notes = newVal
            self.on_element_change()

    @property
    def tags(self):
        return self._tags

    @tags.setter
    def tags(self, newVal):
        if self._tags != newVal:
            self._tags = newVal
            self.on_element_change()

    @property
    def appendToPrev(self):
        return self._appendToPrev

    @appendToPrev.setter
    def appendToPrev(self, newVal):
        if self._appendToPrev != newVal:
            self._appendToPrev = newVal
            self.on_element_change()

    @property
    def goal(self):
        return self._goal

    @goal.setter
    def goal(self, newVal):
        if self._goal != newVal:
            self._goal = newVal
            self.on_element_change()

    @property
    def conflict(self):
        return self._conflict

    @conflict.setter
    def conflict(self, newVal):
        if self._conflict != newVal:
            self._conflict = newVal
            self.on_element_change()

    @property
    def outcome(self):
        return self._outcome

    @outcome.setter
    def outcome(self, newVal):
        if self._outcome != newVal:
            self._outcome = newVal
            self.on_element_change()

    @property
    def plotNotes(self):
        try:
            return dict(self._plotNotes)
        except TypeError:
            return None

    @plotNotes.setter
    def plotNotes(self, newVal):
        if self._plotNotes != newVal:
            self._plotNotes = newVal
            self.on_element_change()

    @property
    def date(self):
        return self._date

    @date.setter
    def date(self, newVal):
        if self._date != newVal:
            if not newVal:
                self._date = None
                self.weekDay = None
            else:
                try:
                    self.weekDay = date.fromisoformat(newVal).weekday()
                except:
                    pass
                else:
                    self._date = newVal
                    self.on_element_change()

    @property
    def time(self):
        return self._time

    @time.setter
    def time(self, newVal):
        if self._time != newVal:
            self._time = newVal
            self.on_element_change()

    @property
    def day(self):
        return self._day

    @day.setter
    def day(self, newVal):
        if self._day != newVal:
            self._day = newVal
            self.on_element_change()

    @property
    def lastsMinutes(self):
        return self._lastsMinutes

    @lastsMinutes.setter
    def lastsMinutes(self, newVal):
        if self._lastsMinutes != newVal:
            self._lastsMinutes = newVal
            self.on_element_change()

    @property
    def lastsHours(self):
        return self._lastsHours

    @lastsHours.setter
    def lastsHours(self, newVal):
        if self._lastsHours != newVal:
            self._lastsHours = newVal
            self.on_element_change()

    @property
    def lastsDays(self):
        return self._lastsDays

    @lastsDays.setter
    def lastsDays(self, newVal):
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
        if self._items != newVal:
            self._items = newVal
            self.on_element_change()

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

    def day_to_date(self, referenceDate):
        if not self._date:
            try:
                deltaDays = timedelta(days=int(self._day))
                refDate = date.fromisoformat(referenceDate)
                self.date = date.isoformat(refDate + deltaDays)
                self._day = None
            except:
                self.date = ''
                return False

        return True

    def date_to_day(self, referenceDate):
        if not self._day:
            try:
                sectionDate = date.fromisoformat(self._date)
                referenceDate = date.fromisoformat(referenceDate)
                self.day = str((sectionDate - referenceDate).days)
                self._date = None
                self.weekDay = None
            except:
                self.day = ''
                return False

        return True



class PlotPoint(BasicElement):

    def __init__(self,
            sectionAssoc=None,
            notes=None,
            **kwargs):
        super().__init__(**kwargs)

        self._sectionAssoc = sectionAssoc

        self._notes = notes

    @property
    def sectionAssoc(self):
        return self._sectionAssoc

    @sectionAssoc.setter
    def sectionAssoc(self, newVal):
        if self._sectionAssoc != newVal:
            self._sectionAssoc = newVal
            self.on_element_change()

    @property
    def notes(self):
        return self._notes

    @notes.setter
    def notes(self, newVal):
        if self._notes != newVal:
            self._notes = newVal
            self.on_element_change()


__all__ = [
    'get_element_text',
    'text_to_xml_element',
    'xml_element_to_text',
    ]


def get_element_text(parent, tag, default=None):
    if parent.find(tag) is not None:
        return parent.find(tag).text
    else:
        return default


def text_to_xml_element(tag, text):
    xmlElement = ET.Element(tag)
    for line in text.split('\n'):
        ET.SubElement(xmlElement, 'p').text = line
    return xmlElement


def xml_element_to_text(xmlElement):
    lines = []
    if xmlElement:
        for paragraph in xmlElement.iterfind('p'):
            lines.append(''.join(t for t in paragraph.itertext()))
    return '\n'.join(lines)



class NovxFile(File):
    DESCRIPTION = _('novelibre project')
    EXTENSION = '.novx'

    MAJOR_VERSION = 1
    MINOR_VERSION = 1

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
        self.xmlTree = ET.parse(self.filePath)
        xmlRoot = self.xmlTree.getroot()
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

        try:
            locale = xmlRoot.attrib['{http://www.w3.org/XML/1998/namespace}lang']
            self.novel.languageCode, self.novel.countryCode = locale.split('-')
        except:
            pass
        self.novel.tree.reset()
        self._read_project(xmlRoot)
        self._read_locations(xmlRoot)
        self._read_items(xmlRoot)
        self._read_characters(xmlRoot)
        self._read_chapters(xmlRoot)
        self._read_plot_lines(xmlRoot)
        self._read_project_notes(xmlRoot)
        self.adjust_section_types()

        xmlWclog = xmlRoot.find('PROGRESS')
        if xmlWclog is not None:
            for xmlWc in xmlWclog.iterfind('WC'):
                wcDate = xmlWc.find('Date').text
                wcCount = xmlWc.find('Count').text
                wcTotalCount = xmlWc.find('WithUnused').text
                if wcDate and wcCount and wcTotalCount:
                    self.wcLog[wcDate] = [wcCount, wcTotalCount]

    def write(self):
        if self.novel.saveWordCount:
            newCountInt, newTotalCountInt = self.count_words()
            newCount = str(newCountInt)
            newTotalCount = str(newTotalCountInt)
            today = date.today().isoformat()
            self.wcLogUpdate[today] = [newCount, newTotalCount]
            for wcDate in self.wcLogUpdate:
                self.wcLog[wcDate] = self.wcLogUpdate[wcDate]
        self.wcLogUpdate = {}
        self.adjust_section_types()
        self.novel.get_languages()
        attrib = {'version':f'{self.MAJOR_VERSION}.{self.MINOR_VERSION}',
                'xml:lang':f'{self.novel.languageCode}-{self.novel.countryCode}',
                }
        xmlRoot = ET.Element('novx', attrib=attrib)
        self._build_element_tree(xmlRoot)
        indent(xmlRoot)
        self.xmlTree = ET.ElementTree(xmlRoot)
        self._write_element_tree(self)
        self._postprocess_xml_file(self.filePath)

    def _build_plot_line_branch(self, xmlPlotLines, prjPlotLine, plId):
        xmlPlotLine = ET.SubElement(xmlPlotLines, 'ARC', attrib={'id':plId})
        if prjPlotLine.title:
            ET.SubElement(xmlPlotLine, 'Title').text = prjPlotLine.title
        if prjPlotLine.shortName:
            ET.SubElement(xmlPlotLine, 'ShortName').text = prjPlotLine.shortName
        if prjPlotLine.desc:
            xmlPlotLine.append(text_to_xml_element('Desc', prjPlotLine.desc))

        if prjPlotLine.sections:
            attrib = {'ids':' '.join(prjPlotLine.sections)}
            ET.SubElement(xmlPlotLine, 'Sections', attrib=attrib)

        for ppId in self.novel.tree.get_children(plId):
            xmlPlotPoint = ET.SubElement(xmlPlotLine, 'POINT', attrib={'id':ppId})
            self._build_plot_point_branch(xmlPlotPoint, self.novel.plotPoints[ppId])

        return xmlPlotLine

    def _build_plot_point_branch(self, xmlPlotPoint, prjPlotPoint):
        if prjPlotPoint.title:
            ET.SubElement(xmlPlotPoint, 'Title').text = prjPlotPoint.title
        if prjPlotPoint.desc:
            xmlPlotPoint.append(text_to_xml_element('Desc', prjPlotPoint.desc))
        if prjPlotPoint.notes:
            xmlPlotPoint.append(text_to_xml_element('Notes', prjPlotPoint.notes))

        if prjPlotPoint.sectionAssoc:
            ET.SubElement(xmlPlotPoint, 'Section', attrib={'id': prjPlotPoint.sectionAssoc})

    def _build_chapter_branch(self, xmlChapters, prjChp, chId):
        xmlChapter = ET.SubElement(xmlChapters, 'CHAPTER', attrib={'id':chId})
        if prjChp.chType:
            xmlChapter.set('type', str(prjChp.chType))
        if prjChp.chLevel == 1:
            xmlChapter.set('level', '1')
        if prjChp.isTrash:
            xmlChapter.set('isTrash', '1')
        if prjChp.noNumber:
            xmlChapter.set('noNumber', '1')
        if prjChp.title:
            ET.SubElement(xmlChapter, 'Title').text = prjChp.title
        if prjChp.desc:
            xmlChapter.append(text_to_xml_element('Desc', prjChp.desc))
        for scId in self.novel.tree.get_children(chId):
            xmlSection = ET.SubElement(xmlChapter, 'SECTION', attrib={'id':scId})
            self._build_section_branch(xmlSection, self.novel.sections[scId])
        return xmlChapter

    def _build_character_branch(self, xmlCrt, prjCrt):
        if prjCrt.isMajor:
            xmlCrt.set('major', '1')
        if prjCrt.title:
            ET.SubElement(xmlCrt, 'Title').text = prjCrt.title
        if prjCrt.fullName:
            ET.SubElement(xmlCrt, 'FullName').text = prjCrt.fullName
        if prjCrt.aka:
            ET.SubElement(xmlCrt, 'Aka').text = prjCrt.aka
        if prjCrt.desc:
            xmlCrt.append(text_to_xml_element('Desc', prjCrt.desc))
        if prjCrt.bio:
            xmlCrt.append(text_to_xml_element('Bio', prjCrt.bio))
        if prjCrt.goals:
            xmlCrt.append(text_to_xml_element('Goals', prjCrt.goals))
        if prjCrt.notes:
            xmlCrt.append(text_to_xml_element('Notes', prjCrt.notes))
        tagStr = list_to_string(prjCrt.tags)
        if tagStr:
            ET.SubElement(xmlCrt, 'Tags').text = tagStr
        if prjCrt.links:
            for path in prjCrt.links:
                xmlLink = ET.SubElement(xmlCrt, 'Link')
                xmlLink.set('path', path)
        if prjCrt.birthDate:
            ET.SubElement(xmlCrt, 'BirthDate').text = prjCrt.birthDate
        if prjCrt.deathDate:
            ET.SubElement(xmlCrt, 'DeathDate').text = prjCrt.deathDate

    def _build_element_tree(self, root):

        xmlProject = ET.SubElement(root, 'PROJECT')
        self._build_project_branch(xmlProject)

        xmlChapters = ET.SubElement(root, 'CHAPTERS')
        for chId in self.novel.tree.get_children(CH_ROOT):
            self._build_chapter_branch(xmlChapters, self.novel.chapters[chId], chId)

        xmlCharacters = ET.SubElement(root, 'CHARACTERS')
        for crId in self.novel.tree.get_children(CR_ROOT):
            xmlCrt = ET.SubElement(xmlCharacters, 'CHARACTER', attrib={'id':crId})
            self._build_character_branch(xmlCrt, self.novel.characters[crId])

        xmlLocations = ET.SubElement(root, 'LOCATIONS')
        for lcId in self.novel.tree.get_children(LC_ROOT):
            xmlLoc = ET.SubElement(xmlLocations, 'LOCATION', attrib={'id':lcId})
            self._build_location_branch(xmlLoc, self.novel.locations[lcId])

        xmlItems = ET.SubElement(root, 'ITEMS')
        for itId in self.novel.tree.get_children(IT_ROOT):
            xmlItm = ET.SubElement(xmlItems, 'ITEM', attrib={'id':itId})
            self._build_item_branch(xmlItm, self.novel.items[itId])

        xmlPlotLines = ET.SubElement(root, 'ARCS')
        for plId in self.novel.tree.get_children(PL_ROOT):
            self._build_plot_line_branch(xmlPlotLines, self.novel.plotLines[plId], plId)

        xmlProjectNotes = ET.SubElement(root, 'PROJECTNOTES')
        for pnId in self.novel.tree.get_children(PN_ROOT):
            xmlProjectNote = ET.SubElement(xmlProjectNotes, 'PROJECTNOTE', attrib={'id':pnId})
            self._build_project_notes_branch(xmlProjectNote, self.novel.projectNotes[pnId])

        if self.wcLog:
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

    def _build_item_branch(self, xmlItm, prjItm):
        if prjItm.title:
            ET.SubElement(xmlItm, 'Title').text = prjItm.title
        if prjItm.aka:
            ET.SubElement(xmlItm, 'Aka').text = prjItm.aka
        if prjItm.desc:
            xmlItm.append(text_to_xml_element('Desc', prjItm.desc))
        tagStr = list_to_string(prjItm.tags)
        if tagStr:
            ET.SubElement(xmlItm, 'Tags').text = tagStr
        if prjItm.links:
            for path in prjItm.links:
                xmlLink = ET.SubElement(xmlItm, 'Link')
                xmlLink.set('path', path)

    def _build_location_branch(self, xmlLoc, prjLoc):
        if prjLoc.title:
            ET.SubElement(xmlLoc, 'Title').text = prjLoc.title
        if prjLoc.aka:
            ET.SubElement(xmlLoc, 'Aka').text = prjLoc.aka
        if prjLoc.desc:
            xmlLoc.append(text_to_xml_element('Desc', prjLoc.desc))
        tagStr = list_to_string(prjLoc.tags)
        if tagStr:
            ET.SubElement(xmlLoc, 'Tags').text = tagStr
        if prjLoc.links:
            for path in prjLoc.links:
                xmlLink = ET.SubElement(xmlLoc, 'Link')
                xmlLink.set('path', path)

    def _build_project_branch(self, xmlProject):
        if self.novel.renumberChapters:
            xmlProject.set('renumberChapters', '1')
        if self.novel.renumberParts:
            xmlProject.set('renumberParts', '1')
        if self.novel.renumberWithinParts:
            xmlProject.set('renumberWithinParts', '1')
        if self.novel.romanChapterNumbers:
            xmlProject.set('romanChapterNumbers', '1')
        if self.novel.romanPartNumbers:
            xmlProject.set('romanPartNumbers', '1')
        if self.novel.saveWordCount:
            xmlProject.set('saveWordCount', '1')
        if self.novel.workPhase is not None:
            xmlProject.set('workPhase', str(self.novel.workPhase))

        if self.novel.title:
            ET.SubElement(xmlProject, 'Title').text = self.novel.title
        if self.novel.authorName:
            ET.SubElement(xmlProject, 'Author').text = self.novel.authorName
        if self.novel.desc:
            xmlProject.append(text_to_xml_element('Desc', self.novel.desc))
        if self.novel.chapterHeadingPrefix:
            ET.SubElement(xmlProject, 'ChapterHeadingPrefix').text = self.novel.chapterHeadingPrefix
        if self.novel.chapterHeadingSuffix:
            ET.SubElement(xmlProject, 'ChapterHeadingSuffix').text = self.novel.chapterHeadingSuffix
        if self.novel.partHeadingPrefix:
            ET.SubElement(xmlProject, 'PartHeadingPrefix').text = self.novel.partHeadingPrefix
        if self.novel.partHeadingSuffix:
            ET.SubElement(xmlProject, 'PartHeadingSuffix').text = self.novel.partHeadingSuffix
        if self.novel.customGoal:
            ET.SubElement(xmlProject, 'CustomGoal').text = self.novel.customGoal
        if self.novel.customConflict:
            ET.SubElement(xmlProject, 'CustomConflict').text = self.novel.customConflict
        if self.novel.customOutcome:
            ET.SubElement(xmlProject, 'CustomOutcome').text = self.novel.customOutcome
        if self.novel.customChrBio:
            ET.SubElement(xmlProject, 'CustomChrBio').text = self.novel.customChrBio
        if self.novel.customChrGoals:
            ET.SubElement(xmlProject, 'CustomChrGoals').text = self.novel.customChrGoals
        if self.novel.wordCountStart:
            ET.SubElement(xmlProject, 'WordCountStart').text = str(self.novel.wordCountStart)
        if self.novel.wordTarget:
            ET.SubElement(xmlProject, 'WordTarget').text = str(self.novel.wordTarget)
        if self.novel.referenceDate:
            ET.SubElement(xmlProject, 'ReferenceDate').text = self.novel.referenceDate

    def _build_project_notes_branch(self, xmlProjectNote, projectNote):
        if projectNote.title:
            ET.SubElement(xmlProjectNote, 'Title').text = projectNote.title
        if projectNote.desc:
            xmlProjectNote.append(text_to_xml_element('Desc', projectNote.desc))

    def _build_section_branch(self, xmlSection, prjScn):
        if prjScn.scType:
            xmlSection.set('type', str(prjScn.scType))
        if prjScn.status > 1:
            xmlSection.set('status', str(prjScn.status))
        if prjScn.scPacing > 0:
            xmlSection.set('pacing', str(prjScn.scPacing))
        if prjScn.appendToPrev:
            xmlSection.set('append', '1')
        if prjScn.title:
            ET.SubElement(xmlSection, 'Title').text = prjScn.title
        if prjScn.desc:
            xmlSection.append(text_to_xml_element('Desc', prjScn.desc))
        if prjScn.goal:
            xmlSection.append(text_to_xml_element('Goal', prjScn.goal))
        if prjScn.conflict:
            xmlSection.append(text_to_xml_element('Conflict', prjScn.conflict))
        if prjScn.outcome:
            xmlSection.append(text_to_xml_element('Outcome', prjScn.outcome))
        if prjScn.plotNotes:
            xmlPlotNotes = ET.SubElement(xmlSection, 'PlotNotes')
            for plId in prjScn.plotNotes:
                if plId in prjScn.scPlotLines:
                    xmlPlotNote = text_to_xml_element('PlotlineNotes', prjScn.plotNotes[plId])
                    xmlPlotNote.set('id', plId)
                    xmlPlotNotes.append(xmlPlotNote)
        if prjScn.notes:
            xmlSection.append(text_to_xml_element('Notes', prjScn.notes))
        tagStr = list_to_string(prjScn.tags)
        if tagStr:
            ET.SubElement(xmlSection, 'Tags').text = tagStr

        if prjScn.date:
            ET.SubElement(xmlSection, 'Date').text = prjScn.date
        elif prjScn.day:
            ET.SubElement(xmlSection, 'Day').text = prjScn.day
        if prjScn.time:
            ET.SubElement(xmlSection, 'Time').text = prjScn.time

        if prjScn.lastsDays and prjScn.lastsDays != '0':
            ET.SubElement(xmlSection, 'LastsDays').text = prjScn.lastsDays
        if prjScn.lastsHours and prjScn.lastsHours != '0':
            ET.SubElement(xmlSection, 'LastsHours').text = prjScn.lastsHours
        if prjScn.lastsMinutes and prjScn.lastsMinutes != '0':
            ET.SubElement(xmlSection, 'LastsMinutes').text = prjScn.lastsMinutes

        if prjScn.characters:
            attrib = {'ids':' '.join(prjScn.characters)}
            ET.SubElement(xmlSection, 'Characters', attrib=attrib)
        if prjScn.locations:
            attrib = {'ids':' '.join(prjScn.locations)}
            ET.SubElement(xmlSection, 'Locations', attrib=attrib)
        if prjScn.items:
            attrib = {'ids':' '.join(prjScn.items)}
            ET.SubElement(xmlSection, 'Items', attrib=attrib)

        sectionContent = prjScn.sectionContent
        if sectionContent:
            if not sectionContent in ('<p></p>', '<p />'):
                xmlSection.append(ET.fromstring(f'<Content>{sectionContent}</Content>'))

    def _get_link_dict(self, parent):
        links = {}
        for xmlLink in parent.iterfind('Link'):
            path = xmlLink.attrib.get('path', None)
            if path:
                links[path] = None
        return links

    def _postprocess_xml_file(self, filePath):
        with open(filePath, 'r', encoding='utf-8') as f:
            text = f.read()
        try:
            with open(filePath, 'w', encoding='utf-8') as f:
                f.write(f'{self.XML_HEADER}{text}')
        except:
            raise Error(f'{_("Cannot write file")}: "{norm_path(filePath)}".')

    def _read_plot_lines(self, root):
        try:
            for xmlPlotLine in root.find('ARCS'):
                plId = xmlPlotLine.attrib['id']
                self.novel.plotLines[plId] = PlotLine(on_element_change=self.on_element_change)
                self.novel.plotLines[plId].title = get_element_text(xmlPlotLine, 'Title')
                self.novel.plotLines[plId].desc = xml_element_to_text(xmlPlotLine.find('Desc'))
                self.novel.plotLines[plId].shortName = get_element_text(xmlPlotLine, 'ShortName')
                self.novel.tree.append(PL_ROOT, plId)
                for xmlPlotPoint in xmlPlotLine.iterfind('POINT'):
                    ppId = xmlPlotPoint.attrib['id']
                    self._read_plot_point(xmlPlotPoint, ppId, plId)
                    self.novel.tree.append(plId, ppId)

                acSections = []
                xmlSections = xmlPlotLine.find('Sections')
                if xmlSections is not None:
                    scIds = xmlSections.get('ids', None)
                    for scId in string_to_list(scIds, divider=' '):
                        if scId and scId in self.novel.sections:
                            acSections.append(scId)
                            self.novel.sections[scId].scPlotLines.append(plId)
                self.novel.plotLines[plId].sections = acSections
        except TypeError:
            pass

    def _read_plot_point(self, xmlPoint, ppId, plId):
        self.novel.plotPoints[ppId] = PlotPoint(on_element_change=self.on_element_change)
        self.novel.plotPoints[ppId].title = get_element_text(xmlPoint, 'Title')
        self.novel.plotPoints[ppId].desc = xml_element_to_text(xmlPoint.find('Desc'))
        self.novel.plotPoints[ppId].notes = xml_element_to_text(xmlPoint.find('Notes'))
        xmlSectionAssoc = xmlPoint.find('Section')
        if xmlSectionAssoc is not None:
            scId = xmlSectionAssoc.get('id', None)
            self.novel.plotPoints[ppId].sectionAssoc = scId
            self.novel.sections[scId].scPlotPoints[ppId] = plId

    def _read_chapters(self, root, partType=0):
        try:
            for xmlChapter in root.find('CHAPTERS'):
                chId = xmlChapter.attrib['id']
                self.novel.chapters[chId] = Chapter(on_element_change=self.on_element_change)
                typeStr = xmlChapter.get('type', '0')
                if typeStr in ('0', '1'):
                    self.novel.chapters[chId].chType = int(typeStr)
                else:
                    self.novel.chapters[chId].chType = 1
                chLevel = xmlChapter.get('level', None)
                if chLevel == '1':
                    self.novel.chapters[chId].chLevel = 1
                else:
                    self.novel.chapters[chId].chLevel = 2
                self.novel.chapters[chId].isTrash = xmlChapter.get('isTrash', None) == '1'
                self.novel.chapters[chId].noNumber = xmlChapter.get('noNumber', None) == '1'
                self.novel.chapters[chId].title = get_element_text(xmlChapter, 'Title')
                self.novel.chapters[chId].desc = xml_element_to_text(xmlChapter.find('Desc'))
                self.novel.tree.append(CH_ROOT, chId)
                if xmlChapter.find('SECTION'):
                    for xmlSection in xmlChapter.iterfind('SECTION'):
                        scId = xmlSection.attrib['id']
                        self._read_section(xmlSection, scId)
                        if self.novel.sections[scId].scType < self.novel.chapters[chId].chType:
                            self.novel.sections[scId].scType = self.novel.chapters[chId].chType
                        self.novel.tree.append(chId, scId)
        except TypeError:
            pass

    def _read_characters(self, root):
        try:
            for xmlCharacter in root.find('CHARACTERS'):
                crId = xmlCharacter.attrib['id']
                self.novel.characters[crId] = Character(on_element_change=self.on_element_change)
                self.novel.characters[crId].isMajor = xmlCharacter.get('major', None) == '1'
                self.novel.characters[crId].title = get_element_text(xmlCharacter, 'Title')
                self.novel.characters[crId].links = self._get_link_dict(xmlCharacter)
                self.novel.characters[crId].desc = xml_element_to_text(xmlCharacter.find('Desc'))
                self.novel.characters[crId].aka = get_element_text(xmlCharacter, 'Aka')
                tags = string_to_list(get_element_text(xmlCharacter, 'Tags'))
                self.novel.characters[crId].tags = self._strip_spaces(tags)
                self.novel.characters[crId].notes = xml_element_to_text(xmlCharacter.find('Notes'))
                self.novel.characters[crId].bio = xml_element_to_text(xmlCharacter.find('Bio'))
                self.novel.characters[crId].goals = xml_element_to_text(xmlCharacter.find('Goals'))
                self.novel.characters[crId].fullName = get_element_text(xmlCharacter, 'FullName')
                self.novel.characters[crId].birthDate = get_element_text(xmlCharacter, 'BirthDate')
                self.novel.characters[crId].deathDate = get_element_text(xmlCharacter, 'DeathDate')
                self.novel.tree.append(CR_ROOT, crId)
        except TypeError:
            pass

    def _read_items(self, root):
        try:
            for xmlItem in root.find('ITEMS'):
                itId = xmlItem.attrib['id']
                self.novel.items[itId] = WorldElement(on_element_change=self.on_element_change)
                self.novel.items[itId].title = get_element_text(xmlItem, 'Title')
                self.novel.items[itId].desc = xml_element_to_text(xmlItem.find('Desc'))
                self.novel.items[itId].aka = get_element_text(xmlItem, 'Aka')
                tags = string_to_list(get_element_text(xmlItem, 'Tags'))
                self.novel.items[itId].tags = self._strip_spaces(tags)
                self.novel.items[itId].links = self._get_link_dict(xmlItem)
                self.novel.tree.append(IT_ROOT, itId)
        except TypeError:
            pass

    def _read_locations(self, root):
        try:
            for xmlLocation in root.find('LOCATIONS'):
                lcId = xmlLocation.attrib['id']
                self.novel.locations[lcId] = WorldElement(on_element_change=self.on_element_change)
                self.novel.locations[lcId].title = get_element_text(xmlLocation, 'Title')
                self.novel.locations[lcId].links = self._get_link_dict(xmlLocation)
                self.novel.locations[lcId].desc = xml_element_to_text(xmlLocation.find('Desc'))
                self.novel.locations[lcId].aka = get_element_text(xmlLocation, 'Aka')
                tags = string_to_list(get_element_text(xmlLocation, 'Tags'))
                self.novel.locations[lcId].tags = self._strip_spaces(tags)
                self.novel.tree.append(LC_ROOT, lcId)
        except TypeError:
            pass

    def _read_project(self, root):
        xmlProject = root.find('PROJECT')
        self.novel.renumberChapters = xmlProject.get('renumberChapters', None) == '1'
        self.novel.renumberParts = xmlProject.get('renumberParts', None) == '1'
        self.novel.renumberWithinParts = xmlProject.get('renumberWithinParts', None) == '1'
        self.novel.romanChapterNumbers = xmlProject.get('romanChapterNumbers', None) == '1'
        self.novel.romanPartNumbers = xmlProject.get('romanPartNumbers', None) == '1'
        self.novel.saveWordCount = xmlProject.get('saveWordCount', None) == '1'
        workPhase = xmlProject.get('workPhase', None)
        if workPhase in ('1', '2', '3', '4', '5'):
            self.novel.workPhase = int(workPhase)
        else:
            self.novel.workPhase = None
        self.novel.title = get_element_text(xmlProject, 'Title')
        self.novel.authorName = get_element_text(xmlProject, 'Author')
        self.novel.desc = xml_element_to_text(xmlProject.find('Desc'))
        self.novel.chapterHeadingPrefix = get_element_text(xmlProject, 'ChapterHeadingPrefix')
        self.novel.chapterHeadingSuffix = get_element_text(xmlProject, 'ChapterHeadingSuffix')
        self.novel.partHeadingPrefix = get_element_text(xmlProject, 'PartHeadingPrefix')
        self.novel.partHeadingSuffix = get_element_text(xmlProject, 'PartHeadingSuffix')
        self.novel.customGoal = get_element_text(xmlProject, 'CustomGoal')
        self.novel.customConflict = get_element_text(xmlProject, 'CustomConflict')
        self.novel.customOutcome = get_element_text(xmlProject, 'CustomOutcome')
        self.novel.customChrBio = get_element_text(xmlProject, 'CustomChrBio')
        self.novel.customChrGoals = get_element_text(xmlProject, 'CustomChrGoals')
        if xmlProject.find('WordCountStart') is not None:
            self.novel.wordCountStart = int(xmlProject.find('WordCountStart').text)
        if xmlProject.find('WordTarget') is not None:
            self.novel.wordTarget = int(xmlProject.find('WordTarget').text)
        self.novel.referenceDate = get_element_text(xmlProject, 'ReferenceDate')

    def _read_project_notes(self, root):
        try:
            for xmlProjectNote in root.find('PROJECTNOTES'):
                pnId = xmlProjectNote.attrib['id']
                self.novel.projectNotes[pnId] = BasicElement()
                self.novel.projectNotes[pnId].title = get_element_text(xmlProjectNote, 'Title')
                self.novel.projectNotes[pnId].desc = xml_element_to_text(xmlProjectNote.find('Desc'))
                self.novel.tree.append(PN_ROOT, pnId)
        except TypeError:
            pass

    def _read_section(self, xmlSection, scId):
        self.novel.sections[scId] = Section(on_element_change=self.on_element_change)
        typeStr = xmlSection.get('type', '0')
        if typeStr in ('0', '1', '2', '3'):
            self.novel.sections[scId].scType = int(typeStr)
        else:
            self.novel.sections[scId].scType = 1
        status = xmlSection.get('status', None)
        if status in ('2', '3', '4', '5'):
            self.novel.sections[scId].status = int(status)
        else:
            self.novel.sections[scId].status = 1
        scPacing = xmlSection.get('pacing', 0)
        if scPacing in ('1', '2'):
            self.novel.sections[scId].scPacing = int(scPacing)
        else:
            self.novel.sections[scId].scPacing = 0
        self.novel.sections[scId].appendToPrev = xmlSection.get('append', None) == '1'
        self.novel.sections[scId].title = get_element_text(xmlSection, 'Title')
        self.novel.sections[scId].desc = xml_element_to_text(xmlSection.find('Desc'))

        if xmlSection.find('Content'):
            xmlStr = ET.tostring(xmlSection.find('Content'),
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
                self.novel.sections[scId].sectionContent = xmlStr
            else:
                self.novel.sections[scId].sectionContent = '<p></p>'
        else:
            self.novel.sections[scId].sectionContent = '<p></p>'

        self.novel.sections[scId].notes = xml_element_to_text(xmlSection.find('Notes'))

        tags = string_to_list(get_element_text(xmlSection, 'Tags'))
        self.novel.sections[scId].tags = self._strip_spaces(tags)

        if xmlSection.find('Date') is not None:
            dateStr = xmlSection.find('Date').text
            try:
                date.fromisoformat(dateStr)
            except:
                self.novel.sections[scId].date = None
            else:
                self.novel.sections[scId].date = dateStr
        elif xmlSection.find('Day') is not None:
            dayStr = xmlSection.find('Day').text
            try:
                int(dayStr)
            except ValueError:
                self.novel.sections[scId].day = None
            else:
                self.novel.sections[scId].day = dayStr

        if xmlSection.find('Time') is not None:
            timeStr = xmlSection.find('Time').text
            try:
                time.fromisoformat(timeStr)
            except:
                self.novel.sections[scId].time = None
            else:
                self.novel.sections[scId].time = timeStr

        self.novel.sections[scId].lastsDays = get_element_text(xmlSection, 'LastsDays')
        self.novel.sections[scId].lastsHours = get_element_text(xmlSection, 'LastsHours')
        self.novel.sections[scId].lastsMinutes = get_element_text(xmlSection, 'LastsMinutes')

        self.novel.sections[scId].goal = xml_element_to_text(xmlSection.find('Goal'))
        self.novel.sections[scId].conflict = xml_element_to_text(xmlSection.find('Conflict'))
        self.novel.sections[scId].outcome = xml_element_to_text(xmlSection.find('Outcome'))

        xmlPlotNotes = xmlSection.find('PlotNotes')
        if xmlPlotNotes is not None:
            plotNotes = {}
            for xmlPlotLineNote in xmlPlotNotes.iterfind('PlotlineNotes'):
                plId = xmlPlotLineNote.get('id', None)
                plotNotes[plId] = xml_element_to_text(xmlPlotLineNote)
            self.novel.sections[scId].plotNotes = plotNotes

        scCharacters = []
        xmlCharacters = xmlSection.find('Characters')
        if xmlCharacters is not None:
            crIds = xmlCharacters.get('ids', None)
            for crId in string_to_list(crIds, divider=' '):
                if crId and crId in self.novel.characters:
                    scCharacters.append(crId)
        self.novel.sections[scId].characters = scCharacters

        scLocations = []
        xmlLocations = xmlSection.find('Locations')
        if xmlLocations is not None:
            lcIds = xmlLocations.get('ids', None)
            for lcId in string_to_list(lcIds, divider=' '):
                if lcId and lcId in self.novel.locations:
                    scLocations.append(lcId)
        self.novel.sections[scId].locations = scLocations

        scItems = []
        xmlItems = xmlSection.find('Items')
        if xmlItems is not None:
            itIds = xmlItems.get('ids', None)
            for itId in string_to_list(itIds, divider=' '):
                if itId and itId in self.novel.items:
                    scItems.append(itId)
        self.novel.sections[scId].items = scItems

    def _strip_spaces(self, lines):
        stripped = []
        for line in lines:
            stripped.append(line.strip())
        return stripped

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
        except Error:
            if backedUp:
                os.replace(f'{xmlProject.filePath}.bak', xmlProject.filePath)
            raise Error(f'{_("Cannot write file")}: "{norm_path(xmlProject.filePath)}".')

from datetime import date


LANGUAGE_TAG = re.compile('\<span xml\:lang=\"(.*?)\"\>')


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
        if self._authorName != newVal:
            self._authorName = newVal
            self.on_element_change()

    @property
    def wordTarget(self):
        return self._wordTarget

    @wordTarget.setter
    def wordTarget(self, newVal):
        if self._wordTarget != newVal:
            self._wordTarget = newVal
            self.on_element_change()

    @property
    def wordCountStart(self):
        return self._wordCountStart

    @wordCountStart.setter
    def wordCountStart(self, newVal):
        if self._wordCountStart != newVal:
            self._wordCountStart = newVal
            self.on_element_change()

    @property
    def languageCode(self):
        return self._languageCode

    @languageCode.setter
    def languageCode(self, newVal):
        if self._languageCode != newVal:
            self._languageCode = newVal
            self.on_element_change()

    @property
    def countryCode(self):
        return self._countryCode

    @countryCode.setter
    def countryCode(self, newVal):
        if self._countryCode != newVal:
            self._countryCode = newVal
            self.on_element_change()

    @property
    def renumberChapters(self):
        return self._renumberChapters

    @renumberChapters.setter
    def renumberChapters(self, newVal):
        if self._renumberChapters != newVal:
            self._renumberChapters = newVal
            self.on_element_change()

    @property
    def renumberParts(self):
        return self._renumberParts

    @renumberParts.setter
    def renumberParts(self, newVal):
        if self._renumberParts != newVal:
            self._renumberParts = newVal
            self.on_element_change()

    @property
    def renumberWithinParts(self):
        return self._renumberWithinParts

    @renumberWithinParts.setter
    def renumberWithinParts(self, newVal):
        if self._renumberWithinParts != newVal:
            self._renumberWithinParts = newVal
            self.on_element_change()

    @property
    def romanChapterNumbers(self):
        return self._romanChapterNumbers

    @romanChapterNumbers.setter
    def romanChapterNumbers(self, newVal):
        if self._romanChapterNumbers != newVal:
            self._romanChapterNumbers = newVal
            self.on_element_change()

    @property
    def romanPartNumbers(self):
        return self._romanPartNumbers

    @romanPartNumbers.setter
    def romanPartNumbers(self, newVal):
        if self._romanPartNumbers != newVal:
            self._romanPartNumbers = newVal
            self.on_element_change()

    @property
    def saveWordCount(self):
        return self._saveWordCount

    @saveWordCount.setter
    def saveWordCount(self, newVal):
        if self._saveWordCount != newVal:
            self._saveWordCount = newVal
            self.on_element_change()

    @property
    def workPhase(self):
        return self._workPhase

    @workPhase.setter
    def workPhase(self, newVal):
        if self._workPhase != newVal:
            self._workPhase = newVal
            self.on_element_change()

    @property
    def chapterHeadingPrefix(self):
        return self._chapterHeadingPrefix

    @chapterHeadingPrefix.setter
    def chapterHeadingPrefix(self, newVal):
        if self._chapterHeadingPrefix != newVal:
            self._chapterHeadingPrefix = newVal
            self.on_element_change()

    @property
    def chapterHeadingSuffix(self):
        return self._chapterHeadingSuffix

    @chapterHeadingSuffix.setter
    def chapterHeadingSuffix(self, newVal):
        if self._chapterHeadingSuffix != newVal:
            self._chapterHeadingSuffix = newVal
            self.on_element_change()

    @property
    def partHeadingPrefix(self):
        return self._partHeadingPrefix

    @partHeadingPrefix.setter
    def partHeadingPrefix(self, newVal):
        if self._partHeadingPrefix != newVal:
            self._partHeadingPrefix = newVal
            self.on_element_change()

    @property
    def partHeadingSuffix(self):
        return self._partHeadingSuffix

    @partHeadingSuffix.setter
    def partHeadingSuffix(self, newVal):
        if self._partHeadingSuffix != newVal:
            self._partHeadingSuffix = newVal
            self.on_element_change()

    @property
    def customGoal(self):
        return self._customGoal

    @customGoal.setter
    def customGoal(self, newVal):
        if self._customGoal != newVal:
            self._customGoal = newVal
            self.on_element_change()

    @property
    def customConflict(self):
        return self._customConflict

    @customConflict.setter
    def customConflict(self, newVal):
        if self._customConflict != newVal:
            self._customConflict = newVal
            self.on_element_change()

    @property
    def customOutcome(self):
        return self._customOutcome

    @customOutcome.setter
    def customOutcome(self, newVal):
        if self._customOutcome != newVal:
            self._customOutcome = newVal
            self.on_element_change()

    @property
    def customChrBio(self):
        return self._customChrBio

    @customChrBio.setter
    def customChrBio(self, newVal):
        if self._customChrBio != newVal:
            self._customChrBio = newVal
            self.on_element_change()

    @property
    def customChrGoals(self):
        return self._customChrGoals

    @customChrGoals.setter
    def customChrGoals(self, newVal):
        if self._customChrGoals != newVal:
            self._customChrGoals = newVal
            self.on_element_change()

    @property
    def referenceDate(self):
        return self._referenceDate

    @referenceDate.setter
    def referenceDate(self, newVal):
        if self._referenceDate != newVal:
            if not newVal:
                self._referenceDate = None
                self.referenceWeekDay = None
            else:
                try:
                    self.referenceWeekDay = date.fromisoformat(newVal).weekday()
                except:
                    pass
                else:
                    self._referenceDate = newVal
                    self.on_element_change()

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

    def get_languages(self):

        def languages(text):
            if text:
                m = LANGUAGE_TAG.search(text)
                while m:
                    text = text[m.span()[1]:]
                    yield m.group(1)
                    m = LANGUAGE_TAG.search(text)

        self.languages = []
        for scId in self.sections:
            text = self.sections[scId].sectionContent
            if text:
                for language in languages(text):
                    if not language in self.languages:
                        self.languages.append(language)

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
        elif parent.startswith(CHAPTER_PREFIX):
            try:
                self.srtSections[parent].append(iid)
            except:
                self.srtSections[parent] = [iid]
        elif parent.startswith(PLOT_LINE_PREFIX):
            try:
                self.srtTurningPoints[parent].append(iid)
            except:
                self.srtTurningPoints[parent] = [iid]

    def delete(self, *items):
        raise NotImplementedError

    def delete_children(self, parent):
        if parent in self.roots:
            self.roots[parent] = []
            if parent == CH_ROOT:
                self.srtSections = {}
            elif parent == PL_ROOT:
                self.srtTurningPoints = {}
        elif parent.startswith(CHAPTER_PREFIX):
            self.srtSections[parent] = []
        elif parent.startswith(PLOT_LINE_PREFIX):
            self.srtTurningPoints[parent] = []

    def get_children(self, item):
        if item in self.roots:
            return self.roots[item]

        elif item.startswith(CHAPTER_PREFIX):
            return self.srtSections.get(item, [])

        elif item.startswith(PLOT_LINE_PREFIX):
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
        elif parent.startswith(CHAPTER_PREFIX):
            try:
                self.srtSections[parent].insert(index, iid)
            except:
                self.srtSections[parent] = [iid]
        elif parent.startswith(PLOT_LINE_PREFIX):
            try:
                self.srtTurningPoints.insert(index, iid)
            except:
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
            elif item == PL_ROOT:
                self.srtTurningPoints = {}
        elif item.startswith(CHAPTER_PREFIX):
            self.srtSections[item] = newchildren[:]
        elif item.startswith(PLOT_LINE_PREFIX):
            self.srtTurningPoints[item] = newchildren[:]




def yw_novx(sourcePath):
    path, extension = os.path.splitext(sourcePath)
    if extension != '.yw7':
        raise ValueError(f'File must be .yw7 type, but is "{extension}".')

    targetPath = f'{path}.novx'
    source = Yw7File(sourcePath)
    target = NovxFile(targetPath)
    source.novel = Novel(tree=NvTree())
    source.read()
    target.novel = source.novel
    target.wcLog = source.wcLog
    target.write()



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

            yw7Path = xmlPath.text
            if yw7Path and os.path.isfile(yw7Path):
                bookPath, bookExt = os.path.splitext(yw7Path)
                if bookExt == '.yw7':
                    novxPath = f'{bookPath}.novx'
                    if not os.path.isfile(novxPath):

                        yw_novx(yw7Path)
                    ET.SubElement(targetElement, 'Path').text = novxPath

    pathRoot , extension = os.path.splitext(sourcePath)
    if extension != '.pwc':
        raise ValueError(f'File must be .pwc type, but is "{extension}".')
    targetPath = f'{pathRoot}.nvcx'

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

