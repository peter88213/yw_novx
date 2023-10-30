"""Convert yw7 to novx.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/yw2novx
License: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.en.html)
"""
SUFFIX = ''

import sys
import os

import re
from datetime import datetime
from html import unescape
import xml.etree.ElementTree as ET
import gettext
import locale

__all__ = [
    'Error',
    '_',
    'LOCALE_PATH',
    'CURRENT_LANGUAGE',
    'norm_path',
    'string_to_list',
    'list_to_string',
    'ROOT_PREFIX',
    'ARC_PREFIX',
    'ARC_POINT_PREFIX',
    'CHAPTER_PREFIX',
    'SCENE_PREFIX',
    'CHARACTER_PREFIX',
    'LOCATION_PREFIX',
    'ITEM_PREFIX',
    'CH_ROOT',
    'AC_ROOT',
    'CR_ROOT',
    'LC_ROOT',
    'IT_ROOT',
    ]
ROOT_PREFIX = 'rt'
CHAPTER_PREFIX = 'ch'
ARC_PREFIX = 'ac'
SCENE_PREFIX = 'sc'
ARC_POINT_PREFIX = 'ap'
CHARACTER_PREFIX = 'cr'
LOCATION_PREFIX = 'lc'
ITEM_PREFIX = 'it'
CH_ROOT = f'{ROOT_PREFIX}{CHAPTER_PREFIX}'
AC_ROOT = f'{ROOT_PREFIX}{ARC_PREFIX}'
CR_ROOT = f'{ROOT_PREFIX}{CHARACTER_PREFIX}'
LC_ROOT = f'{ROOT_PREFIX}{LOCATION_PREFIX}'
IT_ROOT = f'{ROOT_PREFIX}{ITEM_PREFIX}'


class Error(Exception):
    pass


LOCALE_PATH = f'{os.path.dirname(sys.argv[0])}/locale/'
try:
    CURRENT_LANGUAGE = locale.getlocale()[0][:2]
except:
    CURRENT_LANGUAGE = locale.getdefaultlocale()[0][:2]
try:
    t = gettext.translation('novelyst', LOCALE_PATH, languages=[CURRENT_LANGUAGE])
    _ = t.gettext
except:

    def _(message):
        return message


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

from abc import ABC
from urllib.parse import quote


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


LANGUAGE_TAG = re.compile('\[lang=(.*?)\]')


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
        self.scenes = {}
        self.arcPoints = {}
        self.languages = None
        self.arcs = {}
        self.locations = {}
        self.items = {}
        self.characters = {}
        self.projectNotes = {}

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

    def update_scene_arcs(self):
        for scId in self.scenes:
            self.scenes[scId].scArcPoints = {}
            self.scenes[scId].scArcs = []
            for acId in self.arcs:
                if scId in self.arcs[acId].scenes:
                    self.scenes[scId].scArcs.append(acId)
                    for apId in self.tree.get_children(acId):
                        if self.arcPoints[apId].sceneAssoc == scId:
                            self.scenes[scId].scArcPoints[apId] = acId
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
        for scId in self.scenes:
            text = self.scenes[scId].sceneContent
            if text:
                for language in languages(text):
                    if not language in self.languages:
                        self.languages.append(language)

    def check_locale(self):
        if not self._languageCode:
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
        self.scenesSplit = False
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

    def _convert_from_novelyst(self, text:str, quick:bool=False):
        return text.rstrip()

    def _convert_to_novelyst(self, text:str):
        return text.rstrip()



class Arc(BasicElement):

    def __init__(self,
            shortName=None,
            **kwargs):
        super().__init__(**kwargs)

        self._shortName = shortName

        self._scenes = None

    @property
    def shortName(self):
        return self._shortName

    @shortName.setter
    def shortName(self, newVal):
        if self._shortName != newVal:
            self._shortName = newVal
            self.on_element_change()

    @property
    def scenes(self):
        try:
            return self._scenes[:]
        except TypeError:
            return None

    @scenes.setter
    def scenes(self, newVal):
        if self._scenes != newVal:
            self._scenes = newVal
            self.on_element_change()



class ArcPoint(BasicElement):

    def __init__(self,
            sceneAssoc=None,
            notes=None,
            **kwargs):
        super().__init__(**kwargs)

        self._sceneAssoc = sceneAssoc

        self._notes = notes

    @property
    def sceneAssoc(self):
        return self._sceneAssoc

    @sceneAssoc.setter
    def sceneAssoc(self, newVal):
        if self._sceneAssoc != newVal:
            self._sceneAssoc = newVal
            self.on_element_change()

    @property
    def notes(self):
        return self._notes

    @notes.setter
    def notes(self, newVal):
        if self._notes != newVal:
            self._notes = newVal
            self.on_element_change()



class WorldElement(BasicElement):

    def __init__(self,
            aka=None,
            tags=None,
            image=None,
            **kwargs):
        super().__init__(**kwargs)
        self._aka = aka
        self._image = image
        self._tags = tags

    @property
    def aka(self):
        return self._aka

    @aka.setter
    def aka(self, newVal):
        if self._aka != newVal:
            self._aka = newVal
            self.on_element_change()

    @property
    def image(self):
        return self._image

    @image.setter
    def image(self, newVal):
        if self._image != newVal:
            self._image = newVal
            self.on_element_change()

    @property
    def tags(self):
        return self._tags

    @tags.setter
    def tags(self, newVal):
        if self._tags != newVal:
            self._tags = newVal
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
            **kwargs):
        super().__init__(**kwargs)
        self._notes = notes
        self._bio = bio
        self._goals = goals
        self._fullName = fullName
        self._isMajor = isMajor

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


def create_id(elements, prefix=''):
    i = 1
    while f'{prefix}{i}' in elements:
        i += 1
    return f'{prefix}{i}'



ADDITIONAL_WORD_LIMITS = re.compile('--|—|–')
NO_WORD_LIMITS = re.compile('\[.+?\]|\/\*.+?\*\/|-|^\>', re.MULTILINE)


class Scene(BasicElement):
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
            date=None,
            time=None,
            day=None,
            lastsMinutes=None,
            lastsHours=None,
            lastsDays=None,
            scMode=None,
            stageLevel=None,
            **kwargs):
        super().__init__(**kwargs)
        self._sceneContent = None
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
        self._date = date
        self._time = time
        self._day = day
        self._lastsMinutes = lastsMinutes
        self._lastsHours = lastsHours
        self._lastsDays = lastsDays
        self._scMode = scMode
        self._stageLevel = stageLevel

        self._characters = None
        self._locations = None
        self._items = None

        self.scArcs = []
        self.scArcPoints = {}

    @property
    def sceneContent(self):
        return self._sceneContent

    @sceneContent.setter
    def sceneContent(self, text):
        self._sceneContent = text
        text = ADDITIONAL_WORD_LIMITS.sub(' ', text)
        text = NO_WORD_LIMITS.sub('', text)
        wordList = text.split()
        self.wordCount = len(wordList)

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
    def date(self):
        return self._date

    @date.setter
    def date(self, newVal):
        if self._date != newVal:
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
    def scMode(self):
        return self._scMode

    @scMode.setter
    def scMode(self, newVal):
        if self._scMode != newVal:
            self._scMode = newVal
            self.on_element_change()

    @property
    def stageLevel(self):
        return self._stageLevel

    @stageLevel.setter
    def stageLevel(self, newVal):
        if self._stageLevel != newVal:
            self._stageLevel = newVal
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

    def __init__(self, filePath, **kwargs):
        super().__init__(filePath)
        self.tree = None
        self.wcLog = {}
        self._ywApIds = None

    def is_locked(self):
        return os.path.isfile(f'{self.filePath}.lock')

    def read(self):
        if self.is_locked():
            raise Error(f'{_("yWriter seems to be open. Please close first")}.')
        try:
            self.tree = ET.parse(self.filePath)
        except:
            raise Error(f'{_("Can not process file")}: "{norm_path(self.filePath)}".')

        self._ywApIds = []
        self.wcLog = {}
        root = self.tree.getroot()
        self._read_project(root)
        self._read_locations(root)
        self._read_items(root)
        self._read_characters(root)
        self._read_projectvars(root)
        self._read_chapters(root)
        self._read_scenes(root)
        self._read_projectnotes(root)

        xmlWclog = root.find('WCLog')
        if xmlWclog is not None:
            for xmlWc in xmlWclog.findall('WC'):
                wcDate = xmlWc.find('Date').text
                wcCount = xmlWc.find('Count').text
                wcTotalCount = xmlWc.find('TotalCount').text
                self.wcLog[wcDate] = [wcCount, wcTotalCount]

        for scId in self.novel.scenes:
            if self.novel.scenes[scId].characters is None:
                self.novel.scenes[scId].characters = []
            if self.novel.scenes[scId].locations is None:
                self.novel.scenes[scId].locations = []
            if self.novel.scenes[scId].items is None:
                self.novel.scenes[scId].items = []
            if self.novel.scenes[scId].status is None:
                self.novel.scenes[scId].status = 1

        self.novel.check_locale()

    def write(self):
        if self.is_locked():
            raise Error(f'{_("yWriter seems to be open. Please close first")}.')

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

        def build_scene_subtree(xmlScene, prjScn, arcPoint=False):
            i = 1
            i = set_element(xmlScene, 'Title', prjScn.title, i)
            if prjScn.desc is not None:
                ET.SubElement(xmlScene, 'Desc').text = prjScn.desc


            scTypeEncoding = (
                (False, None),
                (True, '1'),
                (True, '2'),
                (True, '0'),
                )
            if arcPoint:
                scType = 2
            elif prjScn.scType is None:
                scType = 0
            else:
                scType = prjScn.scType
            yUnused, ySceneType = scTypeEncoding[scType]
            if yUnused:
                ET.SubElement(xmlScene, 'Unused').text = '-1'
            if ySceneType is not None:
                ET.SubElement(xmlSceneFields[scId], 'Field_SceneType').text = ySceneType
            if arcPoint:
                ET.SubElement(xmlScene, 'Status').text = '1'
            elif prjScn.status is not None:
                ET.SubElement(xmlScene, 'Status').text = str(prjScn.status)

            if arcPoint:
                ET.SubElement(xmlScene, 'SceneContent')
                return

            ET.SubElement(xmlScene, 'SceneContent').text = prjScn.sceneContent
            if prjScn.scMode:
                ET.SubElement(xmlSceneFields[scId], 'Field_SceneMode').text = str(prjScn.scMode)
            if prjScn.notes:
                ET.SubElement(xmlScene, 'Notes').text = prjScn.notes
            if prjScn.tags:
                ET.SubElement(xmlScene, 'Tags').text = list_to_string(prjScn.tags)
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

            if prjScn.scPacing == 1:
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

        def build_chapter_subtree(xmlChapter, prjChp, acId=None):

            chTypeEncoding = (
                (False, '0', '0'),
                (True, '1', '1'),
                (True, '1', '2'),
                (True, '1', '0'),
                )
            if acId is not None:
                chType = 2
            elif prjChp.chType is None:
                chType = 0
            else:
                chType = prjChp.chType
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
            if acId is None and prjChp.isTrash:
                ET.SubElement(xmlChapterFields, 'Field_IsTrash').text = '1'

            if acId is None:
                fields = { 'Field_NoNumber': isTrue(prjChp.noNumber)}
            else:
                fields = {'Field_ArcDefinition': self.novel.arcs[acId].shortName}
            for field in fields:
                if fields[field]:
                    ET.SubElement(xmlChapterFields, field).text = fields[field]
            if acId is None and prjChp.chLevel == 1:
                ET.SubElement(xmlChapter, 'SectionStart').text = '-1'
                i += 1
            i = set_element(xmlChapter, 'Type', yType, i)
            i = set_element(xmlChapter, 'ChapterType', yChapterType, i)

        def build_location_subtree(xmlLoc, prjLoc):
            if prjLoc.title:
                ET.SubElement(xmlLoc, 'Title').text = prjLoc.title
            if prjLoc.image:
                ET.SubElement(xmlLoc, 'ImageFile').text = prjLoc.image
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
            if prjItm.image:
                ET.SubElement(xmlItm, 'ImageFile').text = prjItm.image
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
            if prjCrt.image:
                ET.SubElement(xmlCrt, 'ImageFile').text = prjCrt.image
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

            fields = {
                'Field_WorkPhase': str(self.novel.workPhase),
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
                }
            xmlProjectFields = ET.SubElement(xmlProject, 'Fields')
            for field in fields:
                if fields[field]:
                    ET.SubElement(xmlProjectFields, field).text = fields[field]

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
        scIds = list(self.novel.scenes)
        for scId in scIds:
            xmlScene = ET.SubElement(xmlScenes, 'SCENE')
            ET.SubElement(xmlScene, 'ID').text = scId[2:]
            xmlSceneFields[scId] = ET.SubElement(xmlScene, 'Fields')
            build_scene_subtree(xmlScene, self.novel.scenes[scId])

        newScIds = {}
        for apId in self.novel.arcPoints:
            scId = create_id(scIds, prefix=SCENE_PREFIX)
            scIds.append(scId)
            newScIds[apId] = scId
            xmlScene = ET.SubElement(xmlScenes, 'SCENE')
            ET.SubElement(xmlScene, 'ID').text = scId[2:]
            xmlSceneFields[scId] = ET.SubElement(xmlScene, 'Fields')
            build_scene_subtree(xmlScene, self.novel.arcPoints[apId], arcPoint=True)

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
        arcPart = Chapter(title=_('Arcs'), chType=2, chLevel=1)
        build_chapter_subtree(xmlChapter, arcPart)
        for acId in self.novel.tree.get_children(AC_ROOT):
            chId = create_id(chIds, prefix=CHAPTER_PREFIX)
            chIds.append(chId)
            xmlChapter = ET.SubElement(xmlChapters, 'CHAPTER')
            ET.SubElement(xmlChapter, 'ID').text = chId[2:]
            build_chapter_subtree(xmlChapter, self.novel.arcs[acId], acId=acId)
            srtScenes = self.novel.tree.get_children(acId)
            if srtScenes:
                xmlScnList = ET.SubElement(xmlChapter, 'Scenes')
                for apId in srtScenes:
                    ET.SubElement(xmlScnList, 'ScID').text = newScIds[apId][2:]

        sceneArcs = {}
        sceneAssoc = {}
        for scId in scIds:
            sceneArcs[scId] = []
            sceneAssoc[scId] = []
        for acId in self.novel.arcs:
            for scId in self.novel.arcs[acId].scenes:
                sceneArcs[scId].append(self.novel.arcs[acId].shortName)
            for apId in self.novel.tree.get_children(acId):
                sceneArcs[newScIds[apId]].append(self.novel.arcs[acId].shortName)
        for apId in self.novel.arcPoints:
            if self.novel.arcPoints[apId].sceneAssoc:
                sceneAssoc[self.novel.arcPoints[apId].sceneAssoc].append(newScIds[apId][2:])
                sceneAssoc[newScIds[apId]].append(self.novel.arcPoints[apId].sceneAssoc[2:])
        for scId in scIds:
            fields = {
                'Field_SceneArcs': list_to_string(sceneArcs[scId]),
                'Field_SceneAssoc': list_to_string(sceneAssoc[scId]),
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

    def _postprocess_xml_file(self, filePath):
        with open(filePath, 'r', encoding='utf-8') as f:
            text = f.read()
        lines = text.split('\n')
        newlines = ['<?xml version="1.0" encoding="utf-8"?>']
        for line in lines:
            for tag in self._CDATA_TAGS:
                line = re.sub(f'\<{tag}\>', f'<{tag}><![CDATA[', line)
                line = re.sub(f'\<\/{tag}\>', f']]></{tag}>', line)
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
            self.novel.locations[lcId] = WorldElement()

            if xmlLocation.find('Title') is not None:
                self.novel.locations[lcId].title = xmlLocation.find('Title').text

            if xmlLocation.find('ImageFile') is not None:
                self.novel.locations[lcId].image = xmlLocation.find('ImageFile').text

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
            self.novel.items[itId] = WorldElement()

            if xmlItem.find('Title') is not None:
                self.novel.items[itId].title = xmlItem.find('Title').text

            if xmlItem.find('ImageFile') is not None:
                self.novel.items[itId].image = xmlItem.find('ImageFile').text

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
        self.novel.tree.delete_children(AC_ROOT)
        for xmlChapter in root.find('CHAPTERS'):
            prjChapter = Chapter()

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
                    prjChapter.chType = 2
                elif yChapterType == '1':
                    prjChapter.chType = 1
                elif yUnused:
                    prjChapter.chType = 3
            else:
                if xmlChapter.find('Type') is not None:
                    yType = xmlChapter.find('Type').text
                    if yType == '1':
                        prjChapter.chType = 1
                    elif yUnused:
                        prjChapter.chType = 3

            kwVarYw7 = {}
            for xmlChapterFields in xmlChapter.findall('Fields'):
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
                for scn in xmlChapter.find('Scenes').findall('ScID'):
                    scId = scn.text
                    scenes.append(scId)

            if shortName:
                acId = f"{ARC_PREFIX}{xmlChapter.find('ID').text}"
                self.novel.arcs[acId] = Arc()
                self.novel.arcs[acId].title = prjChapter.title
                self.novel.arcs[acId].desc = prjChapter.desc
                self.novel.arcs[acId].shortName = shortName
                self.novel.tree.append(AC_ROOT, acId)
                for scId in scenes:
                    self.novel.tree.append(acId, f'{ARC_POINT_PREFIX}{scId}')
                    self._ywApIds.append(scId)
            else:
                chId = f"{CHAPTER_PREFIX}{xmlChapter.find('ID').text}"
                self.novel.chapters[chId] = prjChapter
                self.novel.tree.append(CH_ROOT, chId)
                for scId in scenes:
                    self.novel.tree.append(chId, f'{SCENE_PREFIX}{scId}')

    def _read_characters(self, root):
        self.novel.tree.delete_children(CR_ROOT)
        for xmlCharacter in root.find('CHARACTERS'):
            crId = f"{CHARACTER_PREFIX}{xmlCharacter.find('ID').text}"
            self.novel.tree.append(CR_ROOT, crId)
            self.novel.characters[crId] = Character()

            if xmlCharacter.find('Title') is not None:
                self.novel.characters[crId].title = xmlCharacter.find('Title').text

            if xmlCharacter.find('ImageFile') is not None:
                self.novel.characters[crId].image = xmlCharacter.find('ImageFile').text

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
        for xmlProjectFields in xmlProject.findall('Fields'):
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

    def _read_projectnotes(self, root):
        if root.find('PROJECTNOTES') is not None:
            chId = create_id(self.novel.chapters, prefix=CHAPTER_PREFIX)
            self.novel.chapters[chId] = Chapter(title=_('Project notes'), chLevel=1, chType=1)
            self.novel.tree.append(CH_ROOT, chId)
            for xmlProjectnote in root.find('PROJECTNOTES'):
                scId = create_id(self.novel.scenes, prefix=SCENE_PREFIX)
                self.novel.tree.append(chId, scId)
                self.novel.scenes[scId] = Scene(scType=1, scPacing=0, status=0)
                if xmlProjectnote.find('Title') is not None:
                    self.novel.scenes[scId].title = xmlProjectnote.find('Title').text
                if xmlProjectnote.find('Desc') is not None:
                    self.novel.scenes[scId].desc = xmlProjectnote.find('Desc').text

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
            prjScn = Scene()

            if xmlScene.find('Title') is not None:
                prjScn.title = xmlScene.find('Title').text

            if xmlScene.find('Desc') is not None:
                prjScn.desc = xmlScene.find('Desc').text

            if xmlScene.find('SceneContent') is not None:
                sceneContent = xmlScene.find('SceneContent').text
                if sceneContent is not None:
                    prjScn.sceneContent = self._remove_inline_code(sceneContent)



            prjScn.scType = 0
            kwVarYw7 = {}
            for xmlSceneFields in xmlScene.findall('Fields'):
                if xmlSceneFields.find('Field_SceneType') is not None:
                    if xmlSceneFields.find('Field_SceneType').text == '1':
                        prjScn.scType = 1
                    elif xmlSceneFields.find('Field_SceneType').text == '2':
                        prjScn.scType = 2

                for fieldName in self.SCN_KWVAR_YW7:
                    xmlField = xmlSceneFields.find(fieldName)
                    if xmlField  is not None:
                        kwVarYw7[fieldName] = xmlField.text

            ywScnArcs = string_to_list(kwVarYw7.get('Field_SceneArcs', ''))
            for shortName in ywScnArcs:
                for acId in self.novel.arcs:
                    if self.novel.arcs[acId].shortName == shortName:
                        if prjScn.scType == 0:
                            arcScenes = self.novel.arcs[acId].scenes
                            if arcScenes is None:
                                arcScenes = [f"{SCENE_PREFIX}{xmlScene.find('ID').text}"]
                            else:
                                arcScenes.append(f"{SCENE_PREFIX}{xmlScene.find('ID').text}")
                            self.novel.arcs[acId].scenes = arcScenes
                        break

            ywScnAssocs = string_to_list(kwVarYw7.get('Field_SceneAssoc', ''))
            prjScn.arcPoints = [f'{ARC_POINT_PREFIX}{arcPoint}' for arcPoint in ywScnAssocs]

            scMode = kwVarYw7.get('Field_SceneMode', None)
            try:
                prjScn.scMode = int(scMode)
            except:
                prjScn.scMode = None

            if kwVarYw7.get('Field_CustomAR', None) is not None:
                prjScn.scPacing = 2
            elif xmlScene.find('ReactionScene') is not None:
                prjScn.scPacing = 1
            else:
                prjScn.scPacing = 0

            if xmlScene.find('Unused') is not None:
                if prjScn.scType == 0:
                    prjScn.scType = 3

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

            if xmlScene.find('Goal') is not None:
                prjScn.goal = xmlScene.find('Goal').text

            if xmlScene.find('Conflict') is not None:
                prjScn.conflict = xmlScene.find('Conflict').text

            if xmlScene.find('Outcome') is not None:
                prjScn.outcome = xmlScene.find('Outcome').text

            if xmlScene.find('ImageFile') is not None:
                prjScn.image = xmlScene.find('ImageFile').text

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
                ptId = f"{ARC_POINT_PREFIX}{ywScId}"
                self.novel.arcPoints[ptId] = ArcPoint(title=prjScn.title,
                                                      desc=prjScn.desc
                                                      )
                if ywScnAssocs:
                    self.novel.arcPoints[ptId].sceneAssoc = f'{SCENE_PREFIX}{ywScnAssocs[0]}'
            else:
                scId = f"{SCENE_PREFIX}{ywScId}"
                self.novel.scenes[scId] = prjScn

    def _remove_inline_code(self, text):
        if text:
            text = text.replace('<RTFBRK>', '')
            text = re.sub('\[\/*[h|c|r|s|u]\d*\]', '', text)
            YW_SPECIAL_CODES = ('HTM', 'TEX', 'RTF', 'epub', 'mobi', 'rtfimg')
            for specialCode in YW_SPECIAL_CODES:
                text = re.sub(f'\<{specialCode} .+?\/{specialCode}\>', '', text)
        else:
            text = ''
        return text.strip()

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

from html import unescape
from datetime import datetime
from xml import sax


class NovxPParser(sax.ContentHandler):

    def __init__(self):
        super().__init__()
        self.textList = None
        self._paragraph = None
        self._span = None

    def feed(self, xmlString):
        self.textList = []
        if xmlString:
            self._paragraph = False
            self._span = []
            sax.parseString(xmlString, self)

    def characters(self, content):
        if self._paragraph:
            self.textList.append(content)

    def endElement(self, name):
        if name in ('p', 'blockquote'):
            while self._span:
                self.textList.append(self._span.pop())
            self.textList.append('\n')
            self._paragraph = False
        elif name == 'em':
            self.textList.append('[/i]')
        elif name == 'strong':
            self.textList.append('[/b]')
        elif name == 'span':
            if self._span:
                self.textList.append(self._span.pop())
        elif name == 'comment':
            self.textList.append('*/')

    def startElement(self, name, attrs):
        xmlAttributes = {}
        for attribute in attrs.items():
            attrKey, attrValue = attribute
            xmlAttributes[attrKey] = attrValue
        locale = xmlAttributes.get('xml:lang', None)
        if name == 'p':
            self._paragraph = True
        elif name == 'em':
            self.textList.append('[i]')
        elif name == 'strong':
            self.textList.append('[b]')
        elif name == 'span':
            if locale is not None:
                self._span.append(f'[/lang={locale}]')
                self.textList.append(f'[lang={locale}]')
        elif name == 'blockquote':
            self.textList.append('> ')
            self._paragraph = True
        elif name == 'comment':
            self.textList.append('/*')



class NovxFile(File):
    DESCRIPTION = _('novelyst project')
    EXTENSION = '.novx'

    MAJOR_VERSION = 1
    MINOR_VERSION = 0

    XML_HEADER = '''<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE novx SYSTEM "novx_1_0.dtd">
<?xml-stylesheet href="novx.css" type="text/css"?>
'''

    def __init__(self, filePath, **kwargs):
        super().__init__(filePath)
        self.on_element_change = None
        self.xmlTree = None
        self.wcLog = {}

    def adjust_scene_types(self):
        partType = 0
        for chId in self.novel.tree.get_children(CH_ROOT):
            if self.novel.chapters[chId].chLevel == 1:
                partType = self.novel.chapters[chId].chType
            elif partType != 0 and not self.novel.chapters[chId].isTrash:
                self.novel.chapters[chId].chType = partType
            if self.novel.chapters[chId].chType != 0:
                for scId in self.novel.tree.get_children(chId):
                    self.novel.scenes[scId].scType = self.novel.chapters[chId].chType

    def read(self):
        try:
            self.xmlTree = ET.parse(self.filePath)
        except:
            raise Error(f'{_("Can not process file")}: "{norm_path(self.filePath)}".')

        xmlRoot = self.xmlTree.getroot()
        try:
            majorVersionStr, minorVersionStr = xmlRoot.attrib['version'].split('.')
            majorVersion = int(majorVersionStr)
            minorVersion = int(minorVersionStr)
        except:
            raise Error(f'{_("No valid version found in file")}: "{norm_path(self.filePath)}".')

        if majorVersion > self.MAJOR_VERSION:
            raise Error(_('The project "{}" was created with a newer novelyst version.').format(norm_path(self.filePath)))

        elif majorVersion < self.MAJOR_VERSION:
            raise Error(_('The project "{}" was created with an outdated novelyst version.').format(norm_path(self.filePath)))

        elif minorVersion > self.MINOR_VERSION:
            raise Error(_('The project "{}" was created with a newer novelyst version.').format(norm_path(self.filePath)))

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
        xmlChapters = xmlRoot.find('CHAPTERS')
        self._read_chapters(xmlChapters)
        xmlArcs = xmlRoot.find('ARCS')
        self._read_arcs(xmlArcs)
        self.adjust_scene_types()

        xmlWclog = xmlRoot.find('PROGRESS')
        if xmlWclog is not None:
            for xmlWc in xmlWclog.findall('WC'):
                wcDate = xmlWc.get('date')
                wcCount = xmlWc.get('count')
                wcTotalCount = xmlWc.get('withUnused')
                if wcDate and wcCount and wcTotalCount:
                    self.wcLog[wcDate] = [wcCount, wcTotalCount]

    def write(self):
        self.adjust_scene_types()
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

    def _build_arc_branch(self, xmlArcs, prjArc, acId):
        xmlArc = ET.SubElement(xmlArcs, 'ARC', attrib={'id':acId})
        if prjArc.title:
            ET.SubElement(xmlArc, 'Title').text = self._convert_from_novelyst(prjArc.title, quick=True)
        if prjArc.shortName:
            ET.SubElement(xmlArc, 'ShortName').text = self._convert_from_novelyst(prjArc.shortName, quick=True)
        if prjArc.desc:
            ET.SubElement(xmlArc, 'Desc').text = self._convert_from_novelyst(prjArc.desc)

        if prjArc.scenes:
            attrib = {'ids':' '.join(prjArc.scenes)}
            ET.SubElement(xmlArc, 'Sections', attrib=attrib)

        for apId in self.novel.tree.get_children(acId):
            xmlPoint = ET.SubElement(xmlArc, 'POINT', attrib={'id':apId})
            self._build_arcPoint_branch(xmlPoint, self.novel.arcPoints[apId])

        return xmlArc

    def _build_arcPoint_branch(self, xmlPoint, prjArcPoint):
        if prjArcPoint.title:
            ET.SubElement(xmlPoint, 'Title').text = self._convert_from_novelyst(prjArcPoint.title, quick=True)
        if prjArcPoint.desc:
            ET.SubElement(xmlPoint, 'Desc').text = self._convert_from_novelyst(prjArcPoint.desc)
        if prjArcPoint.notes:
            ET.SubElement(xmlPoint, 'Notes').text = self._convert_from_novelyst(prjArcPoint.notes)

        if prjArcPoint.sceneAssoc:
            ET.SubElement(xmlPoint, 'Section', attrib={'id': prjArcPoint.sceneAssoc})

    def _build_chapter_branch(self, xmlChapters, prjChp, chId):
        xmlChapter = ET.SubElement(xmlChapters, 'CHAPTER', attrib={'id':chId})
        xmlChapter.set('type', str(prjChp.chType))
        xmlChapter.set('level', str(prjChp.chLevel))
        if prjChp.isTrash:
            xmlChapter.set('isTrash', '1')
        if prjChp.noNumber:
            xmlChapter.set('noNumber', '1')
        if prjChp.title:
            ET.SubElement(xmlChapter, 'Title').text = self._convert_from_novelyst(prjChp.title, quick=True)
        if prjChp.desc:
            ET.SubElement(xmlChapter, 'Desc').text = self._convert_from_novelyst(prjChp.desc)
        for scId in self.novel.tree.get_children(chId):
            xmlScene = ET.SubElement(xmlChapter, 'SECTION', attrib={'id':scId})
            self._build_scene_branch(xmlScene, self.novel.scenes[scId])
        return xmlChapter

    def _build_character_branch(self, xmlCrt, prjCrt):
        if prjCrt.isMajor:
            xmlCrt.set('major', '1')
        if prjCrt.title:
            ET.SubElement(xmlCrt, 'Title').text = self._convert_from_novelyst(prjCrt.title, quick=True)
        if prjCrt.fullName:
            ET.SubElement(xmlCrt, 'FullName').text = self._convert_from_novelyst(prjCrt.fullName, quick=True)
        if prjCrt.aka:
            ET.SubElement(xmlCrt, 'Aka').text = self._convert_from_novelyst(prjCrt.aka, quick=True)
        if prjCrt.desc:
            ET.SubElement(xmlCrt, 'Desc').text = self._convert_from_novelyst(prjCrt.desc)
        if prjCrt.bio:
            ET.SubElement(xmlCrt, 'Bio').text = self._convert_from_novelyst(prjCrt.bio)
        if prjCrt.goals:
            ET.SubElement(xmlCrt, 'Goals').text = self._convert_from_novelyst(prjCrt.goals)
        if prjCrt.notes:
            ET.SubElement(xmlCrt, 'Notes').text = self._convert_from_novelyst(prjCrt.notes)
        tagStr = list_to_string(prjCrt.tags)
        if tagStr:
            ET.SubElement(xmlCrt, 'Tags').text = self._convert_from_novelyst(tagStr, quick=True)
        if prjCrt.image:
            ET.SubElement(xmlCrt, 'Link').text = self._convert_from_novelyst(prjCrt.image, quick=True)

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

        xmlArcs = ET.SubElement(root, 'ARCS')
        for acId in self.novel.tree.get_children(AC_ROOT):
            self._build_arc_branch(xmlArcs, self.novel.arcs[acId], acId)

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
                xmlWc.set('date', wc)
                xmlWc.set('count', self.wcLog[wc][0])
                xmlWc.set('withUnused', self.wcLog[wc][1])

    def _build_item_branch(self, xmlItm, prjItm):
        if prjItm.title:
            ET.SubElement(xmlItm, 'Title').text = self._convert_from_novelyst(prjItm.title, quick=True)
        if prjItm.aka:
            ET.SubElement(xmlItm, 'Aka').text = self._convert_from_novelyst(prjItm.aka, quick=True)
        if prjItm.desc:
            ET.SubElement(xmlItm, 'Desc').text = self._convert_from_novelyst(prjItm.desc)
        tagStr = list_to_string(prjItm.tags)
        if tagStr:
            ET.SubElement(xmlItm, 'Tags').text = self._convert_from_novelyst(tagStr, quick=True)
        if prjItm.image:
            ET.SubElement(xmlItm, 'Link').text = self._convert_from_novelyst(prjItm.image, quick=True)

    def _build_location_branch(self, xmlLoc, prjLoc):
        if prjLoc.title:
            ET.SubElement(xmlLoc, 'Title').text = self._convert_from_novelyst(prjLoc.title, quick=True)
        if prjLoc.aka:
            ET.SubElement(xmlLoc, 'Aka').text = self._convert_from_novelyst(prjLoc.aka, quick=True)
        if prjLoc.desc:
            ET.SubElement(xmlLoc, 'Desc').text = self._convert_from_novelyst(prjLoc.desc)
        tagStr = list_to_string(prjLoc.tags)
        if tagStr:
            ET.SubElement(xmlLoc, 'Tags').text = self._convert_from_novelyst(tagStr, quick=True)
        if prjLoc.image:
            ET.SubElement(xmlLoc, 'Link').text = self._convert_from_novelyst(prjLoc.image, quick=True)

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
        if self.novel.wordCountStart:
            xmlProject.set('wordCountStart', str(self.novel.wordCountStart))
        if self.novel.wordTarget:
            xmlProject.set('wordTarget', str(self.novel.wordTarget))

        if self.novel.title:
            ET.SubElement(xmlProject, 'Title').text = self._convert_from_novelyst(self.novel.title, quick=True)
        if self.novel.authorName:
            ET.SubElement(xmlProject, 'Author').text = self._convert_from_novelyst(self.novel.authorName, quick=True)
        if self.novel.desc:
            ET.SubElement(xmlProject, 'Desc').text = self._convert_from_novelyst(self.novel.desc)
        if self.novel.chapterHeadingPrefix:
            ET.SubElement(xmlProject, 'ChapterHeadingPrefix').text = self._convert_from_novelyst(self.novel.chapterHeadingPrefix, quick=True)
        if self.novel.chapterHeadingSuffix:
            ET.SubElement(xmlProject, 'ChapterHeadingSuffix').text = self._convert_from_novelyst(self.novel.chapterHeadingSuffix, quick=True)
        if self.novel.partHeadingPrefix:
            ET.SubElement(xmlProject, 'PartHeadingPrefix').text = self._convert_from_novelyst(self.novel.partHeadingPrefix, quick=True)
        if self.novel.partHeadingSuffix:
            ET.SubElement(xmlProject, 'PartHeadingSuffix').text = self._convert_from_novelyst(self.novel.partHeadingSuffix, quick=True)
        if self.novel.customGoal:
            ET.SubElement(xmlProject, 'CustomGoal').text = self._convert_from_novelyst(self.novel.customGoal, quick=True)
        if self.novel.customConflict:
            ET.SubElement(xmlProject, 'CustomConflict').text = self._convert_from_novelyst(self.novel.customConflict, quick=True)
        if self.novel.customOutcome:
            ET.SubElement(xmlProject, 'CustomOutcome').text = self._convert_from_novelyst(self.novel.customOutcome, quick=True)
        if self.novel.customChrBio:
            ET.SubElement(xmlProject, 'CustomChrBio').text = self._convert_from_novelyst(self.novel.customChrBio, quick=True)
        if self.novel.customChrGoals:
            ET.SubElement(xmlProject, 'CustomChrGoals').text = self._convert_from_novelyst(self.novel.customChrGoals, quick=True)

    def _build_scene_branch(self, xmlScene, prjScn):
        if prjScn.stageLevel is not None:
            xmlScene.set('stageLevel', str(prjScn.stageLevel))
            xmlScene.set('type', '3')
        else:
            xmlScene.set('type', str(prjScn.scType))
        if prjScn.status > 1:
            xmlScene.set('status', str(prjScn.status))
        if prjScn.scPacing > 0:
            xmlScene.set('pacing', str(prjScn.scPacing))
        if prjScn.appendToPrev:
            xmlScene.set('append', '1')
        if prjScn.scMode:
            xmlScene.set('mode', str(prjScn.scMode))
        if prjScn.title:
            ET.SubElement(xmlScene, 'Title').text = self._convert_from_novelyst(prjScn.title, quick=True)
        if prjScn.desc:
            ET.SubElement(xmlScene, 'Desc').text = self._convert_from_novelyst(prjScn.desc)
        if prjScn.goal:
            ET.SubElement(xmlScene, 'Goal').text = self._convert_from_novelyst(prjScn.goal)
        if prjScn.conflict:
            ET.SubElement(xmlScene, 'Conflict').text = self._convert_from_novelyst(prjScn.conflict)
        if prjScn.outcome:
            ET.SubElement(xmlScene, 'Outcome').text = self._convert_from_novelyst(prjScn.outcome)
        if prjScn.notes:
            ET.SubElement(xmlScene, 'Notes').text = self._convert_from_novelyst(prjScn.notes)
        tagStr = list_to_string(prjScn.tags)
        if tagStr:
            ET.SubElement(xmlScene, 'Tags').text = self._convert_from_novelyst(tagStr, quick=True)

        if (prjScn.date) and (prjScn.time):
            separator = ' '
            dateTime = f'{prjScn.date}{separator}{prjScn.time}'
            if dateTime != separator:
                if dateTime.count(':') < 2:
                    dateTime = f'{dateTime}:00'
                if dateTime:
                    ET.SubElement(xmlScene, 'SpecificDateTime').text = dateTime
        elif (prjScn.day is not None) or (prjScn.time is not None):
            if prjScn.day or prjScn.time:
                if prjScn.day:
                    ET.SubElement(xmlScene, 'Day').text = prjScn.day
                if prjScn.time:
                    hours, minutes, __ = prjScn.time.split(':')
                    if hours:
                        ET.SubElement(xmlScene, 'Hour').text = hours
                    if minutes:
                        ET.SubElement(xmlScene, 'Minute').text = minutes

        if prjScn.lastsDays:
            ET.SubElement(xmlScene, 'LastsDays').text = prjScn.lastsDays
        if prjScn.lastsHours:
            ET.SubElement(xmlScene, 'LastsHours').text = prjScn.lastsHours
        if prjScn.lastsMinutes:
            ET.SubElement(xmlScene, 'LastsMinutes').text = prjScn.lastsMinutes

        if prjScn.characters:
            attrib = {'ids':' '.join(prjScn.characters)}
            ET.SubElement(xmlScene, 'Characters', attrib=attrib)
        if prjScn.locations:
            attrib = {'ids':' '.join(prjScn.locations)}
            ET.SubElement(xmlScene, 'Locations', attrib=attrib)
        if prjScn.items:
            attrib = {'ids':' '.join(prjScn.items)}
            ET.SubElement(xmlScene, 'Items', attrib=attrib)

        if prjScn.sceneContent:
            ET.SubElement(xmlScene, 'Content').text = self._convert_from_novelyst(prjScn.sceneContent)

    def _convert_from_novelyst(self, text, quick=False):
        if not text:
            return ''

        nvReplacements = [
            ('&', '&amp;'),
            ('>', '&gt;'),
            ('<', '&lt;'),
            ("'", '&apos;'),
            ('"', '&quot;'),
            ]
        if not quick:
            text = text.strip()
            tags = ['i', 'b']
            nvReplacements.extend([
                ('\n', '</p><p>'),
                ('[i]', '<em>'),
                ('[/i]', '</em>'),
                ('[b]', '<strong>'),
                ('[/b]', '</strong>'),
                ('/*', '<comment>'),
                ('*/', '</comment>'),
            ])
            for language in self.novel.languages:
                tags.append(f'lang={language}')
                nvReplacements.append((f'[lang={language}]', f'<span xml:lang="{language}">'))
                nvReplacements.append((f'[/lang={language}]', '</span>'))

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

        for nv, nx in nvReplacements:
            text = text.replace(nv, nx)
        if quick:
            return text

        if text:
            return f'<p>{text}</p>'
        else:
            return ''

    def _convert_to_novelyst(self, text:str):
        nvText = ''
        if text:
            xmlBytes = ET.tostring(text, encoding='utf-8')
            xmlString = xmlBytes.decode('utf-8')
            parser = NovxPParser()
            parser.feed(xmlString)
            nvText = ''.join(parser.textList).strip()
        return nvText

    def _get_xml_text(self, parent, elemName, default=''):
        if parent.find(elemName) is not None:
            return parent.find(elemName).text
        else:
            return default

    def _postprocess_xml_file(self, filePath):
        with open(filePath, 'r', encoding='utf-8') as f:
            text = f.read()
        text = unescape(text)
        try:
            with open(filePath, 'w', encoding='utf-8') as f:
                f.write(f'{self.XML_HEADER}{text}')
        except:
            raise Error(f'{_("Cannot write file")}: "{norm_path(filePath)}".')

    def _read_arcs(self, xmlArcs):
        for xmlArc in xmlArcs.findall('ARC'):
            acId = xmlArc.attrib['id']
            self.novel.arcs[acId] = Arc(on_element_change=self.on_element_change)
            self.novel.arcs[acId].title = self._get_xml_text(xmlArc, 'Title')
            self.novel.arcs[acId].desc = self._convert_to_novelyst(xmlArc.find('Desc'))
            self.novel.arcs[acId].shortName = self._get_xml_text(xmlArc, 'ShortName')
            self.novel.tree.append(AC_ROOT, acId)
            for xmlPoint in xmlArc.findall('POINT'):
                apId = xmlPoint.attrib['id']
                self._read_arcPoint(xmlPoint, apId, acId)
                self.novel.tree.append(acId, apId)

            acSections = []
            xmlSections = xmlArc.find('Sections')
            if xmlSections is not None:
                scIds = xmlSections.get('ids', None)
                for scId in string_to_list(scIds, divider=' '):
                    if scId and scId in self.novel.scenes:
                        acSections.append(scId)
                        self.novel.scenes[scId].scArcs.append(acId)
                self.novel.arcs[acId].scenes = acSections

    def _read_arcPoint(self, xmlPoint, apId, acId):
        self.novel.arcPoints[apId] = ArcPoint(on_element_change=self.on_element_change)
        self.novel.arcPoints[apId].title = self._get_xml_text(xmlPoint, 'Title')
        self.novel.arcPoints[apId].desc = self._convert_to_novelyst(xmlPoint.find('Desc'))
        self.novel.arcPoints[apId].notes = self._convert_to_novelyst(xmlPoint.find('Notes'))
        xmlSceneAssoc = xmlPoint.find('Section')
        if xmlSceneAssoc is not None:
            scId = xmlSceneAssoc.get('id', None)
            self.novel.arcPoints[apId].sceneAssoc = scId
            self.novel.scenes[scId].scArcPoints[apId] = acId

    def _read_chapters(self, parent, partType=0):
        for xmlChapter in parent.findall('CHAPTER'):
            chId = xmlChapter.attrib['id']
            self.novel.chapters[chId] = Chapter(on_element_change=self.on_element_change)
            self.novel.chapters[chId].chType = int(xmlChapter.get('type', 0))
            self.novel.chapters[chId].chLevel = int(xmlChapter.get('level', 2))
            self.novel.chapters[chId].isTrash = xmlChapter.get('isTrash', None) == '1'
            self.novel.chapters[chId].noNumber = xmlChapter.get('noNumber', None) == '1'
            self.novel.chapters[chId].title = self._get_xml_text(xmlChapter, 'Title')
            self.novel.chapters[chId].desc = self._convert_to_novelyst(xmlChapter.find('Desc'))
            self.novel.tree.append(CH_ROOT, chId)
            if xmlChapter.find('SECTION'):
                for xmlScene in xmlChapter.findall('SECTION'):
                    scId = xmlScene.attrib['id']
                    self._read_scene(xmlScene, scId)
                    if self.novel.chapters[chId].chType > 0:
                        self.novel.scenes[scId].scType = self.novel.chapters[chId].chType
                    self.novel.tree.append(chId, scId)

    def _read_characters(self, root):
        for xmlCharacter in root.find('CHARACTERS'):
            crId = xmlCharacter.attrib['id']
            self.novel.characters[crId] = Character(on_element_change=self.on_element_change)
            self.novel.characters[crId].isMajor = xmlCharacter.get('major', None) == '1'
            self.novel.characters[crId].title = self._get_xml_text(xmlCharacter, 'Title')
            self.novel.characters[crId].image = self._get_xml_text(xmlCharacter, 'Link')
            self.novel.characters[crId].desc = self._convert_to_novelyst(xmlCharacter.find('Desc'))
            self.novel.characters[crId].aka = self._get_xml_text(xmlCharacter, 'Aka')
            tags = string_to_list(self._get_xml_text(xmlCharacter, 'Tags'))
            self.novel.characters[crId].tags = self._strip_spaces(tags)
            self.novel.characters[crId].notes = self._convert_to_novelyst(xmlCharacter.find('Notes'))
            self.novel.characters[crId].bio = self._convert_to_novelyst(xmlCharacter.find('Bio'))
            self.novel.characters[crId].goals = self._convert_to_novelyst(xmlCharacter.find('Goals'))
            self.novel.characters[crId].fullName = self._get_xml_text(xmlCharacter, 'FullName')
            self.novel.tree.append(CR_ROOT, crId)

    def _read_items(self, root):
        for xmlItem in root.find('ITEMS'):
            itId = xmlItem.attrib['id']
            self.novel.items[itId] = WorldElement(on_element_change=self.on_element_change)
            self.novel.items[itId].title = self._get_xml_text(xmlItem, 'Title')
            self.novel.items[itId].image = self._get_xml_text(xmlItem, 'Link')
            self.novel.items[itId].desc = self._convert_to_novelyst(xmlItem.find('Desc'))
            self.novel.items[itId].aka = self._get_xml_text(xmlItem, 'Aka')
            tags = string_to_list(self._get_xml_text(xmlItem, 'Tags'))
            self.novel.items[itId].tags = self._strip_spaces(tags)
            self.novel.tree.append(IT_ROOT, itId)

    def _read_locations(self, root):
        for xmlLocation in root.find('LOCATIONS'):
            lcId = xmlLocation.attrib['id']
            self.novel.locations[lcId] = WorldElement(on_element_change=self.on_element_change)
            self.novel.locations[lcId].title = self._get_xml_text(xmlLocation, 'Title')
            self.novel.locations[lcId].image = self._get_xml_text(xmlLocation, 'Link')
            self.novel.locations[lcId].desc = self._convert_to_novelyst(xmlLocation.find('Desc'))
            self.novel.locations[lcId].aka = self._get_xml_text(xmlLocation, 'Aka')
            tags = string_to_list(self._get_xml_text(xmlLocation, 'Tags'))
            self.novel.locations[lcId].tags = self._strip_spaces(tags)
            self.novel.tree.append(LC_ROOT, lcId)

    def _read_project(self, root):
        xmlProject = root.find('PROJECT')
        self.novel.renumberChapters = xmlProject.get('renumberChapters', None) == '1'
        self.novel.renumberParts = xmlProject.get('renumberParts', None) == '1'
        self.novel.renumberWithinParts = xmlProject.get('renumberWithinParts', None) == '1'
        self.novel.romanChapterNumbers = xmlProject.get('romanChapterNumbers', None) == '1'
        self.novel.romanPartNumbers = xmlProject.get('romanPartNumbers', None) == '1'
        self.novel.saveWordCount = xmlProject.get('saveWordCount', None) == '1'
        self.novel.wordCountStart = int(xmlProject.get('wordCountStart', 0))
        self.novel.wordTarget = int(xmlProject.get('wordTarget', 0))
        try:
            self.novel.workPhase = int(xmlProject.get('workPhase', None))
        except TypeError:
            self.novel.workPhase = None
        self.novel.title = self._get_xml_text(xmlProject, 'Title')
        self.novel.authorName = self._get_xml_text(xmlProject, 'Author')
        self.novel.desc = self._convert_to_novelyst(xmlProject.find('Desc'))
        self.novel.chapterHeadingPrefix = self._get_xml_text(xmlProject, 'ChapterHeadingPrefix')
        self.novel.chapterHeadingSuffix = self._get_xml_text(xmlProject, 'ChapterHeadingSuffix')
        self.novel.partHeadingPrefix = self._get_xml_text(xmlProject, 'PartHeadingPrefix')
        self.novel.partHeadingSuffix = self._get_xml_text(xmlProject, 'PartHeadingSuffix')
        self.novel.customGoal = self._get_xml_text(xmlProject, 'CustomGoal')
        self.novel.customConflict = self._get_xml_text(xmlProject, 'CustomConflict')
        self.novel.customOutcome = self._get_xml_text(xmlProject, 'CustomOutcome')
        self.novel.customChrBio = self._get_xml_text(xmlProject, 'CustomChrBio')
        self.novel.customChrGoals = self._get_xml_text(xmlProject, 'CustomChrGoals')

    def _read_scene(self, xmlScene, scId):
        self.novel.scenes[scId] = Scene(on_element_change=self.on_element_change)
        try:
            self.novel.scenes[scId].stageLevel = int(xmlScene.get('stageLevel', None))
        except TypeError:
            self.novel.scenes[scId].stageLevel = None
        if self.novel.scenes[scId].stageLevel is not None:
            self.novel.scenes[scId].scType = 3
        else:
            self.novel.scenes[scId].scType = int(xmlScene.get('type'), 0)
        self.novel.scenes[scId].status = int(xmlScene.get('status', 1))
        self.novel.scenes[scId].scPacing = int(xmlScene.get('pacing', 0))
        try:
            self.novel.scenes[scId].scMode = int(xmlScene.get('mode', None))
        except TypeError:
            self.novel.scenes[scId].scMode = None
        self.novel.scenes[scId].appendToPrev = xmlScene.get('append', None) == '1'
        self.novel.scenes[scId].title = self._get_xml_text(xmlScene, 'Title')
        self.novel.scenes[scId].desc = self._convert_to_novelyst(xmlScene.find('Desc'))
        self.novel.scenes[scId].sceneContent = self._convert_to_novelyst(xmlScene.find('Content'))
        self.novel.scenes[scId].notes = self._convert_to_novelyst(xmlScene.find('Notes'))
        tags = string_to_list(self._get_xml_text(xmlScene, 'Tags'))
        self.novel.scenes[scId].tags = self._strip_spaces(tags)

        self.novel.scenes[scId].date = ''
        self.novel.scenes[scId].time = ''
        self.novel.scenes[scId].day = ''
        self.novel.scenes[scId].time = ''
        if xmlScene.find('SpecificDateTime') is not None:
            dateTimeStr = xmlScene.find('SpecificDateTime').text

            try:
                dateTime = datetime.fromisoformat(dateTimeStr)
            except:
                pass
            else:
                startDateTime = dateTime.isoformat().split('T')
                self.novel.scenes[scId].date = startDateTime[0]
                self.novel.scenes[scId].time = startDateTime[1]
        else:
            if xmlScene.find('Day') is not None:
                day = xmlScene.find('Day').text

                try:
                    int(day)
                except ValueError:
                    pass
                else:
                    self.novel.scenes[scId].day = day

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
                self.novel.scenes[scId].time = f'{hour}:{minute}:00'

        self.novel.scenes[scId].lastsDays = self._get_xml_text(xmlScene, 'LastsDays')
        self.novel.scenes[scId].lastsHours = self._get_xml_text(xmlScene, 'LastsHours')
        self.novel.scenes[scId].lastsMinutes = self._get_xml_text(xmlScene, 'LastsMinutes')
        self.novel.scenes[scId].goal = self._convert_to_novelyst(xmlScene.find('Goal'))
        self.novel.scenes[scId].conflict = self._convert_to_novelyst(xmlScene.find('Conflict'))
        self.novel.scenes[scId].outcome = self._convert_to_novelyst(xmlScene.find('Outcome'))

        scCharacters = []
        xmlCharacters = xmlScene.find('Characters')
        if xmlCharacters is not None:
            crIds = xmlCharacters.get('ids', None)
            for crId in string_to_list(crIds, divider=' '):
                if crId and crId in self.novel.characters:
                    scCharacters.append(crId)
            self.novel.scenes[scId].characters = scCharacters

        scLocations = []
        xmlLocations = xmlScene.find('Locations')
        if xmlLocations is not None:
            lcIds = xmlLocations.get('ids', None)
            for lcId in string_to_list(lcIds, divider=' '):
                if lcId and lcId in self.novel.locations:
                    scLocations.append(lcId)
            self.novel.scenes[scId].locations = scLocations

        scItems = []
        xmlItems = xmlScene.find('Items')
        if xmlItems is not None:
            itIds = xmlItems.get('ids', None)
            for itId in string_to_list(itIds, divider=' '):
                if itId and itId in self.novel.items:
                    scItems.append(itId)
            self.novel.scenes[scId].items = scItems

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



class NvTree:

    def __init__(self):
        self.roots = {
            CH_ROOT:[],
            CR_ROOT:[],
            LC_ROOT:[],
            IT_ROOT:[],
            AC_ROOT:[],
            }
        self.srtScenes = {}
        self.srtArcPoints = {}

    def append(self, parent, iid):
        if parent in self.roots:
            self.roots[parent].append(iid)
            if parent == CH_ROOT:
                self.srtScenes[iid] = []
            elif parent == AC_ROOT:
                self.srtArcPoints[iid] = []
        elif parent.startswith(CHAPTER_PREFIX):
            try:
                self.srtScenes[parent].append(iid)
            except:
                self.srtScenes[parent] = [iid]
        elif parent.startswith(ARC_PREFIX):
            try:
                self.srtArcPoints[parent].append(iid)
            except:
                self.srtArcPoints[parent] = [iid]

    def delete_children(self, parent):
        if parent in self.roots:
            self.roots[parent] = []
            if parent == CH_ROOT:
                self.srtScenes = {}
            elif parent == AC_ROOT:
                self.srtArcPoints = {}
        elif parent.startswith(CHAPTER_PREFIX):
            self.srtScenes[parent] = []
        elif parent.startswith(ARC_PREFIX):
            self.srtArcPoints[parent] = []

    def get_children(self, item):
        if item in self.roots:
            return self.roots[item]

        elif item.startswith(CHAPTER_PREFIX):
            return self.srtScenes.get(item, [])

        elif item.startswith(ARC_PREFIX):
            return self.srtArcPoints.get(item, [])

    def insert(self, parent, index, iid):
        if parent in self.roots:
            self.roots[parent].insert(index, iid)
            if parent == CH_ROOT:
                self.srtScenes[iid] = []
            elif parent == AC_ROOT:
                self.srtArcPoints[iid] = []
        elif parent.startswith(CHAPTER_PREFIX):
            try:
                self.srtScenes[parent].insert(index, iid)
            except:
                self.srtScenes[parent] = [iid]
        elif parent.startswith(ARC_PREFIX):
            try:
                self.srtArcPoints.insert(index, iid)
            except:
                self.srtArcPoints[parent] = [iid]

    def reset(self):
        for item in self.roots:
            self.roots[item] = []
        self.srtScenes = {}
        self.srtArcPoints = {}

    def set_children(self, item, newchildren):
        if item in self.roots:
            self.roots[item] = newchildren[:]
            if item == CH_ROOT:
                self.srtScenes = {}
            elif item == AC_ROOT:
                self.srtArcPoints = {}
        elif item.startswith(CHAPTER_PREFIX):
            self.srtScenes[item] = newchildren[:]
        elif item.startswith(ARC_PREFIX):
            self.srtArcPoints[item] = newchildren[:]




def yw2novx(sourcePath):
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


if __name__ == '__main__':
    yw2novx(sys.argv[1])
    print('Done')
