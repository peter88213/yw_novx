#!/usr/bin/python3
"""novelyst collection upgrader

- Convert a novelyst .nvcx collection file to .nvcx format.
- Convert the collection's .yw7 project files to .novx format. 

Copyright (c) 2024 Peter Triesberger
For further information see https://github.com/peter88213/yw_novx
License: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.en.html)
"""
import sys
import os
import xml.etree.ElementTree as ET



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
#!/usr/bin/python3
SUFFIX = ''


from nvyw7lib.yw7_file import Yw7File


class BasicElement:

    def __init__(
        self,
        on_element_change=None,
        title=None,
        desc=None,
        links=None
    ):
        if on_element_change is None:
            self.on_element_change = self.do_nothing
        else:
            self.on_element_change = on_element_change
        self._title = title
        self._desc = desc
        if links is None:
            self._links = {}
        else:
            self._links = links
        self._fields = {}

    @property
    def title(self):
        return self._title

    @title.setter
    def title(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._title != newVal:
            self._title = newVal
            self.on_element_change()

    @property
    def desc(self):
        return self._desc

    @desc.setter
    def desc(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
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
                    assert type(val) is str
        if self._links != newVal:
            self._links = newVal
            self.on_element_change()

    @property
    def fields(self):
        return self._fields.copy()

    @fields.setter
    def fields(self, newVal):
        if self._fields != newVal:
            self._fields = newVal
            self.on_element_change()

    def do_nothing(self):
        pass




class BasicElementNotes(BasicElement):

    def __init__(
        self,
        notes=None,
        **kwargs
    ):
        super().__init__(**kwargs)
        self._notes = notes

    @property
    def notes(self):
        return self._notes

    @notes.setter
    def notes(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._notes != newVal:
            self._notes = newVal
            self.on_element_change()



class Chapter(BasicElementNotes):

    def __init__(
        self,
        chLevel=None,
        chType=None,
        noNumber=None,
        isTrash=None,
        hasEpigraph=None,
        **kwargs
    ):
        super().__init__(**kwargs)
        self._chLevel = chLevel
        self._chType = chType
        self._noNumber = noNumber
        self._isTrash = isTrash
        self._hasEpigraph = hasEpigraph

    @property
    def chLevel(self):
        return self._chLevel

    @chLevel.setter
    def chLevel(self, newVal):
        if newVal is not None:
            assert type(newVal) is int
        if self._chLevel != newVal:
            self._chLevel = newVal
            self.on_element_change()

    @property
    def chType(self):
        return self._chType

    @chType.setter
    def chType(self, newVal):
        if newVal is not None:
            assert type(newVal) is int
        if self._chType != newVal:
            self._chType = newVal
            self.on_element_change()

    @property
    def noNumber(self):
        return self._noNumber

    @noNumber.setter
    def noNumber(self, newVal):
        if newVal is not None:
            assert type(newVal) is bool
        if self._noNumber != newVal:
            self._noNumber = newVal
            self.on_element_change()

    @property
    def isTrash(self):
        return self._isTrash

    @isTrash.setter
    def isTrash(self, newVal):
        if newVal is not None:
            assert type(newVal) is bool
        if self._isTrash != newVal:
            self._isTrash = newVal
            self.on_element_change()

    @property
    def hasEpigraph(self):
        return self._hasEpigraph

    @hasEpigraph.setter
    def hasEpigraph(self, newVal):
        if newVal is not None:
            assert type(newVal) is bool
        if self._hasEpigraph != newVal:
            self._hasEpigraph = newVal
            self.on_element_change()



class BasicElementTags(BasicElementNotes):

    def __init__(
        self,
        tags=None,
        **kwargs
    ):
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
                    assert type(elem) is str
        if self._tags != newVal:
            self._tags = newVal
            self.on_element_change()



class WorldElement(BasicElementTags):

    def __init__(
        self,
        aka=None,
        **kwargs
    ):
        super().__init__(**kwargs)
        self._aka = aka

    @property
    def aka(self):
        return self._aka

    @aka.setter
    def aka(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._aka != newVal:
            self._aka = newVal
            self.on_element_change()



class Character(WorldElement):

    def __init__(
        self,
        bio=None,
        goals=None,
        fullName=None,
        isMajor=None,
        birthDate=None,
        deathDate=None,
        **kwargs
    ):
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
            assert type(newVal) is str
        if self._bio != newVal:
            self._bio = newVal
            self.on_element_change()

    @property
    def goals(self):
        return self._goals

    @goals.setter
    def goals(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._goals != newVal:
            self._goals = newVal
            self.on_element_change()

    @property
    def fullName(self):
        return self._fullName

    @fullName.setter
    def fullName(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._fullName != newVal:
            self._fullName = newVal
            self.on_element_change()

    @property
    def isMajor(self):
        return self._isMajor

    @isMajor.setter
    def isMajor(self, newVal):
        if newVal is not None:
            assert type(newVal) is bool
        if self._isMajor != newVal:
            self._isMajor = newVal
            self.on_element_change()

    @property
    def birthDate(self):
        return self._birthDate

    @birthDate.setter
    def birthDate(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._birthDate != newVal:
            self._birthDate = newVal
            self.on_element_change()

    @property
    def deathDate(self):
        return self._deathDate

    @deathDate.setter
    def deathDate(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._deathDate != newVal:
            self._deathDate = newVal
            self.on_element_change()

import locale
import re

from calendar import isleap, day_name, month_name
from datetime import date
from datetime import datetime
from datetime import time
from datetime import timedelta

import gettext

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
        t = gettext.translation(
            'novelibre',
            LOCALE_PATH,
            languages=[CURRENT_LANGUAGE],
        )
        _ = t.gettext
    except:

        def _(message):
            return message



class PyCalendar:

    DATE_FORMAT = _("YYYY-MM-DD")
    TIME_FORMAT = _("hh:mm")
    WEEKDAYS = day_name
    MONTHS = month_name
    min = date.min.isoformat()
    max = date.max.isoformat()

    @classmethod
    def age(cls, nowIso, birthDateIso, deathDateIso):
        now = datetime.fromisoformat(nowIso)
        if deathDateIso:
            deathDate = datetime.fromisoformat(deathDateIso)
            if now > deathDate:
                yearsDead = cls._difference_in_years(deathDate, now)
                daysDead = cls._difference_in_days(deathDate, now)
                if birthDateIso:
                    birthDate = datetime.fromisoformat(birthDateIso)
                    yearsOld = cls._difference_in_years(birthDate, deathDate)
                else:
                    yearsOld = None
                return yearsOld, yearsDead, None, daysDead

        if birthDateIso:
            birthDate = datetime.fromisoformat(birthDateIso)
            yearsOld = cls._difference_in_years(birthDate, now)
            daysOld = cls._difference_in_days(birthDate, now)
        return yearsOld, None, daysOld, None

    @classmethod
    def dt_disp(cls, day, dateStr, timeIso):
        dt = []
        if day:
            dt.append(f'{_("Day")} {day}')
        if dateStr:
            dt.append(dateStr)
        if timeIso:
            dt.append(cls.time_disp(timeIso))
        return ' '.join(dt)

    @classmethod
    def duration(cls, startDateIso, startTimeIso, endDateIso, endTimeIso):
        StartDateTime = datetime.fromisoformat(
            f'{startDateIso}T{startTimeIso}'
        )
        endDateTime = datetime.fromisoformat(f'{endDateIso}T{endTimeIso}')
        durationTimedelta = endDateTime - StartDateTime
        lastsHours = durationTimedelta.seconds // 3600
        lastsMinutes = (durationTimedelta.seconds % 3600) // 60
        if durationTimedelta.days:
            daysStr = str(durationTimedelta.days)
        else:
            daysStr = None
        if lastsHours:
            hoursStr = str(lastsHours)
        else:
            hoursStr = None
        if lastsMinutes:
            minutesStr = str(lastsMinutes)
        else:
            minutesStr = None
        return daysStr, hoursStr, minutesStr

    @classmethod
    def duration_disp(cls, lastsDays, lastsHours, lastsMinutes):
        duration = []
        if lastsDays and lastsDays != '0':
            duration.append(f"{lastsDays}{_('d')}")
        if lastsHours and lastsHours != '0':
            duration.append(f"{lastsHours}{_('h')}")
        if lastsMinutes and lastsMinutes != '0':
            duration.append(f"{lastsMinutes}{_('min')}")
        return ' '.join(duration)

    @classmethod
    def get_duration_str(cls, section):
        return cls.duration_disp(
            section.lastsDays,
            section.lastsHours,
            section.lastsMinutes
        )

    @classmethod
    def get_end_date_time(cls, section):
        sectionStart = datetime.fromisoformat(
            f'{section.date} {section.time}'
        )
        sectionEnd = sectionStart + cls._get_duration(section)
        return sectionEnd.isoformat().split('T')

    @classmethod
    def get_end_day_time(cls, section):
        if section.day:
            dayInt = int(section.day)
        else:
            dayInt = 0
        virtualStartDate = (date.min + timedelta(days=dayInt)).isoformat()
        virtualSectionStart = datetime.fromisoformat(
            f'{virtualStartDate} {section.time}'
        )
        virtualSectionEnd = virtualSectionStart + cls._get_duration(section)
        virtualEndDate, endTime = virtualSectionEnd.isoformat().split('T')
        endDay = str((date.fromisoformat(virtualEndDate) - date.min).days)
        return (endDay, endTime)

    @classmethod
    def get_end_time(cls, section):
        virtualSectionStart = datetime.fromisoformat(
            f'{cls.min} {section.time}'
        )
        virtualSectionEnd = virtualSectionStart + cls._get_duration(section)
        return virtualSectionEnd.isoformat().split('T')[1]

    @classmethod
    def get_locale_date(cls, isoDate, localize):
        if localize:
            try:
                localeDateStr = cls.locale_date(isoDate)
            except:
                localeDateStr = ''
            return localeDateStr

        else:
            return isoDate

    @classmethod
    def get_timestamp(cls, section, refIso):
        if not section.time and not section.date and not section.day:
            return

        timeStr = section.time
        if not timeStr:
            timeStr = '00:00'
        if section.date:
            try:
                sectionStart = datetime.fromisoformat(
                    f'{section.date} {timeStr}'
                )
            except:
                return
        else:
            try:
                if section.day:
                    dayInt = int(section.day)
                else:
                    dayInt = 0
                startDate = (
                    date.fromisoformat(refIso) + timedelta(days=dayInt)
                ).isoformat()
                sectionStart = datetime.fromisoformat(f'{startDate} {timeStr}')
            except:
                return

        return int((sectionStart - datetime.min).total_seconds())

    @classmethod
    def h_m_s_str(cls, timeIso):
        return timeIso.split(':')

    @classmethod
    def locale_date(cls, dateIso):
        return date.fromisoformat(dateIso).strftime('%x')

    @classmethod
    def specific_date(cls, dayStr, refIso):
        refDate = date.fromisoformat(refIso)
        return date.isoformat(refDate + timedelta(days=int(dayStr)))

    @classmethod
    def time_disp(cls, timeIso):
        h, m, __ = cls.verified_time(timeIso).split(':')
        return f'{h}:{m}'

    @classmethod
    def unspecific_date(cls, dateIso, refIso):
        refDate = date.fromisoformat(refIso)
        return str((date.fromisoformat(dateIso) - refDate).days)

    @classmethod
    def verified_date(cls, dateIso):
        if dateIso is not None:
            date.fromisoformat(dateIso)
        return dateIso

    @classmethod
    def verified_time(cls, timeIso):
        if  timeIso is not None:
            time.fromisoformat(timeIso)
            while timeIso.count(':') < 2:
                timeIso = f'{timeIso}:00'
        return timeIso

    @classmethod
    def weekday(cls, dateIso):
        return date.fromisoformat(dateIso).weekday()

    @classmethod
    def weekday_str(cls, timestamp):
        return (datetime.min + timedelta(seconds=timestamp)).strftime('%A')

    @classmethod
    def y_m_d_str(cls, dateIso):
        return dateIso.split('-')

    @classmethod
    def _difference_in_years(cls, startDate, endDate):
        diffyears = endDate.year - startDate.year
        difference = endDate - startDate.replace(endDate.year)
        days_in_year = isleap(endDate.year) and 366 or 365
        years = diffyears + (
            difference.days + difference.seconds / 86400.0
            ) / days_in_year
        return int(years)

    @classmethod
    def _difference_in_days(cls, startDate, endDate):
        return (endDate - startDate).days

    @classmethod
    def _get_duration(cls, section):
        if section.lastsDays:
            lastsDays = int(section.lastsDays)
        else:
            lastsDays = 0
        if section.lastsHours:
            lastsSeconds = int(section.lastsHours) * 3600
        else:
            lastsSeconds = 0
        if section.lastsMinutes:
            lastsSeconds += int(section.lastsMinutes) * 60
        return timedelta(days=lastsDays, seconds=lastsSeconds)


LANGUAGE_TAG = re.compile(r'\<(p|span|h.) xml\:lang=\"(.*?)\".*?\>')


class Novel(BasicElement):

    def __init__(
        self,
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
        noSceneField1=None,
        noSceneField2=None,
        noSceneField3=None,
        otherSceneField1=None,
        otherSceneField2=None,
        otherSceneField3=None,
        crField1=None,
        crField2=None,
        referenceDate=None,
        tree=None,
        **kwargs
    ):
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
        self._noSceneField1 = noSceneField1
        self._noSceneField2 = noSceneField2
        self._noSceneField3 = noSceneField3
        self._otherSceneField1 = otherSceneField1
        self._otherSceneField2 = otherSceneField2
        self._otherSceneField3 = otherSceneField3
        self._crField1 = crField1
        self._crField2 = crField2

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
            self.referenceWeekDay = PyCalendar.weekday(referenceDate)
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
            assert type(newVal) is str
        if self._authorName != newVal:
            self._authorName = newVal
            self.on_element_change()

    @property
    def wordTarget(self):
        return self._wordTarget

    @wordTarget.setter
    def wordTarget(self, newVal):
        if newVal is not None:
            assert type(newVal) is int
        if self._wordTarget != newVal:
            self._wordTarget = newVal
            self.on_element_change()

    @property
    def wordCountStart(self):
        return self._wordCountStart

    @wordCountStart.setter
    def wordCountStart(self, newVal):
        if newVal is not None:
            assert type(newVal) is int
        if self._wordCountStart != newVal:
            self._wordCountStart = newVal
            self.on_element_change()

    @property
    def languageCode(self):
        return self._languageCode

    @languageCode.setter
    def languageCode(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._languageCode != newVal:
            self._languageCode = newVal
            self.on_element_change()

    @property
    def countryCode(self):
        return self._countryCode

    @countryCode.setter
    def countryCode(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._countryCode != newVal:
            self._countryCode = newVal
            self.on_element_change()

    @property
    def renumberChapters(self):
        return self._renumberChapters

    @renumberChapters.setter
    def renumberChapters(self, newVal):
        if newVal is not None:
            assert type(newVal) is bool
        if self._renumberChapters != newVal:
            self._renumberChapters = newVal
            self.on_element_change()

    @property
    def renumberParts(self):
        return self._renumberParts

    @renumberParts.setter
    def renumberParts(self, newVal):
        if newVal is not None:
            assert type(newVal) is bool
        if self._renumberParts != newVal:
            self._renumberParts = newVal
            self.on_element_change()

    @property
    def renumberWithinParts(self):
        return self._renumberWithinParts

    @renumberWithinParts.setter
    def renumberWithinParts(self, newVal):
        if newVal is not None:
            assert type(newVal) is bool
        if self._renumberWithinParts != newVal:
            self._renumberWithinParts = newVal
            self.on_element_change()

    @property
    def romanChapterNumbers(self):
        return self._romanChapterNumbers

    @romanChapterNumbers.setter
    def romanChapterNumbers(self, newVal):
        if newVal is not None:
            assert type(newVal) is bool
        if self._romanChapterNumbers != newVal:
            self._romanChapterNumbers = newVal
            self.on_element_change()

    @property
    def romanPartNumbers(self):
        return self._romanPartNumbers

    @romanPartNumbers.setter
    def romanPartNumbers(self, newVal):
        if newVal is not None:
            assert type(newVal) is bool
        if self._romanPartNumbers != newVal:
            self._romanPartNumbers = newVal
            self.on_element_change()

    @property
    def saveWordCount(self):
        return self._saveWordCount

    @saveWordCount.setter
    def saveWordCount(self, newVal):
        if newVal is not None:
            assert type(newVal) is bool
        if self._saveWordCount != newVal:
            self._saveWordCount = newVal
            self.on_element_change()

    @property
    def workPhase(self):
        return self._workPhase

    @workPhase.setter
    def workPhase(self, newVal):
        if newVal is not None:
            assert type(newVal) is int
        if self._workPhase != newVal:
            self._workPhase = newVal
            self.on_element_change()

    @property
    def chapterHeadingPrefix(self):
        return self._chapterHeadingPrefix

    @chapterHeadingPrefix.setter
    def chapterHeadingPrefix(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._chapterHeadingPrefix != newVal:
            self._chapterHeadingPrefix = newVal
            self.on_element_change()

    @property
    def chapterHeadingSuffix(self):
        return self._chapterHeadingSuffix

    @chapterHeadingSuffix.setter
    def chapterHeadingSuffix(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._chapterHeadingSuffix != newVal:
            self._chapterHeadingSuffix = newVal
            self.on_element_change()

    @property
    def partHeadingPrefix(self):
        return self._partHeadingPrefix

    @partHeadingPrefix.setter
    def partHeadingPrefix(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._partHeadingPrefix != newVal:
            self._partHeadingPrefix = newVal
            self.on_element_change()

    @property
    def partHeadingSuffix(self):
        return self._partHeadingSuffix

    @partHeadingSuffix.setter
    def partHeadingSuffix(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._partHeadingSuffix != newVal:
            self._partHeadingSuffix = newVal
            self.on_element_change()

    @property
    def noSceneField1(self):
        return self._noSceneField1

    @noSceneField1.setter
    def noSceneField1(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._noSceneField1 != newVal:
            self._noSceneField1 = newVal
            self.on_element_change()

    @property
    def noSceneField2(self):
        return self._noSceneField2

    @noSceneField2.setter
    def noSceneField2(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._noSceneField2 != newVal:
            self._noSceneField2 = newVal
            self.on_element_change()

    @property
    def noSceneField3(self):
        return self._noSceneField3

    @noSceneField3.setter
    def noSceneField3(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._noSceneField3 != newVal:
            self._noSceneField3 = newVal
            self.on_element_change()

    @property
    def otherSceneField1(self):
        return self._otherSceneField1

    @otherSceneField1.setter
    def otherSceneField1(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._otherSceneField1 != newVal:
            self._otherSceneField1 = newVal
            self.on_element_change()

    @property
    def otherSceneField2(self):
        return self._otherSceneField2

    @otherSceneField2.setter
    def otherSceneField2(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._otherSceneField2 != newVal:
            self._otherSceneField2 = newVal
            self.on_element_change()

    @property
    def otherSceneField3(self):
        return self._otherSceneField3

    @otherSceneField3.setter
    def otherSceneField3(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._otherSceneField3 != newVal:
            self._otherSceneField3 = newVal
            self.on_element_change()

    @property
    def crField1(self):
        return self._crField1

    @crField1.setter
    def crField1(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._crField1 != newVal:
            self._crField1 = newVal
            self.on_element_change()

    @property
    def crField2(self):
        return self._crField2

    @crField2.setter
    def crField2(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._crField2 != newVal:
            self._crField2 = newVal
            self.on_element_change()

    @property
    def referenceDate(self):
        return self._referenceDate

    @referenceDate.setter
    def referenceDate(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._referenceDate != newVal:
            if not newVal:
                self._referenceDate = None
                self.referenceWeekDay = None
                self.on_element_change()
            else:
                try:
                    self.referenceWeekDay = PyCalendar.weekday(newVal)
                except:
                    pass
                else:
                    self._referenceDate = newVal
                    self.on_element_change()

    def check_locale(self):
        if not self._languageCode:
            try:
                sysLng, sysCtr = locale.getlocale()[0].split('_')
            except:
                sysLng, sysCtr = locale.getdefaultlocale()[0].split('_')
            self.languageCode = sysLng
            self.countryCode = sysCtr
            return

        if len(self._languageCode) != 2:
            self.languageCode = 'zxx'
            self.countryCode = None
            return

        if self._countryCode and len(self._countryCode) != 2:
            self.countryCode = None

    def get_languages(self):

        def languages(text):
            m = LANGUAGE_TAG.search(text)
            while m:
                text = text[m.span()[1]:]
                yield m.group(2)
                m = LANGUAGE_TAG.search(text)

        self.languages = []
        for scId in self.sections:
            text = self.sections[scId].sectionContent
            if text:
                for language in languages(text):
                    if not language in self.languages:
                        self.languages.append(language)

    def get_tags(self):
        tags = {}
        for elements in [
            self.sections,
            self.characters,
            self.locations,
            self.items,
        ]:
            for elemId in elements:
                for tag in elements[elemId].tags:
                    if not tag in tags:
                        tags[tag] = [elemId]
                    else:
                        tags[tag].append(elemId)
        return tags

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
CHAPTERLIST_SUFFIX = '_chapterlist_tmp'
CHAPTERS_SUFFIX = '_chapters_tmp'
CHARACTER_REPORT_SUFFIX = '_character_report'
CHARACTERS_SUFFIX = '_characters_tmp'
CHARLIST_SUFFIX = '_charlist_tmp'
DATA_SUFFIX = '_data'
ELEMENT_NOTES_SUFFIX = '_element_note_report',
FULL_MANUSCRIPT_SUFFIX = '_full_tmp'
GRID_SUFFIX = '_grid_tmp'
ITEM_REPORT_SUFFIX = '_item_report'
ITEMLIST_SUFFIX = '_itemlist_tmp'
ITEMS_SUFFIX = '_items_tmp'
LOCATION_REPORT_SUFFIX = '_location_report'
LOCATIONS_SUFFIX = '_locations_tmp'
LOCLIST_SUFFIX = '_loclist_tmp'
MAJOR_MARKER = _('Major Character')
MANUSCRIPT_SUFFIX = '_manuscript_tmp'
METADATA_TEXT_SUFFIX = '_metadata_text_tmp'
MINOR_MARKER = _('Minor Character')
PARTLIST_SUFFIX = '_partlist_tmp'
PARTS_SUFFIX = '_parts_tmp'
PLOTLIST_SUFFIX = '_plotlist'
PLOTLINES_SUFFIX = '_plotlines_tmp'
PROJECTNOTES_SUFFIX = '_projectnote_report'
PROOF_SUFFIX = '_proof_tmp'
SECTIONLIST_SUFFIX = '_sectionlist'
SECTIONS_SUFFIX = '_sections_tmp'
STAGES_SUFFIX = '_structure_tmp'
TIMETABLE_SUFFIX = '_tt_tmp'
XREF_SUFFIX = '_xref'

NO_SCENE_FIELD_1_DEFAULT = _('Plot progress')
NO_SCENE_FIELD_2_DEFAULT = _('Characterization')
NO_SCENE_FIELD_3_DEFAULT = _('World building')
OTHER_SCENE_FIELD_1_DEFAULT = _('Opening')
OTHER_SCENE_FIELD_2_DEFAULT = _('Peak emotional moment')
OTHER_SCENE_FIELD_3_DEFAULT = _('Ending')
CR_FIELD_1_DEFAULT = _('Bio')
CR_FIELD_2_DEFAULT = _('Goals')

STATUS = [
    None,
    _('Outline'),
    _('Draft'),
    _('1st Edit'),
    _('2nd Edit'),
    _('Done')
]

SCENE = ['-', 'A', 'R', 'x']


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


def verified_int_string(intStr):
    if intStr is not None:
        int(intStr)
    return intStr



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
        self.srtPlotPoints = {}

    def append(self, parent, iid):
        if parent in self.roots:
            self.roots[parent].append(iid)
            if parent == CH_ROOT:
                self.srtSections[iid] = []
            elif parent == PL_ROOT:
                self.srtPlotPoints[iid] = []
            return

        if parent.startswith(CHAPTER_PREFIX):
            if parent in self.srtSections:
                self.srtSections[parent].append(iid)
            else:
                self.srtSections[parent] = [iid]
            return

        if parent.startswith(PLOT_LINE_PREFIX):
            if parent in self.srtPlotPoints:
                self.srtPlotPoints[parent].append(iid)
            else:
                self.srtPlotPoints[parent] = [iid]

    def delete(self, *items):
        raise NotImplementedError

    def delete_children(self, parent):
        if parent in self.roots:
            self.roots[parent] = []
            if parent == CH_ROOT:
                self.srtSections.clear()
                return

            if parent == PL_ROOT:
                self.srtPlotPoints.clear()
            return

        if parent.startswith(CHAPTER_PREFIX):
            self.srtSections[parent] = []
            return

        if parent.startswith(PLOT_LINE_PREFIX):
            self.srtPlotPoints[parent] = []

    def get_children(self, item):
        if item in self.roots:
            return self.roots[item]

        if item.startswith(CHAPTER_PREFIX):
            return self.srtSections.get(item, [])

        if item.startswith(PLOT_LINE_PREFIX):
            return self.srtPlotPoints.get(item, [])

    def index(self, item):
        raise NotImplementedError

    def insert(self, parent, index, iid):
        if parent in self.roots:
            self.roots[parent].insert(index, iid)
            if parent == CH_ROOT:
                self.srtSections[iid] = []
            elif parent == PL_ROOT:
                self.srtPlotPoints[iid] = []
            return

        if parent.startswith(CHAPTER_PREFIX):
            if parent in self.srtSections:
                self.srtSections[parent].insert(index, iid)
            else:
                self.srtSections[parent] = [iid]
            return

        if parent.startswith(PLOT_LINE_PREFIX):
            if parent in self.srtPlotPoints:
                self.srtPlotPoints[parent].insert(index, iid)
            else:
                self.srtPlotPoints[parent] = [iid]

    def move(self, item, parent, index):
        raise NotImplementedError

    def next(self, item):
        raise NotImplementedError

    def parent(self, item):
        if item.startswith(PLOT_POINT_PREFIX):
            for plId, ppIds in self.srtPlotPoints.items():
                if item in ppIds:
                    return plId

        elif item.startswith(SECTION_PREFIX):
            for chId, scIds in self.srtSections.items():
                if item in scIds:
                    return chId

        elif item in self.roots:
            return ''

        else:
            for root in self.roots:
                if item in root:
                    return root

        raise KeyError

    def prev(self, item):
        raise NotImplementedError

    def reset(self):
        for item in self.roots:
            self.roots[item] = []
        self.srtSections.clear()
        self.srtPlotPoints.clear()

    def set_children(self, item, newchildren):
        if item in self.roots:
            self.roots[item] = newchildren[:]
            if item == CH_ROOT:
                self.srtSections.clear()
                return

            if item == PL_ROOT:
                self.srtPlotPoints.clear()
            return

        if item.startswith(CHAPTER_PREFIX):
            self.srtSections[item] = newchildren[:]
            return

        if item.startswith(PLOT_LINE_PREFIX):
            self.srtPlotPoints[item] = newchildren[:]



class PlotLine(BasicElementNotes):

    def __init__(
        self,
        shortName=None,
        sections=None,
        **kwargs
    ):
        super().__init__(**kwargs)

        self._shortName = shortName
        self._sections = sections

    @property
    def shortName(self):
        return self._shortName

    @shortName.setter
    def shortName(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
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
                    assert type(elem) is str
        if self._sections != newVal:
            self._sections = newVal
            self.on_element_change()



class PlotPoint(BasicElementNotes):

    def __init__(
        self,
        sectionAssoc=None,
        **kwargs
    ):
        super().__init__(**kwargs)

        self._sectionAssoc = sectionAssoc

    @property
    def sectionAssoc(self):
        return self._sectionAssoc

    @sectionAssoc.setter
    def sectionAssoc(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._sectionAssoc != newVal:
            self._sectionAssoc = newVal
            self.on_element_change()




class WordCounter:

    IGNORE_PATTERN = re.compile(
        r'\<note\>.*?\<\/note\>|\<comment\>.*?\<\/comment\>|\<.+?\>'
    )

    SEPARATOR_PATTERN = re.compile(r'—|–|\<\/p\>')

    def get_word_count(self, text):
        text = text.replace('\n', '')
        text = self.SEPARATOR_PATTERN.sub(' ', text)
        text = self.IGNORE_PATTERN.sub('', text)
        return len(text.split())


class Section(BasicElementTags):

    NULL_DATE = '0001-01-01'
    NULL_TIME = '00:00:00'

    wordCounter = WordCounter()

    def __init__(
        self,
        scType=None,
        scene=None,
        status=None,
        appendToPrev=None,
        viewpoint=None,
        goal=None,
        conflict=None,
        outcome=None,
        plotlineNotes=None,
        scDate=None,
        scTime=None,
        day=None,
        lastsMinutes=None,
        lastsHours=None,
        lastsDays=None,
        characters=None,
        locations=None,
        items=None,
        **kwargs
    ):
        super().__init__(**kwargs)
        self._sectionContent = None
        self.wordCount = 0
        self._hasComment = False

        self._scType = scType
        self._scene = scene
        self._status = status
        self._appendToPrev = appendToPrev
        self._goal = goal
        self._conflict = conflict
        self._outcome = outcome
        self._plotlineNotes = plotlineNotes
        try:
            self._weekDay = PyCalendar.weekday(scDate)
            self._localeDate = PyCalendar.locale_date(scDate)
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
        self._viewpoint = viewpoint
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
            assert type(text) is str
        if self._sectionContent != text:
            self._sectionContent = text
            if text is not None:
                self.wordCount = self.wordCounter.get_word_count(text)
                self._hasComment = '<comment>' in self._sectionContent
            else:
                self.wordCount = 0
                self._hasComment = False
            self.on_element_change()

    @property
    def hasComment(self):
        return self._hasComment

    @property
    def scType(self):
        return self._scType

    @scType.setter
    def scType(self, newVal):
        if newVal is not None:
            assert type(newVal) is int
        if self._scType != newVal:
            self._scType = newVal
            self.on_element_change()

    @property
    def scene(self):
        return self._scene

    @scene.setter
    def scene(self, newVal):
        if newVal is not None:
            assert type(newVal) is int
        if self._scene != newVal:
            self._scene = newVal
            self.on_element_change()

    @property
    def status(self):
        return self._status

    @status.setter
    def status(self, newVal):
        if newVal is not None:
            assert type(newVal) is int
        if self._status != newVal:
            self._status = newVal
            self.on_element_change()

    @property
    def appendToPrev(self):
        return self._appendToPrev

    @appendToPrev.setter
    def appendToPrev(self, newVal):
        if newVal is not None:
            assert type(newVal) is bool
        if self._appendToPrev != newVal:
            self._appendToPrev = newVal
            self.on_element_change()

    @property
    def goal(self):
        return self._goal

    @goal.setter
    def goal(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._goal != newVal:
            self._goal = newVal
            self.on_element_change()

    @property
    def conflict(self):
        return self._conflict

    @conflict.setter
    def conflict(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._conflict != newVal:
            self._conflict = newVal
            self.on_element_change()

    @property
    def outcome(self):
        return self._outcome

    @outcome.setter
    def outcome(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
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
                    assert type(val) is str
        if self._plotlineNotes != newVal:
            self._plotlineNotes = newVal
            self.on_element_change()

    @property
    def date(self):
        return self._date

    @date.setter
    def date(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._date != newVal:
            if not newVal:
                self._date = None
                self._weekDay = None
                self._localeDate = None
                self.on_element_change()
                return

            try:
                self._weekDay = PyCalendar.weekday(newVal)
            except:
                return

            try:
                self._localeDate = PyCalendar.locale_date(newVal)
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
            assert type(newVal) is str
        if self._time != newVal:
            self._time = newVal
            self.on_element_change()

    @property
    def day(self):
        return self._day

    @day.setter
    def day(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._day != newVal:
            self._day = newVal
            self.on_element_change()

    @property
    def lastsMinutes(self):
        return self._lastsMinutes

    @lastsMinutes.setter
    def lastsMinutes(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._lastsMinutes != newVal:
            self._lastsMinutes = newVal
            self.on_element_change()

    @property
    def lastsHours(self):
        return self._lastsHours

    @lastsHours.setter
    def lastsHours(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._lastsHours != newVal:
            self._lastsHours = newVal
            self.on_element_change()

    @property
    def lastsDays(self):
        return self._lastsDays

    @lastsDays.setter
    def lastsDays(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._lastsDays != newVal:
            self._lastsDays = newVal
            self.on_element_change()

    @property
    def viewpoint(self):
        return self._viewpoint

    @viewpoint.setter
    def viewpoint(self, newVal):
        if newVal is not None:
            assert type(newVal) is str
        if self._viewpoint != newVal:
            self._viewpoint = newVal
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
                    assert type(elem) is str
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
                    assert type(elem) is str
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
                    assert type(elem) is str
        if self._items != newVal:
            self._items = newVal
            self.on_element_change()

    def day_to_date(self, referenceDate):
        if self._date:
            return True

        try:
            self.date = PyCalendar.specific_date(self._day, referenceDate)
            self._day = None
            return True

        except:
            self.date = None
            return False

    def date_to_day(self, referenceDate):
        if self._day:
            return True

        try:
            self._day = PyCalendar.unspecific_date(self._date, referenceDate)
            self.date = None
            return True

        except:
            self._day = None
            return False

    def get_end_date_time(self):
        endDate = None
        endTime = None
        endDay = None
        if self.time:
            if self.date:
                try:
                    endDate, endTime = PyCalendar.get_end_date_time(self)
                except:
                    pass
            elif self.day:
                try:
                    endDay, endTime = PyCalendar.get_end_day_time(self)
                except:
                    pass
            else:
                endTime = PyCalendar.get_end_time(self)
        return endDate, endTime, endDay

from datetime import date

from abc import ABC
from urllib.parse import quote



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
            self.projectName = quote(
                tail.replace(f'{suffix}{self.EXTENSION}', '')
            )

    def is_locked(self):
        return False

    def read(self):
        raise NotImplementedError

    def write(self):
        raise NotImplementedError



class BasicElementNovx:

    def import_data(self, element, xmlElement):
        element.title = self._get_element_text(xmlElement, 'Title')
        element.desc = self._xml_element_to_text(xmlElement.find('Desc'))
        element.links = self._get_link_dict(xmlElement)
        element.fields = self._get_fields(xmlElement)

    def export_data(self, element, xmlElement):
        if element.title:
            ET.SubElement(xmlElement, 'Title').text = element.title
        if element.desc:
            xmlElement.append(self._text_to_xml_element('Desc', element.desc))
        for path in element.links:
            xmlLink = ET.SubElement(xmlElement, 'Link')
            ET.SubElement(xmlLink, 'Path').text = path
            if element.links[path]:
                ET.SubElement(xmlLink, 'FullPath').text = element.links[path]
        for tag in element.fields:
            xmlField = ET.SubElement(xmlElement, 'Field')
            xmlField.set('tag', tag)
            xmlField.text = element.fields[tag]

    def _get_element_text(self, xmlElement, tag, default=None):
        if xmlElement.find(tag) is not None:
            return xmlElement.find(tag).text
        else:
            return default

    def _get_fields(self, xmlElement):
        fields = {}
        for xmlField in xmlElement.iterfind('Field'):
            tag = xmlField.get('tag', None)
            if tag is not None:
                fields[tag] = xmlField.text
        return fields

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




class BasicElementNotesNovx(BasicElementNovx):

    def import_data(self, element, xmlElement):
        super().import_data(element, xmlElement)
        element.notes = self._xml_element_to_text(xmlElement.find('Notes'))

    def export_data(self, element, xmlElement):
        super().export_data(element, xmlElement)
        if element.notes:
            xmlElement.append(self._text_to_xml_element('Notes', element.notes))



class ChapterNovx(BasicElementNotesNovx):

    def import_data(self, element, xmlElement):
        super().import_data(element, xmlElement)
        typeStr = xmlElement.get('type', '0')
        if typeStr in ('0', '1'):
            element.chType = int(typeStr)
        else:
            element.chType = 1
        chLevel = xmlElement.get('level', '2')
        if chLevel in ('1', '2'):
            element.chLevel = int(chLevel)
        else:
            element.chLevel = 2
        element.isTrash = xmlElement.get('isTrash', None) == '1'
        element.noNumber = xmlElement.get('noNumber', None) == '1'
        element.hasEpigraph = xmlElement.get('hasEpigraph', None) == '1'

    def export_data(self, element, xmlElement):
        super().export_data(element, xmlElement)
        if element.chType:
            xmlElement.set('type', str(element.chType))
        if element.chLevel == 1:
            xmlElement.set('level', '1')
        if element.isTrash:
            xmlElement.set('isTrash', '1')
        if element.noNumber:
            xmlElement.set('noNumber', '1')
        if element.hasEpigraph:
            xmlElement.set('hasEpigraph', '1')


class BasicElementTagsNovx(BasicElementNotesNovx):

    def import_data(self, element, xmlElement):
        super().import_data(element, xmlElement)
        tags = string_to_list(self._get_element_text(xmlElement, 'Tags'))
        strippedTags = []
        for tag in tags:
            strippedTags.append(tag.strip())
        element.tags = strippedTags

    def export_data(self, element, xmlElement):
        super().export_data(element, xmlElement)
        tagStr = list_to_string(element.tags)
        if tagStr:
            ET.SubElement(xmlElement, 'Tags').text = tagStr



class WorldElementNovx(BasicElementTagsNovx):

    def import_data(self, element, xmlElement):
        super().import_data(element, xmlElement)
        element.aka = self._get_element_text(xmlElement, 'Aka')

    def export_data(self, element, xmlElement):
        super().export_data(element, xmlElement)
        if element.aka:
            ET.SubElement(xmlElement, 'Aka').text = element.aka



class CharacterNovx(WorldElementNovx):

    def import_data(self, element, xmlElement):
        super().import_data(element, xmlElement)
        element.isMajor = xmlElement.get('major', None) == '1'
        element.fullName = self._get_element_text(xmlElement, 'FullName')
        element.bio = self._xml_element_to_text(xmlElement.find('Bio'))
        element.goals = self._xml_element_to_text(xmlElement.find('Goals'))
        element.birthDate = PyCalendar.verified_date(
            self._get_element_text(xmlElement, 'BirthDate')
        )
        element.deathDate = PyCalendar.verified_date(
            self._get_element_text(xmlElement, 'DeathDate')
        )

    def export_data(self, element, xmlElement):
        super().export_data(element, xmlElement)
        if element.isMajor:
            xmlElement.set('major', '1')
        if element.fullName:
            ET.SubElement(xmlElement, 'FullName').text = element.fullName
        if element.bio:
            xmlElement.append(self._text_to_xml_element('Bio', element.bio))
        if element.goals:
            xmlElement.append(self._text_to_xml_element('Goals', element.goals))
        if element.birthDate:
            ET.SubElement(xmlElement, 'BirthDate').text = element.birthDate
        if element.deathDate:
            ET.SubElement(xmlElement, 'DeathDate').text = element.deathDate



class NovelNovx(BasicElementNovx):

    def import_data(self, element, xmlElement):
        super().import_data(element, xmlElement)
        element.renumberChapters = xmlElement.get(
            'renumberChapters', None) == '1'
        element.renumberParts = xmlElement.get(
            'renumberParts', None) == '1'
        element.renumberWithinParts = xmlElement.get(
            'renumberWithinParts', None) == '1'
        element.romanChapterNumbers = xmlElement.get(
            'romanChapterNumbers', None) == '1'
        element.romanPartNumbers = xmlElement.get(
            'romanPartNumbers', None) == '1'
        element.saveWordCount = xmlElement.get(
            'saveWordCount', None) == '1'
        workPhase = xmlElement.get('workPhase', None)
        if workPhase in ('1', '2', '3', '4', '5'):
            element.workPhase = int(workPhase)
        else:
            element.workPhase = None

        element.authorName = self._get_element_text(xmlElement, 'Author')

        element.chapterHeadingPrefix = self._get_element_text(
            xmlElement,
            'ChapterHeadingPrefix'
        )
        element.chapterHeadingSuffix = self._get_element_text(
            xmlElement,
            'ChapterHeadingSuffix'
        )

        element.partHeadingPrefix = self._get_element_text(
            xmlElement,
            'PartHeadingPrefix'
        )
        element.partHeadingSuffix = self._get_element_text(
            xmlElement,
            'PartHeadingSuffix'
        )

        element.noSceneField1 = self._get_element_text(
            xmlElement,
            'CustomPlotProgress',
            default=element.noSceneField1,
        )
        element.noSceneField2 = self._get_element_text(
            xmlElement,
            'CustomCharacterization',
            default=element.noSceneField2,
        )
        element.noSceneField3 = self._get_element_text(
            xmlElement,
            'CustomWorldBuilding',
            default=element.noSceneField3,
        )

        element.otherSceneField1 = self._get_element_text(
            xmlElement,
            'CustomGoal',
            default=element.otherSceneField1,
        )
        element.otherSceneField2 = self._get_element_text(
            xmlElement,
            'CustomConflict',
            default=element.otherSceneField2,
        )
        element.otherSceneField3 = self._get_element_text(
            xmlElement,
            'CustomOutcome',
            default=element.otherSceneField3,
        )

        element.crField1 = self._get_element_text(
            xmlElement,
            'CustomChrBio',
            default=element.crField1,
        )
        element.crField2 = self._get_element_text(
            xmlElement,
            'CustomChrGoals',
            default=element.crField2,
        )

        if xmlElement.find('WordCountStart') is not None:
            element.wordCountStart = int(
                xmlElement.find('WordCountStart').text
            )
        else:
            element.wordCountStart = 0
        if xmlElement.find('WordTarget') is not None:
            element.wordTarget = int(
                xmlElement.find('WordTarget').text
            )

        element.referenceDate = PyCalendar.verified_date(
            self._get_element_text(xmlElement, 'ReferenceDate')
        )

    def export_data(self, element, xmlElement):
        super().export_data(element, xmlElement)
        if element.renumberChapters:
            xmlElement.set('renumberChapters', '1')
        if element.renumberParts:
            xmlElement.set('renumberParts', '1')
        if element.renumberWithinParts:
            xmlElement.set('renumberWithinParts', '1')
        if element.romanChapterNumbers:
            xmlElement.set('romanChapterNumbers', '1')
        if element.romanPartNumbers:
            xmlElement.set('romanPartNumbers', '1')
        if element.saveWordCount:
            xmlElement.set('saveWordCount', '1')
        if element.workPhase is not None:
            xmlElement.set('workPhase', str(element.workPhase))

        if element.authorName:
            ET.SubElement(
                xmlElement,
                'Author',
            ).text = element.authorName

        if element.chapterHeadingPrefix:
            ET.SubElement(
                xmlElement,
                'ChapterHeadingPrefix',
            ).text = element.chapterHeadingPrefix
        if element.chapterHeadingSuffix:
            ET.SubElement(
                xmlElement,
                'ChapterHeadingSuffix',
            ).text = element.chapterHeadingSuffix

        if element.partHeadingPrefix:
            ET.SubElement(
                xmlElement,
                'PartHeadingPrefix',
            ).text = element.partHeadingPrefix
        if element.partHeadingSuffix:
            ET.SubElement(
                xmlElement,
                'PartHeadingSuffix',
            ).text = element.partHeadingSuffix

        if element.noSceneField1:
            ET.SubElement(
                xmlElement,
                'CustomPlotProgress',
            ).text = element.noSceneField1
        if element.noSceneField2:
            ET.SubElement(
                xmlElement,
                'CustomCharacterization',
            ).text = element.noSceneField2
        if element.noSceneField3:
            ET.SubElement(
                xmlElement,
                'CustomWorldBuilding',
            ).text = element.noSceneField3

        if element.otherSceneField1:
            ET.SubElement(
                xmlElement,
                'CustomGoal',
            ).text = element.otherSceneField1
        if element.otherSceneField2:
            ET.SubElement(
                xmlElement,
                'CustomConflict',
            ).text = element.otherSceneField2
        if element.otherSceneField3:
            ET.SubElement(
                xmlElement,
                'CustomOutcome',
            ).text = element.otherSceneField3

        if element.crField1:
            ET.SubElement(
                xmlElement,
                'CustomChrBio',
            ).text = element.crField1
        if element.crField2:
            ET.SubElement(
                xmlElement,
                'CustomChrGoals',
            ).text = element.crField2

        if element.wordCountStart:
            ET.SubElement(
                xmlElement,
                'WordCountStart',
            ).text = str(element.wordCountStart)
        if element.wordTarget:
            ET.SubElement(
                xmlElement,
                'WordTarget',
            ).text = str(element.wordTarget)

        if element.referenceDate:
            ET.SubElement(
                xmlElement,
                'ReferenceDate',
            ).text = element.referenceDate



def new_id(elements, prefix=''):
    i = 1
    while f'{prefix}{i}' in elements:
        i += 1
    return f'{prefix}{i}'



class NovxOpener:

    @classmethod
    def get_xml_root(cls, filePath, majorVersion, minorVersion):
        try:
            xmlTree = ET.parse(filePath)
        except Exception as ex:
            normPath = norm_path(filePath)
            raise RuntimeError(
                f'{_("Cannot process file")}: "{normPath}" - {str(ex)}'
            )

        xmlRoot = xmlTree.getroot()
        if xmlRoot.tag != 'novx':
            msg = _("No valid xml root element found in file")
            raise RuntimeError(f'{msg}: "{norm_path(filePath)}".')

        fileMajorVersion, fileMinorVersion = cls._get_file_version(
            xmlRoot,
            filePath,
        )
        fileMajorVersion, fileMinorVersion = cls._upgrade_file_version(
            xmlRoot,
            fileMajorVersion,
            fileMinorVersion,
        )
        cls._check_version(
            fileMajorVersion,
            fileMinorVersion,
            filePath,
            majorVersion,
            minorVersion,
        )
        return xmlRoot

    @classmethod
    def _check_version(
            cls,
            fileMajorVersion,
            fileMinorVersion,
            filePath,
            majorVersion,
            minorVersion,
    ):
        if fileMajorVersion > majorVersion:
            msg = _('The project "{}" was created with a newer novelibre version.')
            raise RuntimeError(msg.format(norm_path(filePath)))

        if fileMajorVersion < majorVersion:
            msg = _('The project "{}" was created with an outdated novelibre version.')
            raise RuntimeError(msg.format(norm_path(filePath)))

        if fileMinorVersion > minorVersion:
            msg = _('The project "{}" was created with a newer novelibre version.')
            raise RuntimeError(msg.format(norm_path(filePath)))

    @classmethod
    def _upgrade_file_version(
            cls,
            xmlRoot,
            fileMajorVersion,
            fileMinorVersion,
    ):
        if fileMajorVersion == 1 and fileMinorVersion < 7:
            cls._upgrade_to_1_7(xmlRoot)
            fileMinorVersion = 7
        if fileMajorVersion == 1 and fileMinorVersion < 8:
            cls._upgrade_to_1_8(xmlRoot)
            fileMinorVersion = 8
        return fileMajorVersion, fileMinorVersion

    @classmethod
    def _get_file_version(cls, xmlRoot, filePath):
        try:
            (
                fileMajorVersionStr,
                fileMinorVersionStr
            ) = xmlRoot.attrib['version'].split('.')
            fileMajorVersion = int(fileMajorVersionStr)
            fileMinorVersion = int(fileMinorVersionStr)
        except (KeyError, ValueError):
            msg = _("No valid version found in file")
            raise RuntimeError(msg.format(norm_path(filePath)))

        return fileMajorVersion, fileMinorVersion

    @classmethod
    def _upgrade_to_1_7(cls, xmlRoot):
        for xmlSection in xmlRoot.iter('SECTION'):
            xmlCharacters = xmlSection.find('Characters')
            if xmlCharacters is not None:
                crIds = xmlCharacters.get('ids', None)
                if crIds is not None:
                    crId = crIds.split(' ')[0]
                    ET.SubElement(
                        xmlSection,
                        'Viewpoint',
                        attrib={'id':crId},
                    )

    @classmethod
    def _upgrade_to_1_8(cls, xmlRoot):
        allSections = []

        for xmlSection in xmlRoot.iter(tag='SECTION'):
            allSections.append(xmlSection.attrib['id'])

        xmlChapters = xmlRoot.find('CHAPTERS')
        if xmlChapters is None:
            return

        for xmlChapter in xmlChapters.iterfind('CHAPTER'):
            xmlEpigraph = xmlChapter.find('Epigraph')
            xmlEpigraphSrc = xmlChapter.find('EpigraphSrc')
            if xmlEpigraph is not None:
                xmlChapter.remove(xmlEpigraph)

                xmlChapter.set('hasEpigraph', '1')

                xmlNewSection = ET.Element('SECTION')

                newId = new_id(allSections, SECTION_PREFIX)
                allSections.append(newId)
                xmlNewSection.set('id', newId)

                ET.SubElement(xmlNewSection, 'Title').text = _('Epigraph')

                xmlNewSection.append(xmlEpigraph)
                xmlEpigraph.tag = 'Content'

                if xmlEpigraphSrc is not None:
                    xmlChapter.remove(xmlEpigraphSrc)
                    xmlNewSection.append(
                        ET.fromstring(
                           f'<Desc><p>{xmlEpigraphSrc.text}</p></Desc>'
                        )
                    )

                xmlChapter.insert(0, xmlNewSection)



class PlotLineNovx(BasicElementNotesNovx):

    def import_data(self, element, xmlElement):
        super().import_data(element, xmlElement)
        element.shortName = self._get_element_text(xmlElement, 'ShortName')
        plSections = []
        xmlSections = xmlElement.find('Sections')
        if xmlSections is not None:
            scIds = xmlSections.get('ids', None)
            if scIds is not None:
                for scId in string_to_list(scIds, divider=' '):
                    plSections.append(scId)
        element.sections = plSections

    def export_data(self, element, xmlElement):
        super().export_data(element, xmlElement)
        if element.shortName:
            ET.SubElement(xmlElement, 'ShortName').text = element.shortName
        if element.sections:
            attrib = {'ids':' '.join(element.sections)}
            ET.SubElement(xmlElement, 'Sections', attrib=attrib)


class PlotPointNovx(BasicElementNotesNovx):

    def import_data(self, element, xmlElement):
        super().import_data(element, xmlElement)
        xmlSectionAssoc = xmlElement.find('Section')
        if xmlSectionAssoc is not None:
            element.sectionAssoc = xmlSectionAssoc.get('id', None)

    def export_data(self, element, xmlElement):
        super().export_data(element, xmlElement)
        if element.sectionAssoc:
            ET.SubElement(
                xmlElement,
                'Section',
                attrib={'id': element.sectionAssoc},
            )



class SectionNovx(BasicElementTagsNovx):

    def import_data(self, element, xmlElement):
        super().import_data(element, xmlElement)

        typeStr = xmlElement.get('type', '0')
        if typeStr in ('0', '1', '2', '3'):
            element.scType = int(typeStr)
        else:
            element.scType = 1
        status = xmlElement.get('status', '1')
        if status in ('1', '2', '3', '4', '5'):
            element.status = int(status)
        else:
            element.status = 1
        scene = xmlElement.get('scene', '0')
        if scene in ('0', '1', '2', '3'):
            element.scene = int(scene)
        else:
            element.scene = 0

        if not element.scene:
            sceneKind = xmlElement.get('pacing', None)
            if sceneKind in ('1', '2'):
                element.scene = int(sceneKind) + 1

        element.appendToPrev = xmlElement.get('append', None) == '1'

        xmlViewpoint = xmlElement.find('Viewpoint')
        if xmlViewpoint is not None:
            element.viewpoint = xmlViewpoint.get('id', None)

        element.goal = self._xml_element_to_text(xmlElement.find('Goal'))
        element.conflict = self._xml_element_to_text(xmlElement.find('Conflict'))
        element.outcome = self._xml_element_to_text(xmlElement.find('Outcome'))

        xmlPlotlineNotes = xmlElement.find('PlotNotes')
        if xmlPlotlineNotes is None:
            xmlPlotlineNotes = xmlElement
        plotlineNotes = {}
        for xmlPlotlineNote in xmlPlotlineNotes.iterfind('PlotlineNotes'):
            plId = xmlPlotlineNote.get('id', None)
            plotlineNotes[plId] = self._xml_element_to_text(xmlPlotlineNote)
        element.plotlineNotes = plotlineNotes

        if xmlElement.find('Date') is not None:
            element.date = PyCalendar.verified_date(xmlElement.find('Date').text)
        elif xmlElement.find('Day') is not None:
            element.day = verified_int_string(xmlElement.find('Day').text)

        if xmlElement.find('Time') is not None:
            element.time = PyCalendar.verified_time(xmlElement.find('Time').text)

        element.lastsDays = verified_int_string(
            self._get_element_text(xmlElement, 'LastsDays')
        )
        element.lastsHours = verified_int_string(
            self._get_element_text(xmlElement, 'LastsHours')
        )
        element.lastsMinutes = verified_int_string(
            self._get_element_text(xmlElement, 'LastsMinutes')
        )

        scCharacters = []
        xmlCharacters = xmlElement.find('Characters')
        if xmlCharacters is not None:
            crIds = xmlCharacters.get('ids', None)
            if crIds is not None:
                for crId in string_to_list(crIds, divider=' '):
                    scCharacters.append(crId)
        element.characters = scCharacters

        scLocations = []
        xmlLocations = xmlElement.find('Locations')
        if xmlLocations is not None:
            lcIds = xmlLocations.get('ids', None)
            if lcIds is not None:
                for lcId in string_to_list(lcIds, divider=' '):
                    scLocations.append(lcId)
        element.locations = scLocations

        scItems = []
        xmlItems = xmlElement.find('Items')
        if xmlItems is not None:
            itIds = xmlItems.get('ids', None)
            if itIds is not None:
                for itId in string_to_list(itIds, divider=' '):
                    scItems.append(itId)
        element.items = scItems

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
                element.sectionContent = xmlStr
            else:
                element.sectionContent = '<p></p>'
        elif element.scType < 2:
            element.sectionContent = '<p></p>'

    def export_data(self, element, xmlElement):
        super().export_data(element, xmlElement)
        if element.scType:
            xmlElement.set('type', str(element.scType))
        if element.status > 1:
            xmlElement.set('status', str(element.status))
        if element.scene > 0:
            xmlElement.set('scene', str(element.scene))
        if element.appendToPrev:
            xmlElement.set('append', '1')

        if element.viewpoint:
            ET.SubElement(
                xmlElement,
                'Viewpoint',
                attrib={'id':element.viewpoint},
            )

        if element.goal:
            xmlElement.append(
                self._text_to_xml_element('Goal', element.goal)
            )
        if element.conflict:
            xmlElement.append(
                self._text_to_xml_element('Conflict', element.conflict)
            )
        if element.outcome:
            xmlElement.append(
                self._text_to_xml_element('Outcome', element.outcome)
            )

        if element.plotlineNotes:
            for plId in element.plotlineNotes:
                if not plId in element.scPlotLines:
                    continue

                if not element.plotlineNotes[plId]:
                    continue

                xmlPlotlineNotes = self._text_to_xml_element(
                    'PlotlineNotes', element.plotlineNotes[plId]
                )
                xmlPlotlineNotes.set('id', plId)
                xmlElement.append(xmlPlotlineNotes)

        if element.date:
            ET.SubElement(xmlElement, 'Date').text = element.date
        elif element.day:
            ET.SubElement(xmlElement, 'Day').text = element.day
        if element.time:
            ET.SubElement(xmlElement, 'Time').text = element.time

        if element.lastsDays and element.lastsDays != '0':
            ET.SubElement(xmlElement, 'LastsDays').text = element.lastsDays
        if element.lastsHours and element.lastsHours != '0':
            ET.SubElement(xmlElement, 'LastsHours').text = element.lastsHours
        if element.lastsMinutes and element.lastsMinutes != '0':
            ET.SubElement(xmlElement, 'LastsMinutes').text = element.lastsMinutes

        if element.characters:
            ET.SubElement(
                xmlElement,
                'Characters',
                attrib={'ids':' '.join(element.characters)},
            )

        if element.locations:
            ET.SubElement(
                xmlElement,
                'Locations',
                attrib={'ids':' '.join(element.locations)},
            )

        if element.items:
            ET.SubElement(
                xmlElement,
                'Items',
                attrib={'ids':' '.join(element.items)},
            )

        sectionContent = element.sectionContent
        if sectionContent:
            if not sectionContent in ('<p></p>', '<p />'):
                xmlElement.append(
                    ET.fromstring(f'<Content>{sectionContent}</Content>')
                )


def strip_illegal_characters(text):
    return re.sub('[\x00-\x08|\x0b-\x0c|\x0e-\x1f]', '', text)



class NovxFile(File):
    DESCRIPTION = _('novelibre project')
    EXTENSION = '.novx'

    MAJOR_VERSION = 1
    MINOR_VERSION = 9

    XML_HEADER = (
        f'<?xml version="1.0" encoding="utf-8"?>\n'
        f'<!DOCTYPE novx SYSTEM "novx_{MAJOR_VERSION}_{MINOR_VERSION}.dtd">\n'
        '<?xml-stylesheet href="novx.css" type="text/css"?>\n'
    )

    fileOpener = NovxOpener

    def __init__(self, filePath, **kwargs):
        super().__init__(filePath)
        self.on_element_change = None
        self.xmlTree = None

        self.wcLog = {}

        self.wcLogUpdate = {}

        self.timestamp = None

        self.basicElementCnv = BasicElementNovx()
        self.chapterCnv = ChapterNovx()
        self.characterCnv = CharacterNovx()
        self.novelCnv = NovelNovx()
        self.plotLineCnv = PlotLineNovx()
        self.plotPointCnv = PlotPointNovx()
        self.sectionCnv = SectionNovx()
        self.worldElementCnv = WorldElementNovx()

    def adjust_section_types(self):
        partType = 0
        for chId in self.novel.tree.get_children(CH_ROOT):
            if self.novel.chapters[chId].chLevel == 1:
                partType = self.novel.chapters[chId].chType
            elif partType != 0 and not self.novel.chapters[chId].isTrash:
                self.novel.chapters[chId].chType = partType
            for scId in self.novel.tree.get_children(chId):
                if (self.novel.sections[scId].scType
                        < self.novel.chapters[chId].chType
                ):
                    self.novel.sections[scId].scType = (
                        self.novel.chapters[chId].chType
                    )

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

        xmlRoot = self.fileOpener.get_xml_root(
            self.filePath,
            self.MAJOR_VERSION,
            self.MINOR_VERSION,
        )
        try:
            locale = (
                xmlRoot.attrib['{http://www.w3.org/XML/1998/namespace}lang']
            )
        except KeyError:
            pass
        else:
            codes = locale.split('-')
            self.novel.languageCode = codes[0]
            try:
                self.novel.countryCode = codes[1]
            except IndexError:
                self.novel.countryCode = None
        self.novel.tree.reset()
        try:
            self._read_project_data(xmlRoot)
            self._read_locations(xmlRoot)
            self._read_items(xmlRoot)
            self._read_characters(xmlRoot)
            self._read_chapters_and_sections(xmlRoot)
            self._read_plot_lines_and_points(xmlRoot)
            self._read_project_notes(xmlRoot)
            self.adjust_section_types()
            self._read_word_count_log(xmlRoot)
        except Exception as ex:
            raise RuntimeError(f"{_('Corrupt project data')} ({str(ex)})")
        self._get_timestamp()
        self._keep_word_count()

    def write(self):
        self._update_word_count_log()
        self.adjust_section_types()
        self.novel.get_languages()

        if self.novel.countryCode:
            countryCode = f'-{self.novel.countryCode}'
        else:
            countryCode = ''
        attrib = {
            'version': f'{self.MAJOR_VERSION}.{self.MINOR_VERSION}',
            'xml:lang': f'{self.novel.languageCode}{countryCode}',
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
        self.novelCnv.export_data(self.novel, xmlProject)

    def _build_chapters_and_sections(self, root):
        xmlChapters = ET.SubElement(root, 'CHAPTERS')
        for chId in self.novel.tree.get_children(CH_ROOT):
            xmlChapter = ET.SubElement(
                xmlChapters, 'CHAPTER', attrib={'id': chId})
            self.chapterCnv.export_data(self.novel.chapters[chId], xmlChapter)
            for scId in self.novel.tree.get_children(chId):
                self.sectionCnv.export_data(
                    self.novel.sections[scId],
                    ET.SubElement(
                        xmlChapter,
                        'SECTION',
                        attrib={'id': scId},
                    )
                )

    def _build_characters(self, root):
        xmlCharacters = ET.SubElement(root, 'CHARACTERS')
        for crId in self.novel.tree.get_children(CR_ROOT):
            self.characterCnv.export_data(
                self.novel.characters[crId],
                ET.SubElement(
                    xmlCharacters,
                    'CHARACTER',
                    attrib={'id': crId},
                )
            )

    def _build_locations(self, root):
        xmlLocations = ET.SubElement(root, 'LOCATIONS')
        for lcId in self.novel.tree.get_children(LC_ROOT):
            self.worldElementCnv.export_data(
                self.novel.locations[lcId],
                ET.SubElement(
                    xmlLocations,
                    'LOCATION',
                    attrib={'id': lcId},
                )
            )

    def _build_items(self, root):
        xmlItems = ET.SubElement(root, 'ITEMS')
        for itId in self.novel.tree.get_children(IT_ROOT):
            self.worldElementCnv.export_data(
                self.novel.items[itId],
                ET.SubElement(
                    xmlItems,
                    'ITEM',
                    attrib={'id': itId},
                )
            )

    def _build_plot_lines_and_points(self, root):
        xmlPlotLines = ET.SubElement(root, 'ARCS')
        for plId in self.novel.tree.get_children(PL_ROOT):
            xmlPlotLine = ET.SubElement(
                xmlPlotLines,
                'ARC',
                attrib={'id': plId},
            )
            self.plotLineCnv.export_data(self.novel.plotLines[plId], xmlPlotLine)
            for ppId in self.novel.tree.get_children(plId):
                self.plotPointCnv.export_data(
                    self.novel.plotPoints[ppId],
                    ET.SubElement(
                        xmlPlotLine,
                        'POINT',
                        attrib={'id': ppId},
                    )
                )

    def _build_project_notes(self, root):
        xmlProjectNotes = ET.SubElement(root, 'PROJECTNOTES')
        for pnId in self.novel.tree.get_children(PN_ROOT):
            self.basicElementCnv.export_data(
                self.novel.projectNotes[pnId],
                ET.SubElement(
                    xmlProjectNotes,
                    'PROJECTNOTE',
                    attrib={'id': pnId},
                )
            )

    def _build_word_count_log(self, root):
        if not self.wcLog:
            return

        xmlWcLog = ET.SubElement(root, 'PROGRESS')
        wcLastCount = None
        wcLastTotalCount = None
        for wc in self.wcLog:
            wcCount, wcTotalCount = self.wcLog[wc]
            if self.novel.saveWordCount:
                if (
                    wcCount == wcLastCount
                    and wcTotalCount == wcLastTotalCount
                ):
                    continue

                wcLastCount = wcCount
                wcLastTotalCount = wcTotalCount
            xmlWc = ET.SubElement(xmlWcLog, 'WC')
            ET.SubElement(xmlWc, 'Date').text = wc
            ET.SubElement(xmlWc, 'Count').text = str(wcCount)
            ET.SubElement(xmlWc, 'WithUnused').text = str(wcTotalCount)

    def _check_id(self, elemId, elemPrefix):
        if not elemId.startswith(elemPrefix):
            raise RuntimeError(f"bad ID: '{elemId}'")

    def _get_timestamp(self):
        try:
            self.timestamp = os.path.getmtime(self.filePath)
        except Exception:
            self.timestamp = None

    def _keep_word_count(self):

        if not self.wcLog:
            return

        actualCount, actualTotalCount = self.count_words()
        latestDate = list(self.wcLog)[-1]
        latestCount = self.wcLog[latestDate][0]
        latestTotalCount = self.wcLog[latestDate][1]
        if (
            actualCount != latestCount
            or actualTotalCount != latestTotalCount
        ):
            try:
                fileDateIso = date.fromtimestamp(self.timestamp).isoformat()
            except Exception:
                fileDateIso = date.today().isoformat()
            self.wcLogUpdate[fileDateIso] = [actualCount, actualTotalCount]

    def _postprocess_xml_file(self, filePath):

        with open(filePath, 'r', encoding='utf-8') as f:
            text = f.read()
            text = strip_illegal_characters(text)
        try:
            with open(filePath, 'w', encoding='utf-8') as f:
                f.write(f'{self.XML_HEADER}{text}')
        except Exception as ex:
            msg = _("Cannot write file")
            msg = f'{msg}: "{norm_path(filePath)}"'
            msg = f'{msg} - {str(ex)}'
            raise RuntimeError(msg)

    def _read_chapters_and_sections(self, root):
        xmlChapters = root.find('CHAPTERS')
        if xmlChapters is None:
            return

        for xmlChapter in xmlChapters.iterfind('CHAPTER'):
            chId = xmlChapter.attrib['id']
            self._check_id(chId, CHAPTER_PREFIX)
            self.novel.chapters[chId] = Chapter(
                on_element_change=self.on_element_change)
            self.chapterCnv.import_data(self.novel.chapters[chId], xmlChapter)
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
            self.novel.characters[crId] = Character(
                on_element_change=self.on_element_change)
            self.characterCnv.import_data(
                self.novel.characters[crId],
                xmlCharacter
            )
            self.novel.tree.append(CR_ROOT, crId)

    def _read_items(self, root):
        xmlItems = root.find('ITEMS')
        if xmlItems is None:
            return

        for xmlItem in xmlItems.iterfind('ITEM'):
            itId = xmlItem.attrib['id']
            self._check_id(itId, ITEM_PREFIX)
            self.novel.items[itId] = WorldElement(
                on_element_change=self.on_element_change)
            self.worldElementCnv.import_data(self.novel.items[itId], xmlItem)
            self.novel.tree.append(IT_ROOT, itId)

    def _read_locations(self, root):
        xmlLocations = root.find('LOCATIONS')
        if xmlLocations is None:
            return

        for xmlLocation in xmlLocations.iterfind('LOCATION'):
            lcId = xmlLocation.attrib['id']
            self._check_id(lcId, LOCATION_PREFIX)
            self.novel.locations[lcId] = WorldElement(
                on_element_change=self.on_element_change)
            self.worldElementCnv.import_data(
                self.novel.locations[lcId],
                xmlLocation
            )
            self.novel.tree.append(LC_ROOT, lcId)

    def _read_plot_lines_and_points(self, root):
        xmlPlotLines = root.find('ARCS')
        if xmlPlotLines is None:
            return

        for xmlPlotLine in xmlPlotLines.iterfind('ARC'):
            plId = xmlPlotLine.attrib['id']
            self._check_id(plId, PLOT_LINE_PREFIX)
            self.novel.plotLines[plId] = PlotLine(
                on_element_change=self.on_element_change)
            self.plotLineCnv.import_data(self.novel.plotLines[plId], xmlPlotLine)
            self.novel.tree.append(PL_ROOT, plId)

            self.novel.plotLines[plId].sections = intersection(
                self.novel.plotLines[plId].sections, self.novel.sections)

            for scId in self.novel.plotLines[plId].sections:
                self.novel.sections[scId].scPlotLines.append(plId)

            for xmlPlotPoint in xmlPlotLine.iterfind('POINT'):
                ppId = xmlPlotPoint.attrib['id']
                self._check_id(ppId, PLOT_POINT_PREFIX)
                self._read_plot_point(xmlPlotPoint, ppId, plId)
                self.novel.tree.append(plId, ppId)

    def _read_plot_point(self, xmlPlotPoint, ppId, plId):
        self.novel.plotPoints[ppId] = PlotPoint(
            on_element_change=self.on_element_change)
        self.plotPointCnv.import_data(self.novel.plotPoints[ppId], xmlPlotPoint)

        scId = self.novel.plotPoints[ppId].sectionAssoc
        if scId in self.novel.sections:
            self.novel.sections[scId].scPlotPoints[ppId] = plId
        else:
            self.novel.plotPoints[ppId].sectionAssoc = None

    def _read_project_data(self, root):
        xmlProject = root.find('PROJECT')
        if xmlProject is None:
            return

        self.novelCnv.import_data(self.novel, xmlProject)

    def _read_project_notes(self, root):
        xmlProjectNotes = root.find('PROJECTNOTES')
        if xmlProjectNotes is None:
            return

        for xmlProjectNote in xmlProjectNotes.iterfind('PROJECTNOTE'):
            pnId = xmlProjectNote.attrib['id']
            self._check_id(pnId, PRJ_NOTE_PREFIX)
            self.novel.projectNotes[pnId] = BasicElement()
            self.basicElementCnv.import_data(
                self.novel.projectNotes[pnId],
                xmlProjectNote
            )
            self.novel.tree.append(PN_ROOT, pnId)

    def _read_section(self, xmlSection, scId):
        self.novel.sections[scId] = Section(
            on_element_change=self.on_element_change)
        self.sectionCnv.import_data(self.novel.sections[scId], xmlSection)

        self.novel.sections[scId].characters = intersection(
            self.novel.sections[scId].characters, self.novel.characters)
        self.novel.sections[scId].locations = intersection(
            self.novel.sections[scId].locations, self.novel.locations)
        self.novel.sections[scId].items = intersection(
            self.novel.sections[scId].items, self.novel.items)

    def _read_word_count_log(self, xmlRoot):

        def verified_date(dateStr):
            if dateStr is not None:
                date.fromisoformat(dateStr)
            return dateStr

        xmlWclog = xmlRoot.find('PROGRESS')
        if xmlWclog is None:
            return

        for xmlWc in xmlWclog.iterfind('WC'):
            try:
                wcDate = verified_date(xmlWc.find('Date').text)
                self.wcLog[wcDate] = [
                    int(xmlWc.find('Count').text),
                    int(xmlWc.find('WithUnused').text)
                ]
            except:
                pass

    def _update_word_count_log(self):

        if self.novel.saveWordCount:
            newCount, newTotalCount = self.count_words()
            todayIso = date.today().isoformat()
            self.wcLogUpdate[todayIso] = [newCount, newTotalCount]
            for wcDate in self.wcLogUpdate:
                self.wcLog[wcDate] = self.wcLogUpdate[wcDate]
        self.wcLogUpdate.clear()

    def _write_element_tree(self, xmlProject):

        backedUp = False
        if os.path.isfile(xmlProject.filePath):
            try:
                os.replace(xmlProject.filePath, f'{xmlProject.filePath}.bak')
            except Exception as ex:
                raise RuntimeError(str(ex))
            else:
                backedUp = True
        try:
            xmlProject.xmlTree.write(
                xmlProject.filePath, xml_declaration=False, encoding='utf-8')
        except Exception as ex:
            if backedUp:
                os.replace(f'{xmlProject.filePath}.bak', xmlProject.filePath)
            msg = _("Cannot write file")
            msg = f'{msg}: "{norm_path(xmlProject.filePath)}"'
            msg = f'{msg} - {str(ex)}'
            raise RuntimeError(msg)
import zipfile



class ZippedNovxOpener(NovxOpener):

    NOVX_EXTENSIONS = [
        '.novx',
    ]
    ZIP_EXTENSIONS = [
        '.zip',
    ]

    @classmethod
    def get_xml_root(cls, filePath, majorVersion, minorVersion):
        __, extension = os.path.splitext(filePath)
        try:
            if not extension in cls.ZIP_EXTENSIONS:
                raise RuntimeError('File type is not supported')

            with zipfile.ZipFile(filePath, 'r') as z:
                fileNames = z.namelist()
                xmlRoot = None
                for fileName in fileNames:
                    __, extension = os.path.splitext(fileName)
                    if extension in cls.NOVX_EXTENSIONS:
                        with z.open(fileName, 'r') as f:
                            xmlStr = f.read()
                        xmlRoot = ET.fromstring(xmlStr)
                        break

                if xmlRoot is None:
                    raise RuntimeError('File type is not supported')

        except Exception as ex:
            normPath = norm_path(filePath)
            raise RuntimeError(
                f'{_("Cannot process file")}: "{normPath}" - {str(ex)}'
            )

        if xmlRoot.tag != 'novx':
            msg = _("No valid xml root element found in file")
            raise RuntimeError(f'{msg}: "{norm_path(filePath)}".')

        fileMajorVersion, fileMinorVersion = cls._get_file_version(
            xmlRoot,
            filePath,
        )
        fileMajorVersion, fileMinorVersion = cls._upgrade_file_version(
            xmlRoot,
            fileMajorVersion,
            fileMinorVersion,
        )
        cls._check_version(
            fileMajorVersion,
            fileMinorVersion,
            filePath,
            majorVersion,
            minorVersion,
        )
        return xmlRoot



class ZippedNovxFile(NovxFile):

    DESCRIPTION = _('Zipped novelibre project')
    EXTENSION = '.zip'

    fileOpener = ZippedNovxOpener

    def write(self):
        raise NotImplementedError
from pathlib import Path

prefs = {}
launchers = {}

HOME_URL = 'https://github.com/peter88213/novelibre/'
NEWS_URL = 'https://github.com/peter88213/novelibre/discussions/1?sort=new'

HOME_DIR = str(Path.home()).replace('\\', '/')
INSTALL_DIR = f'{HOME_DIR}/.novx'
PROGRAM_DIR = os.path.dirname(sys.argv[0])
if not PROGRAM_DIR:
    PROGRAM_DIR = '.'
USER_STYLES_DIR = f'{INSTALL_DIR}/styles'
USER_STYLES_XML = f'{USER_STYLES_DIR}/styles.xml'

NOT_ASSIGNED = ''


def to_string(text):
    if text is None:
        return ''

    return str(text)



class NovxService:

    def change_word_counter(self, wordCounter):
        Section.wordCounter = wordCounter

    def get_novelibre_home_url(self):
        return HOME_URL

    def get_novx_file_extension(self):
        return NovxFile.EXTENSION

    def get_word_counter(self):
        return Section.wordCounter

    def get_zipped_novx_file_extension(self):
        return ZippedNovxFile.EXTENSION

    def new_basic_element(self, **kwargs):
        return BasicElement(**kwargs)

    def new_chapter(self, **kwargs):
        return Chapter(**kwargs)

    def new_character(self, **kwargs):
        return Character(**kwargs)

    def new_novel(self, **kwargs):
        kwargs['tree'] = kwargs.get('tree', NvTree())
        return Novel(**kwargs)

    def new_novx_file(self, filePath, **kwargs):
        return NovxFile(filePath, **kwargs)

    def new_nv_tree(self, **kwargs):
        return NvTree(**kwargs)

    def new_plot_line(self, **kwargs):
        return PlotLine(**kwargs)

    def new_plot_point(self, **kwargs):
        return PlotPoint(**kwargs)

    def new_section(self, **kwargs):
        return Section(**kwargs)

    def new_world_element(self, **kwargs):
        return WorldElement(**kwargs)

    def new_zipped_novx_file(self, filePath, **kwargs):
        return ZippedNovxFile(filePath, **kwargs)



def yw_novx(sourcePath):
    path, extension = os.path.splitext(sourcePath)
    if extension != '.yw7':
        raise ValueError(f'File must be .yw7 type, but is "{extension}".')

    nvService = NovxService()
    targetPath = f'{path}.novx'
    source = Yw7File(sourcePath, nv_service=nvService)
    target = nvService.new_novx_file(targetPath)
    source.novel = nvService.new_novel()
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

