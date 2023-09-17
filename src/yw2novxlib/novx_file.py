"""Provide a class for ODT chapters and scenes export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2novx
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from novxlib.novx_globals import *
from yw2novxlib.file_export import FileExport


class NovxFile(FileExport):
    """kalliope novel file representation.

    Convert the whole project.
    """
    EXTENSION = '.novx'
    DESCRIPTION = _('kalliope novel project')
    SUFFIX = ''

    _fileHeader = '''<?xml version="1.0" encoding="utf-8" ?>
<?xml-stylesheet href='novx.css' type='text/css' ?>
<!DOCTYPE novx SYSTEM "novx.dtd">
<novx version="0.1.0" xml:lang="$Language-$Country">
  <project>
    <author>
      <name>$AuthorName</name>
      <bio>$AuthorBio</bio>
    </author>
  </project>
  <book>
    <title>$Title</title>
    <desc>$Desc</desc>'''
    _fileFooter = '''
  </world>
</novx>'''

    _partTemplate = '''
      <part ID="pt$ID" type="0">
        <title>$Title</title>
        <desc>$Desc</desc>'''

    _partEndTemplate = '''
      </part>'''

    _notesPartTemplate = '''
      <part ID="ch$ID" type="1">
        <title>$Title</title>
        <desc>$Desc</desc>'''

    _todoPartTemplate = '''
      <part ID="ch$ID" type="2">
        <title>$Title</title>
        <desc>$Desc</desc>'''

    _chapterTemplate = '''
        <chapter ID="ch$ID" type="0">
          <title>$Title</title>
          <desc>$Desc</desc>'''

    _chapterEndTemplate = '''
        </chapter>'''

    _notesChapterTemplate = '''
        <chapter ID="ch$ID" type="1">
          <title>$Title</title>
          <desc>$Desc</desc>'''

    _notesChapterEndTemplate = '''
        </chapter>'''

    _todoChapterTemplate = '''
        <chapter ID="ch$ID" type="2">
          <title>$Title</title>
          <desc>$Desc</desc>'''

    _todoChapterEndTemplate = '''
        </chapter>'''

    _unusedChapterTemplate = '''
        <chapter ID="ch$ID" type="3">
          <title>$Title</title>
          <desc>$Desc</desc>'''

    _unusedChapterEndTemplate = '''
        </chapter>'''

    _sceneTemplate = '''
          <section ID="sc$ID" type="0">
            <title>$Title</title>
            <desc>$Desc</desc>
            <date>$Date</date>
            <time>$Time</time>
            <day>$Day</day>
            <hour>$Hour</hour>
            <minute>$Minute</minute>
            <lasts_days>$LastsDays</lasts_days>
            <lasts_hours>$LastsHours</lasts_hours>
            <lasts_minutes>$LastsMinutes</lasts_minutes>
            <content>$SceneContent</content>
          </section>'''

    _appendedSceneTemplate = '''
          <section ID="sc$ID" type="0" append="1">
            <title>$Title</title>
            <desc>$Desc</desc>
            <date>$Date</date>
            <time>$Time</time>
            <day>$Day</day>
            <hour>$Hour</hour>
            <minute>$Minute</minute>
            <lasts_days>$LastsDays</lasts_days>
            <lasts_hours>$LastsHours</lasts_hours>
            <lasts_minutes>$LastsMinutes</lasts_minutes>
            <content>$SceneContent</content>
          </section>'''

    _notesSceneTemplate = '''
          <section ID="sc$ID" type="1">
            <title>$Title</title>
            <desc>$Desc</desc>
            <date>$Date</date>
            <time>$Time</time>
            <day>$Day</day>
            <hour>$Hour</hour>
            <minute>$Minute</minute>
            <lasts_days>$LastsDays</lasts_days>
            <lasts_hours>$LastsHours</lasts_hours>
            <lasts_minutes>$LastsMinutes</lasts_minutes>
            <content>$SceneContent</content>
          </section>'''

    _todoSceneTemplate = '''
          <section ID="sc$ID" type="2">
            <title>$Title</title>
            <desc>$Desc</desc>
            <date>$Date</date>
            <time>$Time</time>
            <day>$Day</day>
            <hour>$Hour</hour>
            <minute>$Minute</minute>
            <lasts_days>$LastsDays</lasts_days>
            <lasts_hours>$LastsHours</lasts_hours>
            <lasts_minutes>$LastsMinutes</lasts_minutes>
            <content>$SceneContent</content>
          </section>'''

    _unusedSceneTemplate = '''
          <section ID="sc$ID" type="3">
            <title>$Title</title>
            <desc>$Desc</desc>
            <date>$Date</date>
            <time>$Time</time>
            <day>$Day</day>
            <hour>$Hour</hour>
            <minute>$Minute</minute>
            <lasts_days>$LastsDays</lasts_days>
            <lasts_hours>$LastsHours</lasts_hours>
            <lasts_minutes>$LastsMinutes</lasts_minutes>
            <content>$SceneContent</content>
          </section>'''

    _characterSectionHeading = '''
  </book>
  <world>'''

    _characterTemplate = '''
      <character ID="cr$ID" importance="$Status">
        <title>$Title</title>
        <desc>$Desc</desc>
      </character>'''

    _locationTemplate = '''
      <location ID="lc$ID">
        <title>$Title</title>
        <desc>$Desc</desc>
      </location>'''

    _itemTemplate = '''
      <item ID="it$ID">
        <title>$Title</title>
        <desc>$Desc</desc>
      </item>'''

    def _convert_from_yw(self, text, quick=False):
        """Return text, converted from yw7 markup to target format.
        
        Positional arguments:
            text -- string to convert.
        
        Optional arguments:
            quick -- bool: if True, apply a conversion mode for one-liners without formatting.
        
        Overrides the superclass method.
        """
        if text:
            tags = ['i', 'b']
            odtReplacements = [
                ('&', '&amp;'),
                ('>', '&gt;'),
                ('<', '&lt;'),
                ("'", '&apos;'),
                ('"', '&quot;'),
                ]
            if quick:
                for yw, od in odtReplacements:
                    text = text.replace(yw, od)
                return text

            odtReplacements.extend([
                ('\n', '</p>\r<p>'),
                ('\r', '\n'),
                ('[i]', '<em>'),
                ('[/i]', '</em>'),
                ('[b]', '<strong>'),
                ('[/b]', '</strong>'),
                ('/*', '<!--'),
                ('*/', '-->'),
            ])
            for i, language in enumerate(self.languages, 1):
                tags.append(f'lang={language}')
                odtReplacements.append((f'[lang={language}]', f'<text:span text:style-name="T{i}">'))
                odtReplacements.append((f'[/lang={language}]', '</text:span>'))

            #--- Process markup reaching across linebreaks.
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

            #--- Apply odt formating.
            for yw, od in odtReplacements:
                text = text.replace(yw, od)

            text = re.sub('\[\/lang=(.+?)\]', '</span>', text)
            text = re.sub('\[lang=(.+?)\]', '<span xml:lang="\\1">', text)

            # Remove highlighting, alignment,
            # strikethrough, and underline tags.
            text = re.sub('\[\/*[h|c|r|s|u]\d*\]', '', text)
            text = f'<p>{text}</p>'
        else:
            text = ''
        return text

    def write(self):
        """Determine the languages used in the document before writing.
        
        Extends the superclass method.
        """
        if self.languages is None:
            self.get_languages()
        super().write()

