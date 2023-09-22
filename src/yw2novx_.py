"""Convert yw7 to novx.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/yw2novx
License: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.en.html)
"""
SUFFIX = ''

import sys
import os

from pywriter.yw.yw7_reader import Yw7Reader
from pywriter.novx.novx_file import NovxFile
from pywriter.model.novel import Novel


def main(sourcePath):
    path, extension = os.path.splitext(sourcePath)
    if extension != '.yw7':
        print(f'Error: File must be .yw7 type, but is "{extension}".')
        return

    targetPath = f'{path}.novx'
    source = Yw7Reader(sourcePath)
    target = NovxFile(targetPath)
    source.novel = Novel()
    source.read()
    target.novel = source.novel
    target.write()
    print('Done')


if __name__ == '__main__':
    main(sys.argv[1])
