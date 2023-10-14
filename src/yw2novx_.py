"""Convert yw7 to novx.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/yw2novx
License: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.en.html)
"""
SUFFIX = ''

import sys
import os

from novxlib.yw.yw7_reader import Yw7Reader
from novxlib.novx.novx_file import NovxFile
from novxlib.model.novel import Novel
from novxlib.model.nv_tree import NvTree


def main(sourcePath):
    path, extension = os.path.splitext(sourcePath)
    if extension != '.yw7':
        print(f'Error: File must be .yw7 type, but is "{extension}".')
        return

    targetPath = f'{path}.novx'
    source = Yw7Reader(sourcePath)
    target = NovxFile(targetPath)
    source.novel = Novel(tree=NvTree())
    source.read()
    target.novel = source.novel
    target.wcLog = source.wcLog
    target.write()
    print('Done')


if __name__ == '__main__':
    main(sys.argv[1])
