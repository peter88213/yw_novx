#!/usr/bin/python3
"""Convert yw7 to novx.

Copyright (c) 2024 Peter Triesberger
For further information see https://github.com/peter88213/yw_novx
License: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.en.html)
"""
SUFFIX = ''

import sys
import os

from nvyw7lib.yw7_file import Yw7File
from nvlib.novx_service import NovxService


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


if __name__ == '__main__':
    yw_novx(sys.argv[1])
    print('Done')

