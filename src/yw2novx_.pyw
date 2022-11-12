"""Convert yWriter chapters and scenes to odt format.

This is a PyWriter sample application.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2novx
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
SUFFIX = ''

import sys

from pywriter.ui.ui_cmd import UiCmd
from pywriter.ui.ui_tk import UiTk
from yw2novxlib.yw7_exporter import Yw7Exporter


def run(sourcePath, suffix=''):
    ui = UiTk('yWriter import/export')
    converter = Yw7Exporter()
    converter.ui = ui
    kwargs = {'suffix': suffix}
    converter.run(sourcePath, **kwargs)
    ui.start()


if __name__ == '__main__':
    run(sys.argv[1], SUFFIX)
