""" Build a python script for the yw_novx distribution.
        
In order to distribute a single script without dependencies, 
this script "inlines" all modules imported from the novxlib package.

The novxlib project (see see https://github.com/peter88213/novxlib)
must be located on the same directory level as the yw_novx project. 

For further information see https://github.com/peter88213/yw_novx
License: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.en.html)
"""
import os
import sys

import inliner

sys.path.insert(0, f'{os.getcwd()}/../../novxlib/src')

SOURCE_DIR = '../src/'
TEST_DIR = '../build/'
SOURCE_FILE = f'{SOURCE_DIR}upgrade_collection_.py'
TEST_FILE = f'{TEST_DIR}upgrade_collection.py'
NVLIB = 'nvlib'
NV_PATH = '../../novelibre/src/'
NOVXLIB = 'novxlib'
NOVX_PATH = '../../novxlib/src/'


def inline_modules():
    inliner.run(SOURCE_FILE, TEST_FILE, 'yw_novx_', './')
    inliner.run(TEST_FILE, TEST_FILE, NVLIB, NV_PATH)
    inliner.run(TEST_FILE, TEST_FILE, NOVXLIB, NOVX_PATH)
    print('Done.')


def main():
    os.makedirs(TEST_DIR, exist_ok=True)
    inline_modules()


if __name__ == '__main__':
    main()
