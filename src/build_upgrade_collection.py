""" Build a python script for the yw_novx distribution.
        
In order to distribute a single script without dependencies, 
this script "inlines" all modules imported from the novxlib package.

The novxlib project (see see https://github.com/peter88213/novxlib)
must be located on the same directory level as the yw_novx project. 

For further information see https://github.com/peter88213/yw_novx
License: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.en.html)
"""
import sys
sys.path.append('../../novxlib-Alpha/src')
import inliner

SRC = '../src/'
BUILD = '../build/'
SOURCE_FILE = f'{SRC}upgrade_collection_.py'
TARGET_FILE = f'{BUILD}upgrade_collection.py'


def main():
    inliner.run(SOURCE_FILE, TARGET_FILE, 'yw_novx_', './')
    inliner.run(TARGET_FILE, TARGET_FILE, 'novxlib', '../../novxlib-Alpha/src/')
    print('Done.')


if __name__ == '__main__':
    main()
