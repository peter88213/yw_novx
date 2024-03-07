""" Build a python script for the yw_novx distribution.
        
In order to distribute a single script without dependencies, 
this script "inlines" all modules imported from the novxlib package.

The novxlib project (see see https://github.com/peter88213/novxlib)
must be located on the same directory level as the yw_novx project. 

For further information see https://github.com/peter88213/yw_novx
License: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.en.html)
"""
import sys
sys.path.append('../../novxlib/src')
sys.path.append('../../nv_yw7/src')
import inliner

SRC = '../src/'
BUILD = '../build/'
SOURCE_FILE = f'{SRC}yw_novx_.py'
TARGET_FILE = f'{BUILD}yw_novx.py'


def main():
    inliner.run(SOURCE_FILE, TARGET_FILE, 'nvywlib', '../../nv_yw7/src/')
    inliner.run(TARGET_FILE, TARGET_FILE, 'novxlib', '../../novxlib/src/')
    print('Done.')


if __name__ == '__main__':
    main()
