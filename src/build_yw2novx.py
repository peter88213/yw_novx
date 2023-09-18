""" Build a python script for the yw2novx distribution.
        
In order to distribute a single script without dependencies, 
this script "inlines" all modules imported from the pywriter package.

The pywriter project (see see https://github.com/peter88213/PyWriter)
must be located on the same directory level as the yw2novx project. 

For further information see https://github.com/peter88213/yw2novx
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
sys.path.append('../../PyWriter/src')
import inliner

SRC = '../src/'
BUILD = '../test/'
SOURCE_FILE = f'{SRC}yw2novx_.py'
TARGET_FILE = f'{BUILD}yw2novx.py'


def main():
    inliner.run(SOURCE_FILE, TARGET_FILE, 'pywriter', '../../novxlib/src/')
    print('Done.')


if __name__ == '__main__':
    main()
