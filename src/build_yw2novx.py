""" Build a python script for the yw2novx distribution.
        
In order to distribute a single script without dependencies, 
this script "inlines" all modules imported from the novxlib package.

The novxlib project (see see https://github.com/peter88213/novxlib)
must be located on the same directory level as the yw2novx project. 

For further information see https://github.com/peter88213/yw2novx
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import inliner

SRC = '../src/'
BUILD = '../test/'
SOURCE_FILE = f'{SRC}yw2novx_.pyw'
TARGET_FILE = f'{BUILD}yw2novx.pyw'


def main():
    inliner.run(SOURCE_FILE, TARGET_FILE, 'yw2novxlib', '../src/', copynovxlib=False)
    inliner.run(TARGET_FILE, TARGET_FILE, 'novxlib', '../../novxlib/src/', copynovxlib=False)
    # inliner.run(SOURCE_FILE, TARGET_FILE, 'yw2novxlib', '../src/')
    # inliner.run(TARGET_FILE, TARGET_FILE, 'novxlib', '../src/')
    print('Done.')


if __name__ == '__main__':
    main()
