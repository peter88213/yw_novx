"""Provide a converter class for universal export from a yWriter 7 project. 

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw2novx
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from novxlib.converter.yw_cnv_ff import YwCnvFf
from novxlib.yw.yw7_file import Yw7File
from yw2novxlib.novx_file import NovxFile


class Yw7Exporter(YwCnvFf):
    """A converter for universal export from a yWriter 7 project.

    Instantiate a Yw7File object as sourceFile and a
    Novel subclass object as targetFile for file conversion.

    Overrides the superclass constants EXPORT_SOURCE_CLASSES, EXPORT_TARGET_CLASSES.    
    """
    EXPORT_SOURCE_CLASSES = [Yw7File]
    EXPORT_TARGET_CLASSES = [NovxFile]
