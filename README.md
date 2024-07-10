# Convert *novelyst* projects and collections for *novelibre* 

The *yw_novx.py* and *upgrade_collection.py* scripts help with the migration from [novelyst](https://peter88213.github.io/novelyst/) to [novelibre](https://github.com/peter88213/novelibre).

## Requirements

- A Python installation (version 3.6 or newer).

# yw_novx

The *yw_novx.py* script converts a single *.yw7* file to the *.novx* format and vice versa.
It is a stand-alone alternative to the [nv_yw7](https://github.com/peter88213/nv_yw7) plugin, 
mainly intended for converting larger quantities of project files, e.g. by scripting.

## Download

Save the file [yw_novx.py](https://raw.githubusercontent.com/peter88213/yw_novx/main/build/yw_novx.py).

## Usage

Just drag the icon of your source file, and drop it on the *yw_novx.py* icon.

Alternatively, you can

- launch the program on the command line passing the source file as an argument, or
- launch the program via a batch file.

usage: `yw_novx.py Sourcefile`

# upgrade_collection

If you have got a book/series collection created with the *novelyst_collection* plugin,
you can convert it to the new *.nvcx* format for the 
[nv_collection](https://github.com/peter88213/nv_collection) plugin.
In addition, *.novx* files are automatically generated for all *.yw7* projects contained in the collection. 

## Download

Save the files [upgrade_collection.py](https://raw.githubusercontent.com/peter88213/yw_novx/main/build/upgrade_collection.py) 
and [yw_novx.py](https://raw.githubusercontent.com/peter88213/yw_novx/main/build/yw_novx.py). Both files must be
located in the same directory. 

## Usage

Just drag the icon of your *.pwc* collection file, and drop it on the *upgrade_collection.py* icon.
Alternatively, you can launch the program on the command line passing the source file as an argument.

usage: `upgrade_collection.py Sourcefile`

## License

This is Open Source software, and the *novelibre_collection* plugin is licenced under GPLv3. See the
[GNU General Public License website](https://www.gnu.org/licenses/gpl-3.0.en.html) for more
details, or consult the [LICENSE](https://github.com/peter88213/nv_collection/blob/main/LICENSE) file.






