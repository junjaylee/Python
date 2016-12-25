from distutils.core import setup
import py2exe
 
setup(
    options = {'py2exe': {
        'includes': ["pandas", "numpy"]
    }},
    console = [{'script': 'Data_Uploader.py'}],
    zipfile = None
)
#       'bundle_files': 1,
#'dll_excludes': ["MSVCP90.dll","HID.DLL", "w9xpopen.exe", "libzmq.pyd"]