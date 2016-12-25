import sys
import distutils
from cx_Freeze import setup, Executable

build_exe_options = {
    'packages': ['os'],
    'excludes': ['tkinter'], # tkinterは使わないので除外する
    'includes': ["pandas", "numpy"],
    'include_files': ['Data_Uploader.plist'],
    'include_msvcr': True, # PySideを使うのでMicrosoftのCランタイムを含めないと起動できない
    'compressed'   : True
}
name = 'Data_Uploader'
programfiles_dir = 'ProgramFiles64Folder' if distutils.util.get_platform() == 'win-amd64' else 'ProgramFilesFolder'
bdist_msi_options = {
    'add_to_path': False,
    'initial_target_dir': '[%s]\%s' % (programfiles_dir, name)
}
options = {
    'build_exe': build_exe_options,
    'bdist_msi': bdist_msi_options
}
base = 'Win32GUI' if sys.platform == 'win32' else None
mainexe = Executable(
    'Data_Uploader.py',
    targetName = 'Data_Uploader.exe',
    base = base,
    copyDependentFiles = True
)
setup(
    name=name,
    version='1.1',
    description='Upload PDCA and Matrix data to MongoDb',
    options=options,
    executables=[mainexe]
)
#python Build_cx_Freeze.py bdist_msi