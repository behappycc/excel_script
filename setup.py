from distutils.core import setup
import py2exe

# py -3 setup.py py2exe

setup(
    options = {'py2exe': {
        'bundle_files': 2,
    }},
    console = [{'script': 'script.py',
    "icon_resource":[(1, "icon.jpg")]}],
    zipfile = None
    )
