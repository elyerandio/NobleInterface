from distutils.core import setup
import py2exe

setup(windows=['NobleInterface.py'],
		options = {"py2exe": {'includes':'decimal'}})
