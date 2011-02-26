# -*- coding: utf-8 -*-
# Copyright (c) 2011 Wensheng Wang
#
# This Software is released under the MIT License:
# http://www.opensource.org/licenses/mit-license.html

import os
try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

package_dir = os.path.join(os.path.dirname(__file__),'ooxml')

setup(
    name='ooxml',
    version=file(os.path.join(package_dir,'version.txt')).read().strip(),
    author='Wensheng Wang',
    author_email='wenshengwang@gmail.com',
    license='MIT',
    description="Python interface for working with OOXML files such as docx, xlsx.",
    long_description=open(os.path.join(package_dir,'description.txt')).read(),
    url='https://bitbucket.org/wensheng/pyooxml',
    keywords="word excel powerpoint ooxml",
    classifiers=[
		'Operating System :: OS Independent',
		'Programming Language :: Python',
    	'License :: OSI Approved :: MIT License',
    	'Development Status :: 3 - Alpha',
	    'Intended Audience :: Developers',
    	'Topic :: Office/Business :: Financial :: Spreadsheet',
		'Topic :: Software Development :: Libraries :: Python Modules',
    ],
	platforms = 'Platform Independent',
    packages=['ooxml'],
    zip_safe=False,
    include_package_data=True,
    install_requires=[],
    entry_points = {},
    extras_require={},
    data_files=[
          ('ooxml/ooxml-templates', ['ooxml/ooxml-templates/workbook.xlsx']),
          ]
    )
