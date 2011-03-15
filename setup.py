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

setup(
    name='ooxml',
    version=open('version.txt').read().strip(),
    author='Wensheng Wang',
    author_email='wenshengwang@gmail.com',
    license='MIT',
    description="Python interface for working with OOXML files such as docx, xlsx.",
    long_description=open('README.txt').read(),
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
    data_files=[
        ('ooxml/ooxml-templates', ['ooxml/ooxml-templates/workbook.xlsx']),
    ]
)
