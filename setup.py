#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
easyxlsx setup.py
"""

from setuptools import setup, find_packages

__version__ = '0.3.9'

INSTALL_REQUIRES = ['XlsxWriter>=1.1.1']

setup(
    name='easyxlsx',
    version=__version__,
    author='codingcat',
    packages=find_packages(),
    description='easy way to use xlsxwriter.',
    author_email='istommao@gmail.com',
    install_requires=INSTALL_REQUIRES,
    include_package_data=True,
    zip_safe=False,
    license='http://opensource.org/licenses/MIT',
    url='https://github.com/istommao/easyxlsx'
)
