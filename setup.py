#!/usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import absolute_import, division, print_function, unicode_literals
import io
import os
import os.path
import re

from setuptools import setup, find_packages, Extension

PACKAGE = 'xlsx_streaming'
description = 'Export your data as an xlsx stream'


def _open(*file_path):
    here = os.path.abspath(os.path.dirname(__file__))
    return io.open(os.path.join(here, *file_path), 'r', encoding='utf-8')


def get_version(package):
    version_pattern = re.compile(r"^__version__ = [\"']([\w_.-]+)[\"']$", re.MULTILINE)
    with _open(*(package.split('.') + ['__init__.py'])) as version_file:
        matched = version_pattern.search(version_file.read())
    return matched.groups()[0]


def get_long_description():
    with _open('README.rst') as readme_file:
        return readme_file.read()


REQUIREMENTS = [
    'zipstream>=1.1.3',
]

setup(
    name=PACKAGE,
    version=get_version(PACKAGE),
    description=description,
    long_description=get_long_description(),
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3.4',
    ],
    keywords=['xlsx', 'excel', 'streaming'],
    author='Polyconseil',
    author_email='opensource+%s@polyconseil.fr' % PACKAGE,
    url='https://github.com/Polyconseil/%s/' % PACKAGE,
    license='GNU GPLv3',
    packages=find_packages(exclude=['docs', 'tests']),
    setup_requires=[
        'setuptools',
    ],
    install_requires=REQUIREMENTS,
    test_suite='tests'
)
