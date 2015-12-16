# -*- coding: utf-8 -*-

from __future__ import unicode_literals

import sys


if sys.version_info < (3,):
    text_type = unicode
else:
    text_type = str

def itertree(tree):
    if sys.version_info < (2, 7):
        # getiterator returns a list in python 2.6!
        return iter(tree.getiterator())
    return tree.iter()
