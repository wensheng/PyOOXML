# -*- coding: utf-8 -*-
"""
read test for xlsx
"""
# Copyright (C) 2011 Wensheng Wang
#
# MIT License

import sys
import os
import unittest
import doctest
import random
import sys
from StringIO import StringIO

test_dir = os.path.dirname(__file__)
sys.path.insert(0,os.path.join(test_dir,".."))

from ooxml.spreadsheet import Spreadsheet

def _randomName():
    _randy = random.Random()
    result = u""
    for _ in range(5 + _randy.randint(0, 10)):
        result += unichr(_randy.randint(ord("a"), ord("z")))
    return result

class TestXlsxRead(unittest.TestCase):
    def setUp(self):
        self.workbook = Spreadsheet(os.path.join(test_dir,"sample1.xlsx"))

    def test_rows(self):
        rows = self.workbook.sheet(1).rows()
        self.assertEqual(rows[1].values(),['1','5','5'])
        self.assertEqual(rows[2].values(),['2','5','10'])
        self.assertEqual(rows[3].values(),['3','5','15'])
        self.assertEqual(rows[4].values(),['4','5','20'])
        self.assertEqual(rows[5].values(),['5','5','25'])
        self.assertEqual(rows[6].values(),['6','5','30'])
        self.assertEqual(rows[7].values(),['7','5','35'])
        self.assertEqual(rows[8].values(),['8','5','40'])
        self.assertEqual(rows[9].values(),['9','5','45'])

    def test_row(self):
        self.assertEqual(self.workbook.sheet(1).row(1).values(), ['1','5','5'])

    def test_cell(self):
        self.assertEqual(self.workbook.sheet(1).cell(1,1).value, '1')
    
if __name__ == "__main__":
    unittest.main()
