#coding:utf-8
'''
Created on 2014-12-15

@author: Shawn
'''


import json
import unittest

from parser import parse_xlsm

Encoder = json.JSONEncoder()
Decoder = json.JSONDecoder()


def suite():
    testSuite1 = unittest.makeSuite(TestParseXlsm, "test")
    alltestCase = unittest.TestSuite([testSuite1, ])
    return alltestCase


class TestParseXlsm(unittest.TestCase):
    '''
    测试武将相关
    '''
    def setUp(self):
        self.excelFilePath = 'excelfile'



    def test_parse_xlsm(self):
        """
        解析 .xlsm 格式的文件
        :return:
        """
        wb = parse_xlsm.XlsmWorkBook(self.excelFilePath + "/officer.xlsm")














