#coding:utf-8
'''
Created on 2014-12-17

@author: Shawn
'''

import json
import unittest

import parser

Encoder = json.JSONEncoder()
Decoder = json.JSONDecoder()


def suite():
    testSuite1 = unittest.makeSuite(TestParser, "test")
    alltestCase = unittest.TestSuite([testSuite1, ])
    return alltestCase


class TestParser(unittest.TestCase):
    '''
    测试武将相关
    '''
    def setUp(self):
        self.excelFilePath = 'excelfile'


    def test_getAllFilenameS(self):
        """
        获得指定目录下所有文件的文件名
        :return:
        """
        theParser = parser.Parser(self.excelFilePath)
        print theParser.excelFilenameS


    def test_load(self):
        """
        读取
        :return:
        """
        theParser = parser.Parser(self.excelFilePath)
        theParser.load()



    def test_setInfoArray(self):
        """
        解析 成 infoArray 的数据格式
        :return:
        """
        theParser = parser.Parser(self.excelFilePath)
        theParser.load()
        for ws in theParser.getAllWorksheet():
            print ws.infoArray



