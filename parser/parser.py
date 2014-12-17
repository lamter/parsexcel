#coding:utf-8
'''
Created on 2014-12-15

@author: Shawn
'''

import os


import common



class Parser(object):
    """
    解析指定路径下全部的excel文件
    """
    ''' 可以解析的 excel 文件类型 '''
    EXCEL_FILEL_TYPE_LIST = ['xlsx', 'xlsm', 'xls']

    def __init__(self, path):
        self.path = path
        self。


    def parse_all(self):
        """
        解析所有文件
        :return:
        """
        for fn in self.excelFilenameS:
            filepath = os.path.join(self.floder, fn)
            return filepath


    @property
    def floder(self):
        """
        文件夹路径
        :return:
        """

        return os.path.join(os.getcwd(), self.path)


    @property
    def excelFilenameS(self):
        """
        获得路径下所有excel文件的文件名
        :return:
        """
        filenameS = []
        for d, fd, fl in os.walk(self.floder):
            filenameS = fl
            break

        excelFilesnameS = []
        for filename in filenameS:
            sufix = os.path.splitext(filename)[1][1:]
            if sufix in self.EXCEL_FILEL_TYPE_LIST:
                excelFilesnameS.append(filename)

        return excelFilesnameS
