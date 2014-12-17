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
    EXCEL_FILEL_TYPE_LIST = ['.xlsx', '.xlsm', '.xls']

    def __init__(self, path):
        self.path = path


    def parse_all(self):
        """
        :return:
        """

    @property
    def floder(self):
        """
        文件夹路径
        :return:
        """

        return os.path.join(os.getcwd(), self.path)


    @property
    def filenameS(self):
        """
        获得路径下所有文件的文件名
        :return:
        """
        for d, fd, fl in os.walk(self.floder):




    @classmethod
    def isFileType(cls, filename):
        """
        是否是可以解析的文件类型
        :param filename:
        :return:
        """
        filename

        return False
