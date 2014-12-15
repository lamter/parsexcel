#coding:utf-8
'''
Created on 2014-12-15

@author: Shawn
'''

import os
import traceback


def getAllFileAndDirByPath(path):
    '''获取某个路径下的所有文件和文件夹.不会递归去查询文件夹下的文件'''
    dirArray = []
    fileArray = []
    try:
        fileNameArray = os.listdir(path)
        for fileName in fileNameArray:
            fullPath = os.path.join(path, fileName)

            #print "fullPath --> ", fullPath
            #文件夹
            if os.path.isdir(fullPath) == True:
                dirArray.append(fileName)

            #文件
            elif os.path.isfile(fullPath) == True:
                fileArray.append(fileName)
    except:
        print traceback.print_exc()
    finally:
        return fileArray, dirArray


def getFileBySuffix(path, suffix):
    '''获取这个目录以某个后缀结尾的文件.ect: .txt'''
    resultFileArray = []
    fileArray, dirArray = getAllFileAndDirByPath(path)
    for fileName in fileArray:
        try:
            fileNameSuffix = os.path.splitext(fileName)[1]
            if fileNameSuffix == suffix:
                resultFileArray.append(fileName)
        except:
            print traceback.print_exc()

    return resultFileArray