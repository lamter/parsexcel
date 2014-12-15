#coding:utf-8
'''
Created on 2014-12-15

解析excel导出的 *.xml 文件

@author: Shawn
'''

try:
    import xml.etree.cElementTree as ET             # C语言编译的xml包
except:
    import xml.etree.ElementTree as ET

import os

import common


EXCEL_OPEN_XML_WORK_SHEET = "{urn:schemas-microsoft-com:office:spreadsheet}Worksheet"               # 微软excel的电子表格节点
EXCEL_OPEN_XML_WORK_SHEET_NAME = "{urn:schemas-microsoft-com:office:spreadsheet}Name"               # 微软excel的电子表格名属性
EXCEL_OPEN_XML_WORK_SHEET_TBALE = "{urn:schemas-microsoft-com:office:spreadsheet}Table"             # 微软excel的表格内容标签
EXCEL_OPEN_XML_WORK_SHEET_ROW = "{urn:schemas-microsoft-com:office:spreadsheet}Row"                 # 微软excel的行标签
EXCEL_OPEN_XML_WORK_SHEET_CELL = "{urn:schemas-microsoft-com:office:spreadsheet}Cell"               # 微软excel的单元格标签
EXCEL_OPEN_XML_WORK_SHEET_COMMENT = "{urn:schemas-microsoft-com:office:spreadsheet}Comment"         # 微软excel的单元格内注释标签
EXCEL_OPEN_XML_WORK_SHEET_DATA = "{urn:schemas-microsoft-com:office:spreadsheet}Data"               # 微软excel的单元格内数据标签
EXCEL_OPEN_XML_WORK_SHEET_CELL_INDEX = "{urn:schemas-microsoft-com:office:spreadsheet}Index"        # 微软excel的单元格下标属性



class XmlWorkBook(object):
    '''
    每个excel文件会导出一份*.xml文件，会包含多个Worksheet
    '''
    def __init__(self, name, floderPath):
        ''' 文件名 '''
        self.fileName = name
        ''' 所在路径 '''
        self.fileFloder = floderPath

        ''' worksheet名字 '''
        self.worksheetNameArray = []
        self.worksheetDic = {}


    def log(self):
        print "%s================>" % self.__class__.__name__
        for k,v in self.__dict__.items():
            print k, ':', v
        print "%s<================" % self.__class__.__name__


    def getFilePath(self):
        '''
        获得文件路径
        '''
        return os.path.join(self.fileFloder, self.fileName)


    def parseWorkSheet(self):
        '''
        解析出workSheet
        '''
        tree = ET.parse(self.getFilePath())    #载入数据
        root = tree.getroot()

        ''' 这样便可以遍历根元素的所有子元素(这里是worksheet元素) '''
        for worksheetTree in root.findall(EXCEL_OPEN_XML_WORK_SHEET):
            ''' 用.tag得到该子元素的名称 '''
            sheetName = worksheetTree.get(EXCEL_OPEN_XML_WORK_SHEET_NAME)

            ''' 用于检查是否有重名的sheet表 '''
            self.worksheetNameArray.append(sheetName)
            ''' 生成 worksheet实例 '''
            worksheet = XmlWorkSheet(sheetName, worksheetTree)
            self.worksheetDic[sheetName] = worksheet


    def parseSheetTree(self):
        '''
        解析sheetTree
        :return:
        '''
        for worksheet in self.worksheetDic.values():
            worksheet.parseTree()


    def getSheetNameArray(self):
        '''
        :return [sheetName and NO None]
        '''

        return [sheetName for sheetName in self.worksheetNameArray if sheetName != None]


    def getSheetArray(self):
        '''
        :return [worksheet]
        '''
        return [worksheet for worksheet in self.worksheetDic.values() if worksheet != None]



class XmlWorkSheet(object):
    '''
    这里就是每张表
    :param object:
    :return:
    '''
    def __init__(self, sheetName, worksheetTree):
        '''
        :param tree: 有lxml解出来的数据表worksheet的树分支
        :return:
        '''
        self.sheetName = sheetName
        self.worksheetTree = worksheetTree

        self.data = []


    def parseTree(self):
        '''
        解析tree
        '''
        dataArray = []
        ''' 获得表分支 '''
        tableBranch = self.parseTableTree(self.worksheetTree)

        ''' 将数据组织成矩阵 '''
        rowTreeArray = self.parseRowTreeArray(tableBranch)

        infoArray = []
        ''' 组织数据为矩阵 '''
        for rowTree in rowTreeArray:
            datas = []
            for cellTree in self.parseCellArray(rowTree):
                ''' 单元格是否指定了特定的下标 '''
                cellIndex = cellTree.get(EXCEL_OPEN_XML_WORK_SHEET_CELL_INDEX)
                if cellIndex != None:
                    cellIndex = int(cellIndex)
                    ''' 有下标需要处理，要在datas里面填None到下标处，excel的下标是从0开始的 '''
                    while len(datas) < cellIndex-1:
                        datas.append(None)

                ''' 取出每个单元格中的数据 '''
                dat = self.parseData(cellTree)

                datas.append(dat)

            infoArray.append(datas)

        self.infoArray = XmlWorkSheet.parseInfoArray(infoArray)


    def getInfoArray(self):
        '''
        返回一个
        :return:
        '''
        return [info.copy() for info in self.infoArray]


    @staticmethod
    def parseTitleArray(infoArray):
        '''
        :param infoArray:二维数组应该为一个包含表头的[[title], [data1], [data2], ...]
        :return:[title, ...]
        '''
        tittleArray = []
        for title in infoArray[0]:
            if title == None:
                break
            tittleArray.append(title)
        return tittleArray


    @staticmethod
    def parseDataArray(infoArray):
        '''
        :param infoArray:二维数组应该为一个包含表头的[[title], [data1], [data2], ...]
        :return:[[data1],[data2] ...]
        '''
        tmpDatasArray = infoArray[1:]
        dataArray = []
        titleArray = XmlWorkSheet.parseTitleArray(infoArray)
        for tmpDatas in tmpDatasArray:
            datas = tmpDatas[:len(titleArray)]
            while len(datas) < len(titleArray):
                datas.append(None)
            dataArray.append(datas)
        return dataArray


    @staticmethod
    def parseInfoArray(infoArray):
        '''
        这个函数仅为在方法 parseTree()中使用
        :param infoArray:二维数组应该为一个包含表头的[[title], [data1], [data2], ...]
        :return:[{info}, ...]
        '''
        if len(infoArray) < 3:
            ''' 最少需要一行表头和一行数据 '''
            return []

        ''' 默认第一行为key '''
        titleArray = XmlWorkSheet.parseTitleArray(infoArray)
        dataArray = XmlWorkSheet.parseDataArray(infoArray)

        # dataArray.insert(0, titleArray)
        # excel = makeTable.MakeTable(dataArray, True)
        # htmlCode = excel.makeTable()
        # common.makeTextFile(htmlCode, user='lamter')
        # exit()

        infoDicArray = []
        for datas in dataArray:
            ''' datas = [[data1], [data2], ...] '''
            if XmlWorkSheet.isNullString(datas):
                ''' 如果是空行就跳过 '''
                continue
            info = {}
            for i, title in enumerate(titleArray):
                data = datas[i]
                ''' 如果这个数据是None, 不加入info中 '''
                if data == None:
                    continue
                info[title] = data
            infoDicArray.append(info)

        # dataArray = []
        # for info in infoDicArray:
        #     if info == None:
        #         print 'info=>', info
        #         exit()
        #     dataArray.append(['%s' % i for i in info.keys()])
        #     print 'key=>', len(['%s' % i for i in info.keys()])
        #     # print ['%s' % i for i in info.keys()]
        #     dataArray.append(['%s' % i for i in  info.values()])
        #     # print ['%s' % i for i in info.keys()]
        #     print 'values=>', len(['%s' % i for i in info.values()])
        #
        # excel = makeTable.MakeTable(dataArray, True)
        # htmlCode = excel.makeTable()
        # common.makeTextFile(htmlCode, user='lamter')
        # exit()

        return infoDicArray

    @staticmethod
    def isNullString(datas):
        '''
        判断这一行是否在表中为空行，即title对应的data全部为None
        :return:
        '''
        for data in datas:
            if data != None:
                ''' 只要有一个data不为None,就是非空行 '''
                return False

        ''' 否则为空行 '''
        return True



    @staticmethod
    def parseTableTree(worksheetTree=None):
        '''
        从worksheetTree里拿到tableTree
        :return:
        '''
        for branch in worksheetTree:
            if branch.tag == EXCEL_OPEN_XML_WORK_SHEET_TBALE:
                ''' 根据标签名，拿到表分支，一个worksheet只有一个table分支 '''
                return branch


    @staticmethod
    def parseRowTreeArray(tableTree):
        '''
        从
        :param tableTree:
        :return:
        '''
        rowTreeArray = []
        for row in tableTree:
            if row.tag != EXCEL_OPEN_XML_WORK_SHEET_ROW:
                ''' 只选中行标签 '''
                continue
            rowTreeArray.append(row)
        return rowTreeArray


    @staticmethod
    def parseCellArray(rowTree):
        '''
        从行树中获得cell树
        :param rwoTree:
        :return:
        '''
        cellArray = []
        for cell in rowTree:
            if cell.tag != EXCEL_OPEN_XML_WORK_SHEET_CELL:
                ''' 只选中单元格标签 '''
                continue
            cellArray.append(cell)
        return cellArray


    @staticmethod
    def parseData(cellTree):
        '''
        从单元树中获得数据
        :return:
        '''
        ''' 处理了单元格 '''
        for data in cellTree:
            if data.tag == EXCEL_OPEN_XML_WORK_SHEET_DATA:
                ''' 只选中数据标签，且读取一次后跳出这个循环 '''
                # if len(data) > 0:
                    # print data.text_content()
                    # print data.text
                    # exit()
                while len(data) > 0:
                    ''' 如果data还有更低级标签的话，反复挖掘分支 '''
                    # print data.tag
                    data = data[0]

                return data.text


class Parser(object):
    '''
    由这个类来操作解表
    '''
    def __init__(self, path):
        '''
        :param path: xml文件的路径
        :return:
        '''
        self.xmlFloderPath = path
        print u'解析xml文件的路径为...'
        print path
        self.workBookDic = {}
        self.xmlSuffix = ".xml"
        ''' 解析成功与否 '''
        self.isSuss = False


    def isHaveXmlData(self):
        '''
        给出的文件夹路径是否正确
        :return:
        '''

        if self.xmlFloderPath == None or os.path.exists(self.xmlFloderPath) == False:
            print u'找不到xml文件路径:' , self.xmlFloderPath
            return False

        xmlFilseNameArray = common.getFileBySuffix(self.xmlFloderPath, self.xmlSuffix)
        if len(xmlFilseNameArray) == 0:
            print u'路径没有*.xml文件'
            return False


    def importXmlWorkBook(self):
        '''
        生成xmlWorkSheet对象
        :return:
        '''
        xmlFilseNameArray = common.getFileBySuffix(self.xmlFloderPath, self.xmlSuffix)
        for xmlFileName in xmlFilseNameArray:
            ''' 设成*.xml的实例 '''
            workBook = XmlWorkBook(xmlFileName, self.xmlFloderPath)
            # workBook.log()
            self.workBookDic[xmlFileName] = workBook


    def importXmlWorkSheet(self):
        '''
        从WorkBook中解析出WorkSheet
        :return:
        '''
        for workBook in self.workBookDic.values():
            workBook.parseWorkSheet()



    def isRepeatedSheetName(self):
        '''
        :return: bool(是否存在重复的表)
        '''

        sheetNameArray = []
        for workBook in self.workBookDic.values():
            sheetNameArray.extend(workBook.getSheetNameArray())

        ''' 比较名字去重后的表数量 '''
        return len(sheetNameArray) != len(set(sheetNameArray))


    def parseXmlWorkSheet(self):
        '''
        解析workSheetTree
        :return:
        '''
        for workBook in self.workBookDic.values():
            workBook.parseSheetTree()


    def getWorksheetArray(self):
        '''
        :return:
        '''
        worksheetArray = []
        for workbook in self.workBookDic.values():
            worksheetArray.extend(workbook.getSheetArray())

        return worksheetArray


    def getWorkBookArray(self):
        '''
        :return:
        '''
        return self.workBookDic.values()


    def xml(self):
        '''
        整个解析xml的流程都在这个函数里进行
        :return:
        '''
        if self.isHaveXmlData() == False:
            return

        ''' 根据指定的xmlFolderPath给出的*.xml文件来生成指定的*.xml文件对象 '''
        self.importXmlWorkBook()

        ''' 遍历所有WorkBook，根据给出的路径，生成未初始化的WorkShet '''
        self.importXmlWorkSheet()

        # ''' 检查是否存在重复的表 '''
        # if self.isRepeatedSheetName() == True:
        #     print u'存在重复的表名'
        #     return

        self.parseXmlWorkSheet()

        ''' 解析成功 '''
        self.isSuss = True





if __name__ == "__main__":
    parse = Parser("/Users/lamter/workspace/BlackSG/table")
    parse.xml()




