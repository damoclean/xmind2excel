from xmindparser import xmind_to_dict
import os, xlwt
from tkinter.messagebox import showinfo
import tkinter.messagebox


class XlwtSeting(object):

    @staticmethod  # 静态方法装饰器，使用此装饰器装饰后，可以直接使用类名.方法名调用（XlwtSeting.styles()），并且不需要self参数
    def template_one(worksheet):
        dicts = {"horz": "CENTER", "vert": "CENTER"}
        sizes = [30, 30, 30, 60, 45, 15, 15, 15, 15]
        se = XlwtSeting()
        style = se.styles()
        style.alignment = se.alignments(**dicts)
        style.font = se.fonts(bold=True)
        style.borders = se.borders()
        style.pattern = se.patterns(7)
        se.heights(worksheet, 0)
        for i in range(len(sizes)):
            se.widths(worksheet, i, size=sizes[i])
        return style
    @staticmethod
    def template_two():
        dicts2 = {"vert": "CENTER"}
        se = XlwtSeting()
        style = se.styles()
        style.borders = se.borders()
        style.alignment = se.alignments(**dicts2)
        return style

    @staticmethod
    def styles():
        """设置单元格的样式的基础方法"""
        style = xlwt.XFStyle()
        return style

    @staticmethod
    def borders(status=1):
        """设置单元格的边框，
        细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13"""
        border = xlwt.Borders()
        border.left = status
        border.right = status
        border.top = status
        border.bottom = status
        return border

    @staticmethod
    def heights(worksheet, line, size=4):
        """设置单元格的高度"""

        worksheet.row(line).height_mismatch = True
        worksheet.row(line).height = size * 128

    @staticmethod
    def widths(worksheet, line, size=11):
        """设置单元格的宽度"""

        worksheet.col(line).width = size * 256

    @staticmethod
    def alignments(wrap=1, **kwargs):
        """设置单元格的对齐方式，
        ：接收一个对齐参数的字典{"horz": "CENTER", "vert": "CENTER"}horz（水平），vert（垂直）
        ：horz中的direction常用的有：CENTER（居中）,DISTRIBUTED（两端）,GENERAL,CENTER_ACROSS_SEL（分散）,RIGHT（右边）,LEFT（左边）
        ：vert中的direction常用的有：CENTER（居中）,DISTRIBUTED（两端）,BOTTOM(下方),TOP（上方）"""

        alignment = xlwt.Alignment()

        if "horz" in kwargs.keys():
            alignment.horz = eval(f"xlwt.Alignment.HORZ_{kwargs['horz'].upper()}")
        if "vert" in kwargs.keys():
            alignment.vert = eval(f"xlwt.Alignment.VERT_{kwargs['vert'].upper()}")
        alignment.wrap = wrap  # 设置自动换行
        return alignment

    @staticmethod
    def fonts(name='宋体', bold=False, underline=False, italic=False, colour='black', height=11):
        """设置单元格中字体的样式，
        默认字体为宋体，不加粗，没有下划线，不是斜体，黑色字体"""

        font = xlwt.Font()
        # 字体
        font.name = name
        # 加粗
        font.bold = bold
        # 下划线
        font.underline = underline
        # 斜体
        font.italic = italic
        # 颜色
        font.colour_index = xlwt.Style.colour_map[colour]
        # 大小
        font.height = 20 * height
        return font

    @staticmethod
    def patterns(colors=1):
        """设置单元格的背景颜色，该数字表示的颜色在xlwt库的其他方法中也适用，默认颜色为白色
        0 = Black, 1 = White,2 = Red, 3 = Green, 4 = Blue,5 = Yellow, 6 = Magenta, 7 = Cyan,
        16 = Maroon, 17 = Dark Green,18 = Dark Blue, 19 = Dark Yellow ,almost brown), 20 = Dark Magenta,
        21 = Teal, 22 = Light Gray,23 = Dark Gray, the list goes on..."""

        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = colors
        return pattern


class Xmind2Excel(XlwtSeting):
    def __init__(self):
        self.workbook = xlwt.Workbook(encoding='utf-8')
        self.worksheet = self.workbook.add_sheet('sheet1', cell_overwrite_ok=True)
    def save(self, name):
        self.workbook.save(name)

    def writeExcel(self, row, case, excelName):
        sort = 0
        row0 = ["用例目录", '用例名称', '前置条件', '用例步骤', '预期结果', "用例类型", '用例状态', '用例等级', '创建人']
        # 生成第一行
        style2 = self.template_one(self.worksheet)
        for i in range(0, len(row0)):
            self.worksheet.write(0, i, row0[i],style2)

        for key, value in case.items():
            self.worksheet.write(row, sort, value)
            sort = sort + 1
        self.save(excelName)

    def numberLen(self, value, errornum=None):
        try:
            return len(value['topics'])
        except KeyError:
            if errornum == 2:
                 tkinter.messagebox.askokcancel('提示', '用例有前置条件 "{0}",没有测试步骤和预期结果喔！ 请确认是否如此！'.format(value['title']))
                #print('案例 "{0}",没有测试步骤和预期结果喔！ 请确认是否如此！'.format(value['title']))
            if errornum == 3:
                 tkinter.messagebox.showinfo(title='提示', message='用例有前置条件 "{0}",没有预期结果喔！ 请填写后重新执行！'.format(value['title']))
                #print('案例 "{0}",没有预期结果喔！ 请填写后重新执行！'.format(value['title']))
            return 0

    def writeCase(self, catalogue, value):
        #TestStepMunFlag = self.numberLen(value, 2)
        caseList = []
        fList = []
        sList = []
        tList = []
        self.caseDict['myTestCase'] = '-'.join(catalogue)
        self.caseDict['TestCase'] = value['title']
        if 'topics' in value:
            for i1 in value['topics']:
                fList.append(i1['title'])
                if 'topics' in i1:
                    for i2 in i1['topics']:
                        sList.append(i2['title'])
                        if 'topics' in i2:
                            for i3 in i2['topics']:
                                tList.append(i3['title'])
        if len(fList) > 0:
            caseList.append('\n'.join(fList))
            if len(sList) > 0:
                caseList.append('\n'.join(sList))
                if len(tList) > 0:
                    caseList.append('\n'.join(tList))
        if len(caseList) == 3:
            self.caseDict['Testprecondition'] = caseList[0]
            self.caseDict['TestStep'] = caseList[1]
            self.caseDict['TestResult'] = caseList[2]
        elif len(caseList) == 2:
            self.caseDict['Testprecondition'] = ''
            self.caseDict['TestStep'] = caseList[0]
            self.caseDict['TestResult'] = caseList[1]
        elif len(caseList) == 1:
            self.caseDict['Testprecondition'] = ''
            self.caseDict['TestStep'] = ''
            self.caseDict['TestResult'] = caseList[0]
        elif len(caseList) == 0:
            self.caseDict['Testprecondition'] = ''
            self.caseDict['TestStep'] = ''
            self.caseDict['TestResult'] = ''
        self.caseDict['Testbelong'] = '功能测试'
        self.caseDict['Teststatus'] = '正常'
        Testpriority1 = value['makers'][0]
        if Testpriority1 == 'priority-1':
            Testpriority = 'P0'
        elif Testpriority1 == 'priority-2':
            Testpriority = 'P1'
        elif Testpriority1 == 'priority-3':
            Testpriority = 'P2'
        elif Testpriority1 == 'priority-4':
            Testpriority = 'P3'
        else:
            Testpriority = 'p3'
        self.caseDict['Testpriority'] = Testpriority
        self.caseDict['Testoperator'] = self.operator
        self.rowNum = self.rowNum + 1
        self.writeExcel(self.rowNum, self.caseDict, self.filePath)

    def xmind_title(self, value):
        """获取xmind标题内容"""
        return value['title']

    def xmind_traversal(self, d, n_tab=-1):
        if isinstance(d, list):
            for i in d:
                self.xmind_traversal(i, n_tab)

        elif isinstance(d, dict):
            n_tab += 1
            if 'makers' in d.keys():
                self.writeCase(self.dirList, d)
            else:
                for key, value in d.items():
                    self.xmind_traversal(value, n_tab)
        else:
            # print("{}{}".format(n_tab, d))
            self.dirList = self.dirList[0:n_tab]
            self.dirList.append(d)
            # print('list:' + str(self.dirList))

    def xmind2excel(self, FileName, operator):

        self.rowNum = 0  # 计算测试用例的条数
        self.caseDict = {}
        self.dirList = []
        self.operator = operator
        self.filePath = FileName.replace('.xmind','.xls')
        self.XmindContent = xmind_to_dict(FileName)[0]['topic']  # xmind内容
        self.xmind_traversal(self.XmindContent)
        #print('生成{0}条用例，请检查是否有误。'.format(self.rowNum))
        showinfo(title='转换结束', message='生成{0}条用例，请检查是否有误。'.format(self.rowNum))


#if __name__ == '__main__':
#    XmindFile='/Users/xenos_liu/xmind2execl/data/app.xmind'
#    operator = '江彪'
#    Xmind2Excel().xmind2excel(XmindFile, operator)