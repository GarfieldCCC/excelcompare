import os
import sys
import xlrd
import xlwt
import ctypes
import numpy as np
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myappid")


class ExcelOutput:
    def __init__(self):
        # 格式: 居中加边框
        font = xlwt.Font()
        font.bold = True
        self.style = xlwt.XFStyle()
        self.style_head = xlwt.XFStyle()
        self.style_note = xlwt.XFStyle()
        # 设置居中
        al = xlwt.Alignment()
        al.horz = 0x02
        al.vert = 0x01
        self.style.alignment = al
        self.style_head.alignment = al
        # 加边框
        borders = xlwt.Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        self.style.borders = borders
        self.style_head.borders = borders
        self.style_head.font = font
        self.style_note.font = font

    def adjust_col(self, filename):
        """行宽自适应"""
        max_list = [0, 0, 0, 0, 0, 0, 0]
        data = xlrd.open_workbook(filename)
        sheet = data.sheet_by_index(0)
        for i in range(sheet.nrows):
            for j in range(sheet.ncols):
                max_list[j] = max(max_list[j], len(sheet.cell(i, j).value.encode('gb18030')))
        return max_list

    def output_head(self, ws):
        """打印表头"""
        ws.write_merge(0, 1, 0, 0, "部件", self.style_head)
        ws.write_merge(0, 1, 6, 6, "备注", self.style_head)
        ws.write_merge(0, 1, 1, 1, "代号", self.style_head)
        ws.write_merge(0, 0, 2, 3, "变更之前", self.style_head)
        ws.write_merge(0, 0, 4, 5, "变更之后", self.style_head)
        ws.write(1, 2, "变更字段", self.style_head)
        ws.write(1, 4, "变更字段", self.style_head)
        ws.write(1, 3, "变更内容", self.style_head)
        ws.write(1, 5, "变更内容", self.style_head)

    def output_excel(self, ws, wb, out_path, change, delete, add, name, start_row, max_list=[]):
        """打印并输出至Excel"""
        end_row = start_row
        merge = {}

        if len(max_list) != 0:
            ws.col(0).width = 256 * (max_list[0] + 2)
            ws.col(1).width = 256 * (max_list[1] + 2)
            ws.col(2).width = 256 * (max_list[2] + 2)
            ws.col(3).width = 256 * (max_list[3] + 2)
            ws.col(4).width = 256 * (max_list[4] + 2)
            ws.col(5).width = 256 * (max_list[5] + 2)
            ws.col(6).width = 256 * (max_list[6] + 2)

        # 1. 变动部分
        start = end_row
        for i in change:
            value = change[i]
            length = len(value[0])
            if length > 1:
                merge[i] = [end_row, end_row + length - 1]
            for j in range(length):
                ws.write(end_row, 0, label=name, style=self.style)
                ws.write(end_row, 1, label=i, style=self.style)
                ws.write(end_row, 2, label=value[0][j], style=self.style)
                ws.write(end_row, 3, label=value[1][j], style=self.style)
                ws.write(end_row, 4, label=value[2][j], style=self.style)
                ws.write(end_row, 5, label=value[3][j], style=self.style)
                end_row = end_row + 1
        end = end_row - 1
        if start <= end:
            ws.write_merge(start, end, 6, 6, "变动", self.style)

        # 2. 去掉部分
        start = end_row
        for i in delete:
            ws.write(end_row, 0, label=name, style=self.style)
            ws.write(end_row, 1, label=i[1], style=self.style)
            ws.write(end_row, 2, label='名称', style=self.style)
            ws.write(end_row, 3, label=i[3], style=self.style)
            ws.write(end_row, 4, style=self.style)
            ws.write(end_row, 5, style=self.style)
            end_row = end_row + 1
        end = end_row - 1
        if start <= end:
            ws.write_merge(start, end, 6, 6, "去掉", self.style)

        # 3. 新增部分
        start = end_row
        for i in add:
            ws.write(end_row, 0, label=name, style=self.style)
            ws.write(end_row, 1, label=i[1], style=self.style)
            ws.write(end_row, 2, style=self.style)
            ws.write(end_row, 3, style=self.style)
            ws.write(end_row, 4, label='名称', style=self.style)
            ws.write(end_row, 5, label=i[3], style=self.style)
            end_row = end_row + 1
        end = end_row - 1
        if start <= end:
            ws.write_merge(start, end, 6, 6, "新增", self.style)

        # 4. 合并单元格
        if start_row <= end_row - 1:
            ws.write_merge(start_row, end_row - 1, 0, 0, name, self.style)
        for i in merge:
            ws.write_merge(merge[i][0], merge[i][1], 1, 1, i, self.style)
        wb.save(out_path)
        return end_row


class ExcelCompare:
    def __init__(self):

        self.title_A = ['序号', '代号', '', '名称', '', '', '数量', '材料', '', '单重', '总重', '备料', '锻造', '压力试验', '叶轮焊接', '喷砂喷丸',
                        '涂装', '热处理', '机加工', '冷作', '平衡', '超转', '装配', '外协', '外购', '', '附注']
        self.title_B = ['序号', '代号', '', '名称', '', '', '数量', '材料', '', '单重', '总重', '备料', '锻造', '铸造', '热处理', '喷砂喷丸', '涂装',
                        '压力试验', '机加工', '焊接', '平衡超转', '装配', '酸洗', '外购', '其它', '', '附注']

        self.dic = {}

    def get_info(self, path):
        """读取Excel"""
        data = xlrd.open_workbook(path)
        print(data.sheet_names())
        return data

    def get_all_index(self, list, value):
        """获取某个元素的全部索引"""
        res = []
        for i in range(len(list)):
            if value == list[i]:
                res.append(i)
        return res

    def generate_dic(self, set_, list_):
        """生成字典，键：名(str)；值：全部索引(list)"""
        dic = {}
        for i in set_:
            dic[i] = self.get_all_index(list_, i)
        return dic

    def find_end(self, sheet):
        """找到最后一个不为空的数据的索引"""
        i = 19
        while sheet.row(i)[3].value == '':
            i = i - 1
        return i

    def generate_mat_sheet(self, sheet):
        """将Excel的一个sheet转换成二维矩阵"""
        end = self.find_end(sheet)
        temp = []
        res = []
        for i in range(2, end + 1):
            for j in sheet.row(i):
                temp.append(j.value)
            res.append(temp)
            temp = []
        return res

    def generate_mat_complete(self, index_list, excel):
        """根据索引生成完整的二维矩阵"""
        res = []
        for i in index_list:
            sheet = excel.sheet_by_index(i)
            res = res + self.generate_mat_sheet(sheet)
        return res

    def compress_mat(self, mat):
        """实现矩阵的压缩，去除空白行、合并分行的数据等"""
        matrix = mat
        i = len(matrix) - 1
        while i != -1:
            if matrix[i][0] == '':
                matrix[i - 1] = [i + j for i, j in zip(matrix[i - 1], matrix[i])]
                del matrix[i]
            i = i - 1
        return matrix

    def generate_index_dic(self, mat):
        """生成字典，键：代号(str)；值：索引(int)"""
        res = {}
        for i in mat:
            res[i[1]] = int(i[0]) - 1
        return res

    def get_dif(self, l1, l2):
        """比较两个列表，返回True、False列表"""
        return (np.array(l1) == np.array(l2)).tolist()

    def conversion(self, boo):
        """将True变成False，False变成True"""
        integer = np.array(boo).astype(int) - 1
        return integer.astype(bool)

    def compare_components(self, seta, setb):
        """比较两个表格的整体上的差异"""
        res = "注: "
        if len(seta) == len(setb):
            if seta == setb:  # 长度相等且内容相等
                res = res + "整体上不存在部件的增删改"
            else:  # 长度相等但是内容不相等
                a_has = seta - setb
                b_has = setb - seta
                res = res + "旧版文件有但是新版文件没有的部件是: " + str(a_has) + "; 新版文件有但是旧版文件没有的部件是: " + str(b_has)
        elif len(seta) > len(setb):  # 旧版数量大于新版数量
            a_has = seta - setb
            b_has = setb - seta
            if len(b_has) == 0:
                res = res + "旧版文件有但是新版文件没有的部件是: " + str(a_has)
            else:
                res = res + "旧版文件有但是新版文件没有的部件是: " + str(a_has) + "; 新版文件有但是旧版文件没有的部件是: " + str(b_has)
        elif len(seta) < len(setb):  # 旧版数量小于新版数量
            a_has = seta - setb
            b_has = setb - seta
            if len(a_has) == 0:
                res = res + "新版文件有但是旧版文件没有的部件是: " + str(b_has)
            else:
                res = res + "旧版文件有但是新版文件没有的部件是: " + str(a_has) + "; 新版文件有但是旧版文件没有的部件是: " + str(b_has)
        return res

    def compare(self, mata, matb):
        """表格比较"""
        index_mat_a = self.generate_index_dic(mata)
        index_mat_b = self.generate_index_dic(matb)

        change_res = {}
        delete_res = []
        add_res = []

        for ia in index_mat_a:
            if ia in index_mat_b:
                value = self.get_dif(mata[index_mat_a[ia]], matb[index_mat_b[ia]])
                if value[1: len(value)].count(False) > 0:
                    if ia == "dh":
                        value[0] = True
                    change_res[ia] = value
            else:
                delete_res.append(ia)
        for ib in index_mat_b:
            if ib not in index_mat_a:
                add_res.append(ib)
        return change_res, delete_res, add_res

    def output(self, excel_old, excel_new, dic_old, dic_new, mata, matb, name):
        """将比较结果输出, 变动部分为字典, 去掉和新增部分为列表"""
        change_res = {}
        delete_res = []
        add_res = []

        temp = int(mata[-1][0]) + 1
        mata = mata + [
            [temp, 'dh', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
             excel_old.sheet_by_index(dic_old[name][0]).cell(21, 20).value, '', '', '', '', '', '', '']]

        temp = int(matb[-1][0]) + 1
        matb = matb + [
            [temp, 'dh', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
             excel_new.sheet_by_index(dic_new[name][0]).cell(21, 20).value, '', '', '', '', '', '', '']]

        change, delete, add = self.compare(mata, matb)
        index_a, index_b = self.generate_index_dic(mata), self.generate_index_dic(matb)
        print("\n", name, "变动的部分: ")
        for i in change:  # i就是代号
            temp = self.conversion(change[i])
            title_a = np.array(self.title_A)[temp]
            title_b = np.array(self.title_B)[temp]
            content_a = np.array(mata[index_a[i]])[temp]
            content_b = np.array(matb[index_b[i]])[temp]
            print(i, ": ", title_a, content_a, title_b, content_b)
            if i != "dh":
                change_res[i] = [title_a.tolist(), content_a.tolist(), title_b.tolist(), content_b.tolist()]
            else:
                change_res[i] = [['图号'], content_a.tolist(), ['图号'], content_b.tolist()]

        print("\n", name, "去掉的部分: ")
        for i in delete:
            content = mata[index_a[i]]
            print(content)
            delete_res.append(content)
        print("\n", name, "新增的部分: ")
        for i in add:
            content = matb[index_b[i]]
            print(content)
            add_res.append(content)
        return change_res, delete_res, add_res


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        self.count = 1
        self.excel_compare = ExcelCompare()
        self.excel_write = ExcelOutput()
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(808, 503)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(120, 30, 611, 101))
        self.label.setObjectName("label")
        self.label.setFont(QtGui.QFont("幼圆", 10, QtGui.QFont.Bold))
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(120, 150, 151, 41))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.setFont(QtGui.QFont("幼圆", 9, QtGui.QFont.Bold))
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(120, 220, 151, 41))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.setFont(QtGui.QFont("幼圆", 9, QtGui.QFont.Bold))
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(120, 290, 151, 41))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.setFont(QtGui.QFont("幼圆", 9, QtGui.QFont.Bold))
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(120, 360, 151, 41))
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_4.setFont(QtGui.QFont("幼圆", 9, QtGui.QFont.Bold))
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(300, 150, 440, 31))
        self.label_2.setObjectName("label_2")
        self.label_2.setFont(QtGui.QFont("幼圆"))
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(300, 220, 440, 31))
        self.label_3.setObjectName("label_3")
        self.label_3.setFont(QtGui.QFont("幼圆"))
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 808, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.pushButton.clicked.connect(self.get_file_path_old)
        self.pushButton_2.clicked.connect(self.get_file_path_new)
        self.pushButton_3.clicked.connect(self.compare)
        self.pushButton_4.clicked.connect(self.clear)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        while not self.verify():
            self.verify()

    def verify(self):
        """身份验证, 在某台计算机上第一次使用, 则开启验证"""
        with open('ID.txt', "r") as f:
            mark = f.read()
        if mark == "False":

            dialog = QtWidgets.QInputDialog()
            dialog.setWindowIcon(QtGui.QIcon("info.ico"))
            dialog.setWindowTitle("验证")
            dialog.setLabelText("<font size='4'>请输入密码以继续使用 </font>")
            dialog.setOkButtonText("确定")
            dialog.setCancelButtonText("取消")

            if dialog.exec_() == QtWidgets.QInputDialog.Rejected:
                sys.exit()
            else:
                if dialog.textValue() != "sgxl8105369":
                    self.info_pwd()
                    return False
            with open('ID.txt', "w") as f:
                f.write("True")
            return True
        else:
            return True

    def info_pwd(self):  # 消息：密码错误
        msg_box = QtWidgets.QMessageBox()
        msg_box.setWindowIcon(QtGui.QIcon("info.ico"))
        msg_box.setWindowTitle("注意! ")
        msg_box.setText("<font size='4'>密码错误请重试  </font>")
        msg_box.setStandardButtons(QtWidgets.QMessageBox.Yes)
        btn_yes = msg_box.button(QtWidgets.QMessageBox.Yes)
        btn_yes.setText("确定")
        msg_box.exec_()

    def info_wrong_file(self):  # 消息：旧版文件不能为空
        msg_box = QtWidgets.QMessageBox()
        msg_box.setWindowIcon(QtGui.QIcon("info.ico"))
        msg_box.setWindowTitle("注意! ")
        msg_box.setText("<font size='4'>文件格式不正确, 请选择正确格式的文件  </font>")
        msg_box.setStandardButtons(QtWidgets.QMessageBox.Yes)
        btn_yes = msg_box.button(QtWidgets.QMessageBox.Yes)
        btn_yes.setText("确定")
        msg_box.exec_()

    def info_old(self):  # 消息：旧版文件不能为空
        msg_box = QtWidgets.QMessageBox()
        msg_box.setWindowIcon(QtGui.QIcon("info.ico"))
        msg_box.setWindowTitle("注意! ")
        msg_box.setText("<font size='4'>旧版文件为空, 请重新选择!  </font>")
        msg_box.setStandardButtons(QtWidgets.QMessageBox.Yes)
        btn_yes = msg_box.button(QtWidgets.QMessageBox.Yes)
        btn_yes.setText("确定")
        msg_box.exec_()

    def info_new(self):  # 消息：新版文件不能为空
        msg_box = QtWidgets.QMessageBox()
        msg_box.setWindowIcon(QtGui.QIcon("info.ico"))
        msg_box.setWindowTitle("注意! ")
        msg_box.setText("<font size='4'>新版文件为空, 请重新选择!  </font>")
        msg_box.setStandardButtons(QtWidgets.QMessageBox.Yes)
        btn_yes = msg_box.button(QtWidgets.QMessageBox.Yes)
        btn_yes.setText("确定")
        msg_box.exec_()

    def info_success(self, dirpath):  # 消息：导出成功
        msg_box = QtWidgets.QMessageBox()
        msg_box.setWindowIcon(QtGui.QIcon("info.ico"))
        msg_box.setWindowTitle("提示~ ")
        msg_box.setText("<font size='4'>导出成功! 是否打开导出文件所在目录? </font>")
        msg_box.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        btn_yes = msg_box.button(QtWidgets.QMessageBox.Yes)
        btn_yes.setText("好的")
        btn_no = msg_box.button(QtWidgets.QMessageBox.No)
        btn_no.setText("不了")
        msg_box.exec_()

        if msg_box.clickedButton() == btn_yes:
            start_directory = dirpath[0:[i for i, x in enumerate(dirpath) if x == "/"][-1]]
            os.startfile(start_directory)
        else:
            pass

    def get_file_path_old(self):
        """获取旧版文件路径"""
        path = QtWidgets.QFileDialog.getOpenFileName(filter="Excel Files (*.xlsx;*.xls;)")[0]
        self.label_2.setText("旧版文件名: " + path.split("/")[-1])
        self.excel_compare.dic[1] = path

    def get_file_path_new(self):
        """获取新版文件路径"""
        path = QtWidgets.QFileDialog.getOpenFileName(filter="Excel Files (*.xlsx;*.xls;)")[0]
        self.label_3.setText("新版文件名: " + path.split("/")[-1])
        self.excel_compare.dic[2] = path

    def compare(self):
        """比较输出"""

        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('所有差异', cell_overwrite_ok=True)
        output_path = "对比结果_" + str(self.count) + ".xls"
        if 1 not in self.excel_compare.dic or self.excel_compare.dic[1] == "":
            self.info_old()
        elif 2 not in self.excel_compare.dic or self.excel_compare.dic[2] == "":
            self.info_new()
        else:
            # 1. 读取A版、B版两张Excel表格
            excel_old = self.excel_compare.get_info(self.excel_compare.dic[1])
            excel_new = self.excel_compare.get_info(self.excel_compare.dic[2])

            # 2. 获取所有sheet及索引，形成字典
            try:
                name_old = [str(excel_old.sheet_by_name(x).row(20)[9].value) for x in excel_old.sheet_names()]
                name_new = [str(excel_new.sheet_by_name(x).row(20)[9].value) for x in excel_new.sheet_names()]

                name_set_old = set(name_old)
                name_set_new = set(name_new)
                self.excel_compare.compare_components(name_set_old, name_set_new)

                dic_old = self.excel_compare.generate_dic(name_set_old, name_old)
                dic_new = self.excel_compare.generate_dic(name_set_new, name_new)

                print()
                print(dic_old)
                print(dic_new, "\n")

                # 第一遍, 为了获取最大列宽
                self.excel_write.output_head(ws)
                start_row = 2
                for name in dic_old:
                    if name in dic_new:
                        mat_old = self.excel_compare.generate_mat_complete(dic_old[name], excel_old)
                        mat_new = self.excel_compare.generate_mat_complete(dic_new[name], excel_new)

                        c_mat_old = self.excel_compare.compress_mat(mat_old)
                        c_mat_new = self.excel_compare.compress_mat(mat_new)

                        change_res, delete_res, add_res = self.excel_compare.output(excel_old, excel_new, dic_old,
                                                                                    dic_new,
                                                                                    c_mat_old, c_mat_new, name)
                        start_row = self.excel_write.output_excel(ws, wb, output_path, change_res, delete_res,
                                                                  add_res, name, start_row)

                max_list = self.excel_write.adjust_col(output_path)

                # 第二遍, 带上列宽和格式写入
                self.excel_write.output_head(ws)
                start_row = 2
                for name in dic_old:
                    if name in dic_new:
                        mat_old = self.excel_compare.generate_mat_complete(dic_old[name], excel_old)
                        mat_new = self.excel_compare.generate_mat_complete(dic_new[name], excel_new)

                        c_mat_old = self.excel_compare.compress_mat(mat_old)
                        c_mat_new = self.excel_compare.compress_mat(mat_new)

                        change_res, delete_res, add_res = self.excel_compare.output(excel_old, excel_new, dic_old,
                                                                                    dic_new,
                                                                                    c_mat_old, c_mat_new, name)
                        start_row = self.excel_write.output_excel(ws, wb, output_path, change_res, delete_res,
                                                                  add_res, name, start_row,
                                                                  max_list)
                ws.write(start_row + 1, 0, label=self.excel_compare.compare_components(name_set_old, name_set_new),
                         style=self.excel_write.style_note)
                os.remove(output_path)
                dirpath = QFileDialog.getSaveFileName(self, "选择保存目录", output_path, "xls(*.xls);;xlsx(*.xlsx)")
                if dirpath[0] != '':
                    wb.save(dirpath[0])
                    self.count += 1
                    self.info_success(dirpath[0])
            except:
                self.info_wrong_file()

    def clear(self):
        """一键清空文件"""
        self.excel_compare.dic = {}
        self.label_2.setText("旧版文件名")
        self.label_3.setText("新版文件名")

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Excel文件差异对比"))
        MainWindow.setWindowIcon(QtGui.QIcon("Icon.ico"))
        file = QtCore.QFile('css.qss')
        file.open(QtCore.QFile.ReadOnly)
        stylesheet = file.readAll()
        QtWidgets.qApp.setStyleSheet(str(stylesheet, encoding='utf-8'))
        self.label.setText(_translate("MainWindow",
                                      "<html><head/><body><p><span style=\" font-size:12pt; font-weight:600; color:#000000;\">点击第一个按钮选择旧版文件，点击第二个按钮选择新版文件</span></p><p><span style=\" font-size:12pt; font-weight:600; color:#000000;\">本软件将基于旧版文件进行比较</span></p></body></html>"))
        self.pushButton.setText(_translate("MainWindow", "选择旧版文件"))
        self.pushButton_2.setText(_translate("MainWindow", "选择新版文件"))
        self.pushButton_3.setText(_translate("MainWindow", "对比导出文件"))
        self.pushButton_4.setText(_translate("MainWindow", "一键清空文件"))
        self.label_2.setText(_translate("MainWindow", "旧版文件名"))
        self.label_3.setText(_translate("MainWindow", "新版文件名"))


class MyWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)
        self.setFixedSize(self.width(), self.height())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWin = MyWindow()
    palette = QtGui.QPalette()
    palette.setBrush(QtGui.QPalette.Background, QtGui.QBrush(QtGui.QPixmap("background.jpg")))
    myWin.setPalette(palette)
    myWin.show()
    sys.exit(app.exec_())
