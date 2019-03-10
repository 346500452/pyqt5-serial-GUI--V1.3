import sys

from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

import xlrd
import xlwt

import serial
import serial.tools.list_ports

from demoSerial1 import Ui_Form

 

class mainWindow(QMainWindow, Ui_Form):
    ser = serial.Serial()
    receiveCnt = 0
    sendCnt = 0
    def __init__(self, parent=None):
        super(mainWindow, self).__init__(parent)

        self.setupUi(self)
        #禁止拉伸窗口大小
        self.setFixedSize(self.width(), self.height())
        
        self.setWindowIcon(QIcon('./images/cartoon.ico'))
        
        tmp = '发送:' + '{:d}'.format(self.sendCnt) + ' 接收:' + '{:d}'.format(self.receiveCnt)
        self.cntLabel.setText(tmp)
        #信号与槽函数的连接
        self.initUi()
        
        #receiveThread = workThread()
        #定时器调用读取串口接收数据
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.receiveData)
        #定时发送
        self.timerSend = QTimer(self)
        self.timerSend.timeout.connect(self.send1)
        #显示时间
        self.timerShow = QTimer(self)
        self.timerShow.timeout.connect(self.showTime)
        self.timerShow.start(1000)
        
        #自动获取当前可用串口
        self.serialCOMComboBox.addItems(self.serialList())
        
    def initUi(self):
        #设置默认参数
        self.serialBaudRateComboBox.setCurrentIndex(2)
        self.serialDataBitComboBox.setCurrentIndex(3)
        self.serialStopBitComboBox.setCurrentIndex(0)
        self.serialParityComboBox.setCurrentIndex(0)
        #连接打开串口按钮，并初始化可用
        self.openSerialButton.clicked.connect(self.serialOpen)
        self.openSerialButton.setEnabled(True)
        #连接关闭串口按钮，并初始化不可用
        self.closeSerialButton.clicked.connect(self.serialClose)
        self.closeSerialButton.setEnabled(False)
        #刷新当前可用串口，并刷新com列表显示
        self.refreshSerialButton.clicked.connect(self.serialRefresh)
        #清除接收区
        self.clearButtonReceive.clicked.connect(self.receiveClear)
        #清除发送区
        self.clearButtonSend.clicked.connect(self.sendClear)
        #发送信息
        #单条发送
        self.sendButton.clicked.connect(self.send1)
        #多条发送
        self.sendButton0.clicked.connect(self.send2)
        self.sendButton1.clicked.connect(self.send3)
        self.sendButton2.clicked.connect(self.send4)
        self.sendButton3.clicked.connect(self.send5)
        self.sendButton4.clicked.connect(self.send6)
        self.sendButton5.clicked.connect(self.send7)
        self.sendButton6.clicked.connect(self.send8)
        self.sendButton7.clicked.connect(self.send9)
        
        #定时发送
        self.timingSendCheckBox.stateChanged.connect(self.timerSendBox)
        
        #导入文件
        self.importButton.clicked.connect(self.importFile)
        
        #导出文件
        self.outportButton.clicked.connect(self.outportFile)
    
    #扫描获取端口号列表，并以列表形式返回
    def serialList(self):
        comList = []
        #获取当前可用串口信息
        portList = list(serial.tools.list_ports.comports())
        portList.sort()
        for port in portList:
            comList.append(port[0])
            '''
            #将串口信息格式重新调整，便于使用
            str = port[1]
            str = str[:-7]
            str = '%s:%s' % (port[0], str)
            #将调整后的串口信息加入列表中
            comList.append(str)
            #print(str)
            '''
        return comList
        
    def serialRefresh(self):
        self.serialCOMComboBox.clear()
        #自动获取当前可用串口
        self.serialCOMComboBox.addItems(self.serialList())
    
    def serialOpen(self):
        try:
            self.ser.port = self.serialCOMComboBox.currentText()
            self.ser.baudRate = self.serialBaudRateComboBox.currentText()
            self.ser.dataBits = self.serialDataBitComboBox.currentText()
            parityValue = self.serialParityComboBox.currentText()
            self.ser.parity = parityValue[0]
            self.ser.stopBits = int(self.serialStopBitComboBox.currentText())
            
            self.ser.open()
            #定时器开启，2ms
            self.timer.start(10)
        except:
            QMessageBox.critical(self, '错误提示','打开串口失败!!!\r\n没有可用的串口或当前串口被占用')
            return None
        
        if self.ser.isOpen():
            print("串口打开成功")
            self.openSerialButton.setEnabled(False)
            self.closeSerialButton.setEnabled(True)
        else:
            print("串口打开失败")
        
    def serialClose(self):
        self.receiveCnt = 0
        self.sendCnt = 0
        #定时器关闭
        if self.timer.isActive():
            self.timer.stop()
        if self.timerSend.isActive():
            self.timerSend.stop()
        try:
            self.ser.close()
        except:
            QMessageBox.critical(self, '错误提示','关闭串口失败!!!')
            return None
        if self.ser.isOpen():
            print("串口关闭失败")
        else:
            print("串口关闭成功")
            self.openSerialButton.setEnabled(True)
            self.closeSerialButton.setEnabled(False)
            
    def receiveClear(self):
        self.receiveTextEdit.clear()
        print("清除接收内容成功！")
        self.sendCnt = 0
        self.receiveCnt = 0
        tmp = '发送:' + '{:d}'.format(self.sendCnt) + ' 接收:' + '{:d}'.format(self.receiveCnt)
        self.cntLabel.setText(tmp)
        
    def sendClear(self):
        self.sendTextEdit.clear()
        print("清除发送内容成功！")
        
    def timerSendBox(self):
        if self.ser.isOpen():
            if self.timingSendCheckBox.checkState():
                time = self.timingLineEdit.text()
                try:
                    timeValue = int(time, 10)
                except ValueError:
                    QMessageBox.critical(self, '错误提示','请输入有效的定时时间！')
                    print("请输入有效的定时时间！")
                    return None
                if timeValue <= 0:
                    QMessageBox.critical(self, '错误提示','定时时间必须大于0！')
                    print("定时时间必须大于0！")
                    return None
                #开启定时发送定时器
                self.timerSend.start(timeValue)
            else:
                #定时器关闭
                if self.timerSend.isActive():
                    self.timerSend.stop()
        else:
            #定时器关闭
            if self.timerSend.isActive():
                self.timerSend.stop()
        
    def sendData(self, ch):
        #有串口打开，才进行发送操作
        if self.ser.isOpen():
            if ch == 1:
                data = self.sendTextEdit.toPlainText()
            elif ch == 2:
                data = self.sendLineEdit0.text()
            elif ch == 3:
                data = self.sendLineEdit1.text()
            elif ch == 4:
                data = self.sendLineEdit2.text()
            elif ch == 5:
                data = self.sendLineEdit3.text()
            elif ch == 6:
                data = self.sendLineEdit4.text()
            elif ch == 7:
                data = self.sendLineEdit5.text()
            elif ch == 8:
                data = self.sendLineEdit6.text()
            elif ch == 9:
                data = self.sendLineEdit7.text()
                
            if data != '':
                #字符发送
                if self.hexSendCheckBox.checkState() == False:
                    #发送新行
                    if self.sendNewLineCheckBox.checkState():
                        data = data + '\r\n'
                    data = data.encode('gbk')
                #十六进制发送
                else:
                    data = data.strip()     #删除前后空格
                    send_list = []
                    while data != '':
                        try:
                            num = int(data[0:2], 16)
                        except ValueError:
                            QMessageBox.critical(self, '错误提示','请输入十六进制数据，以空格分开！')
                            print("请输入十六进制数据，以空格分开！")
                            return None
                        data = data[2:]
                        data = data.strip()
                        #添加到发送列表中
                        send_list.append(num)
                    data = bytes(send_list)
                    #data = data.encode('utf-8')
                #发送数据
                try:
                    self.ser.write(data)
                    print("发送成功")
                except:
                    print("发送失败!!!\r\n没有可用的串口或当前串口被占用")
                    QMessageBox.critical(self, '错误提示','发送失败!!!\r\n没有可用的串口或当前串口被占用')
                    return None
                #显示统计信息
                self.sendCnt = self.sendCnt + len(data)
                tmp = '发送:' + '{:d}'.format(self.sendCnt) + ' 接收:' + '{:d}'.format(self.receiveCnt)
                self.cntLabel.setText(tmp)
            else:
                print("无数据输入")
                QMessageBox.warning(self, '警告提示','无数据输入!')
        else:
            QMessageBox.critical(self, '错误提示','发送失败!!!\r\n没有可用的串口或当前串口被占用')
            return None
            
    def send1(self):
        self.sendData(1)
        
    def send2(self):
        self.sendData(2)
        
    def send3(self):
        self.sendData(3)
        
    def send4(self):
        self.sendData(4)
        
    def send5(self):
        self.sendData(5)
        
    def send6(self):
        self.sendData(6)
        
    def send7(self):
        self.sendData(7)
        
    def send8(self):
        self.sendData(8)
        
    def send9(self):
        self.sendData(9)
            
    def receiveData(self):
        if self.ser.isOpen():
            num = self.ser.inWaiting()            
            if num > 0:
                print(num)
                bytes = self.ser.read(self.ser.inWaiting() )
                #bytes = self.ser.readline()
                for i in range(0, len(bytes)):
                    print(hex(bytes[i]))
                #十六进制显示
                if self.hexShowCheckBox.checkState():
                    showData = ''
                    for i in range(0, len(bytes)):
                        showData = showData + '{:02x}'.format(bytes[i]) + ' '
                else:
                    showData = ''
#                    while len(bytes) > 1:
#                        if bytes[0] < 0x7f:
#                            showData += chr(bytes[0])
#                            bytes = bytes[1:]
#                        else:
#                            try:
#                                hanzi = bytes[:2].decode('gbk')#GB18030
#                                showData += hanzi
#                                bytes = bytes[2:]
#                            except Exception as e:
#                                showData += '{:02x}'.format(bytes[0])#'\\x%02X' % int(bytes[0])
#                                bytes = bytes[1:]
#                    if len(bytes) > 0:
#                        if bytes[0] < 0x7f:
#                            showData += chr(bytes[0])
#                            bytes = bytes[1:]
                    try:
                        showData = bytes.decode('utf-8')
                        print(showData)
                        #showData = unicode(QtCore.QString(data)).decode('utf-8')
                    except:
                        print("编码出错")
                        for i in range(0, len(bytes)):
                            showData = showData + '{:02x}'.format(bytes[i]) + ' '
                    
                #先把光标移到最后
                cursor = self.receiveTextEdit.textCursor()
                if cursor != cursor.End:
                    cursor.movePosition(cursor.End)
                    self.receiveTextEdit.setTextCursor(cursor)
                #把字符显示到窗口中
                self.receiveTextEdit.insertPlainText(showData)
                #显示统计信息
                self.receiveCnt = self.receiveCnt + num
                tmp = '发送:' + '{:d}'.format(self.sendCnt) + ' 接收:' + '{:d}'.format(self.receiveCnt)
                self.cntLabel.setText(tmp)
                #最后再将光标移到最后
                cursor = self.receiveTextEdit.textCursor()
                cursor.movePosition(cursor.End)
                self.receiveTextEdit.setTextCursor(cursor)
            else:
                pass
        else:
            self.serialClose()
            print("接收出错")
            QMessageBox.critical(self, '错误提示','串口被拔出')
            
    def importFile(self):
        fileName, _ = QFileDialog().getOpenFileName(self, 'Open file', '.', "excl file(*.xls)")
        #打开Excel文件读取数据
        data = xlrd.open_workbook(fileName)
        #获取一个工作表
        table = data.sheets()[0] 
        #获取行数
        row = table.nrows
        #print(row)
        #获取单元格,并填充到相应位置
        for i in range(0, row):
            value = table.cell(i,0).value
            #print(i, value)
            if i == 0:
                self.sendLineEdit0.setText(str(value))
            elif i == 1:
                self.sendLineEdit1.setText(str(value))
            elif i == 2:
                self.sendLineEdit2.setText(str(value))
            elif i == 3:
                self.sendLineEdit3.setText(str(value))
            elif i == 4:
                self.sendLineEdit4.setText(str(value))
            elif i == 5:
                self.sendLineEdit5.setText(str(value))
            elif i == 6:
                self.sendLineEdit6.setText(str(value))
            elif i == 7:
                self.sendLineEdit7.setText(str(value))
    
    def outportFile(self):
        wb = xlwt.Workbook()
        ws = wb.add_sheet('Sheet1')
        for i in range(0, 8):
            if i == 0:
                value = self.sendLineEdit0.text()
            elif i == 1:
                value = self.sendLineEdit1.text()
            elif i == 2:
                value = self.sendLineEdit2.text()
            elif i == 3:
                value = self.sendLineEdit3.text()
            elif i == 4:
                value = self.sendLineEdit4.text()
            elif i == 5:
                value = self.sendLineEdit5.text()
            elif i == 6:
                value = self.sendLineEdit6.text()
            elif i == 7:
                value = self.sendLineEdit7.text()
            ws.write(i, 0, value)
            #print(i, value)
        #保存,文件名为当天日期
        exclName = QDate.currentDate()
        exclName = exclName.toString("yyyy-MM-dd.xls")
        wb.save(exclName)
        QMessageBox.about(self, '提示','导出成功！')
        
    def showTime(self):
        date = QDateTime.currentDateTime()
        date = date.toString("yyyy-MM-dd hh:mm:ss ddd")
        self.timeLabel.setText(date)
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = mainWindow()
    win.show()
    sys.exit(app.exec_())
