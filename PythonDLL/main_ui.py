from PyQt5 import QtWidgets, QtCore
from main_face2 import Ui_MainWindow
import PythonDLL3X as device
import sys
import os



sys.path.append('../')


# sys.path.append(os.path.abspath(os.path.dirname(__file__) + '/' + '..'))


class EventMessageBox(QtWidgets.QMessageBox):
    def __init__(self, timeout=3, parent=None, msg="", event=None, param=None):
        super(EventMessageBox, self).__init__(parent)
        self.setWindowTitle("Scanning…………")
        # self.showEvent()
        self.setStandardButtons(QtWidgets.QMessageBox.NoButton)
        # self.showinfo("Scanning…………")

    def closeEvent(self, event):
        event.accept()


# def createFolder(directory):
#     try:
#         if not os.path.exists(directory):
#             os.makedirs(directory)
#     except OSError:
#         print('Error: Creating directory. ' + directory)

def createFolder(directory):
    try:
        path = r'D:/Pukon/'
        if not os.path.exists(directory):
            os.makedirs(path + directory, exist_ok=True)
    except OSError:
        print('Error: Creating directory. ' + directory)


class MainUI(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        self.device = device.Device()
        self.setupUi(self)  # 连接界面
        self.init_ui()  # 创建文件
        # self.contrl_com_enable = 0
        # self.moto_com_enable = 0
        # self.COM_ControlBoard = str(self.comboBox_Control_COM.currentText())
        # self.COM_Moto = str(self.comboBox_Moto_COM.currentText())
        # self.comboBox_Control_COM_indexChange()
        self.connectionTimer = QtCore.QTimer()
        self.connectionTimer.setInterval(1000)
        self.connectionTimer.timeout.connect(self.Button_Start)
        # self.connectionTime = QtCore.QTimer()
        # self.connectionTime.setInterval(2000)
        # self.connectionTime.timeout.connect(self.SetUiInfo)
        # self.connectionTime.start()
        self.time_cnt = 0
        self.message = EventMessageBox()
        '''
        self.ScanTimer = QtCore.QTimer()
        self.ScanTimer.setInterval(5000)
        self.ScanTimer.timeout.connect(self.pushButton_scan_one)
        self.ScanTimer.start()
        '''

    def create_Folter(self):  # 创建文件目录
        createFolder("Pukon")
        # for i in range(0, len(self.device.LoadConfig.Pre_Function)):
        #     if self.device.LoadConfig.Pre_Function[i] != '':
        #         folder_name = "Pukon_" + str(self.device.LoadConfig.Pre_Function[i])
        #         createFolder(folder_name)
        folder_name = "Pukon_" + str("原始光谱")
        createFolder(folder_name)
        folder_name = "Pukon_" + str("透射光谱")
        createFolder(folder_name)
        folder_name = "Pukon_" + str("吸收光谱")
        createFolder(folder_name)
        folder_name = "Pukon_" + str("吸收预处理")
        createFolder(folder_name)
        folder_name = "Pukon_" + str("理化指标")
        createFolder(folder_name)
        createFolder("Pukon_" + str("环境温度"))

    def init_ui(self):  # 设计窗口标题，按钮功能
        self.setWindowTitle("中科谱康 在线水分检测仪 API   V1.0.0")
        # self.comboBox_Control_COM.currentIndexChanged.connect(self.comboBox_Control_COM_indexChange)
        # self.comboBox_Moto_COM.currentIndexChanged.connect(self.comboBox_Moto_COM_indexChange)
        self.pushButton_Stop.clicked.connect(self.Button_Stop)
        self.pushButton_Start.clicked.connect(self.Button_Start)
        self.pushButton_back.clicked.connect(self.device.Button_back)
        # Initialize
        # time.sleep(1)
        self.SetUiInfo()
        self.create_Folter()

    def SetUiInfo(self):
        if self.device.IsConnected():
            # self.label_connected.setText('Connected!')
            self.label_6.setStyleSheet(
                "min-width:20px;min-height:20px;max-width:20px;max-height:20px;border-radius:10px;border:1px solid black;background:green")
        else:
            self.label_6.setStyleSheet(
                "min-width:20px;min-height:20px;max-width:20px;max-height:20px;border-radius:10px;border:1px solid black;background:red")
        #
        if self.device.IsCOM_Connected():
            self.label_7.setStyleSheet(
                "min-width:20px;min-height:20px;max-width:20px;max-height:20px;border-radius:10px;border:1px solid black;background:green")
        else:
            self.label_7.setStyleSheet(
                "min-width:20px;min-height:20px;max-width:20px;max-height:20px;border-radius:10px;border:1px solid black;background:red")
    '''
    def comboBox_Control_COM_indexChange(self):
        if (self.contrl_com_enable == 1):
            self.contrl_com.close()
        try:
            self.COM_ControlBoard = self.comboBox_Control_COM.currentText()
            self.contrl_com = serial.Serial(self.COM_ControlBoard, 9600, timeout=1)
            self.label_COM_connected.setText(' ')
            self.contrl_com_enable = 1
        except:
            self.contrl_com_enable = 0
            self.label_COM_connected.setText('串口打开失败!')

    def Contrl_Communication(self, data):
        if (self.contrl_com_enable == 1):
            self.contrl_com.write(data.encode())
            print(data)
            self.label_COM_connected.setText(' ')
            try:
                receive_frame = self.contrl_com.readline()
                print(receive_frame)
                if receive_frame == b'WOK\r\n':
                    print(type(receive_frame))
                    print(receive_frame)
                # else:
                # QMessageBox.information(self,  "提示", "数据接收错误！")
            except:
                print('error')

            byte_number_1 = 0
            tmp_cnt = 0
            while True:
                byte_number_1 = self.contrl_com.inWaiting()
                print(byte_number_1)
                time.sleep(1)
                tmp_cnt += 1
                if tmp_cnt == 10:
                    tmp_cnt = 0
                    break
                if byte_number_1 != 0:
                    receive_frame = self.contrl_com.readline()
                    print(receive_frame)

        else:
            self.label_COM_connected.setText('串口打开失败!')

    def comboBox_Moto_COM_indexChange(self):
        if (self.moto_com_enable == 1):
            self.moto_com.close()
        try:
            self.COM_Moto = self.comboBox_Moto_COM.currentText()
            self.moto_com = serial.Serial(self.COM_Moto, 9600)
            self.moto_com_enable = 1
        except:
            self.moto_com_enable = 0
            self.label_COM2_connected.setText('串口打开失败!')

    def MOTO_Communication(self, data):
        if (self.moto_com_enable == 1):
            self.moto_com.write(data.encode())
            self.label_COM2_connected.setText(' ')
        else:
            self.label_COM2_connected.setText('串口打开失败!')
    '''
    def _timeout(self):
        # self.SetUiInfo()
        self.time_cnt += 1
        print(self.time_cnt)
        # print(int(self.device.LoadConfig.LoopTime))
        # if self.time_cnt == 5:
        if self.time_cnt == int(self.device.LoadConfig.LoopTime):
            # self.connectionTime.stop()
            self.SetUiInfo()
            self.message.show()
            self_WaterData = self.device.WriteDataToExcl()
            self_WaterData = '%.2f' % (int(self_WaterData[1] * 100) * 0.01)
            self.message.done(1)
            self.time_cnt = 0
            # time.sleep(0.1)
            # self.connectionTime.start()
            # string_temp = 'W' + str(self_WaterData) + '\r' + '\n'
            # self.Contrl_Communication(string_temp)
            self.lineEdit_Water.setText(self_WaterData)

    def Button_Stop(self):
        self.connectionTimer.stop()
        self.lineEdit_Statue.setText('停止检测')
        # self.connectionTime.stop()



    def Button_Start(self):
        self.connectionTimer.start()
        self.lineEdit_Statue.setText('正在检测')
        self._timeout()

