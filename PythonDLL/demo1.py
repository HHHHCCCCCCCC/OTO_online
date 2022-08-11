import math
import serial
from ctypes import *
import openpyxl
import datetime
import pandas as pd
import numpy as np
import base64

class Device:
    def __init__(self):
        self.VID = 1592
        self.PID = 2732
        # self.pre = pre()
        self.errStatus = 0
        self.contrl_com_enable = 0
        self.OTOdll = CDLL("UserApplication.dll")
        # self.LoadConfig = IRConfig()
        self.intFramesize = self.GetFramesize()
        # self.Temperature = self.TEC()
        self.IsCOM_Communication()

        # print(self.LoadConfig.Head_Information)

    def IsConnected(self):
        #Check how many device is connected with PC.
        self.intDeviceamout = c_int(0)
        self.OTOdll.UAI_SpectrometerGetDeviceAmount(self.VID,self.PID,byref(self.intDeviceamout))
        return self.intDeviceamout.value

    def IsCOM_Connected(self):
        if self.contrl_com_enable == 1:
            return 1
        else:
            return 0

    def IsCOM_Communication(self):
        if self.contrl_com_enable == 1:
            count = self.contrl_com.inWaiting()                #获取串口缓冲区数据
            self.contrl_com.flushInput()  # 清空缓冲区
            if count != 0:
                self.receive_frame = self.contrl_com.readline(count)
                # print(self.receive_frame)
                # self.contrl_com.flushInput()
        else:
            try:
                self.COM_Control = (self.LoadConfig.comnumber)
                self.contrl_com = serial.Serial(self.COM_Control , 9600, timeout=1)
                self.receive_frame = self.contrl_com.readline()
                self.contrl_com_enable = 1
            except:
                self.contrl_com_enable = 0
                return 0
        return self.receive_frame
                # self.label_COM_connected.setText('串口打开失败!')
    def Contrl_Communication(self, data):
        if(self.contrl_com_enable == 1):
            self.contrl_com.write(data.encode())
            # print(data)
        else:
            self.contrl_com_enable = 0
            # QMessageBox.information(self,  "提示", "数据接收错误！")
            # self.IsCOM_Connected()


    def Open(self):
        #Open Device
        self.DeviceHandle = c_int(0)
        self.OTOdll.UAI_SpectrometerOpen(0,byref(self.DeviceHandle),self.VID,self.PID)
        return self.DeviceHandle.value

    def GetFramesize(self):
        self.Open()
        #Get Framesize
        self.intFramesize = c_int(0)
        self.OTOdll.UAI_SpectromoduleGetFrameSize(self.DeviceHandle,byref(self.intFramesize))
        print("Device framesize:")
        return self.intFramesize

    def GetModuleName(self):
        #Get Module name
        self.charModulename = create_string_buffer(16)
        self.OTOdll.UAI_SpectrometerGetModelName(self.DeviceHandle,byref(self.charModulename))
        # print("Module name:")
        return repr(self.charModulename.value)

    def GetSerilaNum(self):
        #Get Serial number
        self.charSerialnumber = create_string_buffer(16)
        self.OTOdll.UAI_SpectrometerGetSerialNumber(self.DeviceHandle,byref(self.charSerialnumber))
        print("Serial number:")
        return repr(self.charSerialnumber.value)

    def decode(self,ciphertext):
        plaintext = ''
        ciphertext = base64.b64decode(ciphertext)
        for i in ciphertext:
            s = ord(i)-16 ######################
            s = s^32
            plaintext += chr(s)
        return plaintext

# strData = str(a.value, encoding="utf-8")
# print("strdata",strdata)


a = Device().GetSerilaNum()
b = "a"

# Device().decode(b)
for i in range (16):
    print(ord(a[i]))

#
# cipher = 'XlNkVmtUI1MgXWBZXCFeKY+AaXNt'
# flag = Device().decode(cipher)
# print(flag)
