import sys
import os
import time
import pandas as pd
from datetime import datetime
import pyvisa
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
class MSO64Controller:
    def __init__(self):
        self.rm = pyvisa.ResourceManager()
        self.scope = None
        self.connected = False
        self.save_path = ""  # 基本保存路径
        self.image_path = ""  # 图片保存路径
        self.excel_path = ""  # Excel保存路径
        
    def connect(self, ip_address):
        """连接到示波器"""
        try:
            resource_string = f"TCPIP::{ip_address}::INSTR"
            self.scope = self.rm.open_resource(resource_string)
            self.scope.timeout = 10000  # 设置超时时间为10秒
            self.scope.write_termination = '\n'
            self.scope.read_termination = '\n'
            idn = self.scope.query("*IDN?")
            if "MSO64" in idn:
                self.connected = True
                return True, idn
            else:
                return False, f"发现仪器但不是MSO64: {idn}"
        except Exception as e:
            return False, str(e)
        
    def clear_measurements(self):
        """清除所有测量项"""
        if not self.connected:
            return False, "未连接到示波器"
        
        try:
            self.scope.write("MEASUREMENT:DELETEALL")
            return True, "测量项已清除"
        except Exception as e:
            return False, f"清除测量项失败: {str(e)}"

    def add_voltage_measurements(self):
        """添加电压测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        
        try:
            # 清除所有现有测量
            self.clear_measurements()
            
            # 添加电压参数
            self.scope.write("MEASUREMENT:ADDMEAS MAXIMUM")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE CH1")
            
            self.scope.write("MEASUREMENT:ADDMEAS MINIMUM")
            self.scope.write("MEASUREMENT:MEAS2:SOURCE CH1")
            
            self.scope.write("MEASUREMENT:ADDMEAS TOP")
            self.scope.write("MEASUREMENT:MEAS3:SOURCE CH1")
            
            self.scope.write("MEASUREMENT:ADDMEAS BASE")
            self.scope.write("MEASUREMENT:MEAS4:SOURCE CH1")
            
            # 添加CH2的电压参数
            self.scope.write("MEASUREMENT:ADDMEAS MAXIMUM")
            self.scope.write("MEASUREMENT:MEAS5:SOURCE CH2")
            
            self.scope.write("MEASUREMENT:ADDMEAS MINIMUM")
            self.scope.write("MEASUREMENT:MEAS6:SOURCE CH2")
            
            self.scope.write("MEASUREMENT:ADDMEAS TOP")
            self.scope.write("MEASUREMENT:MEAS7:SOURCE CH2")
            
            self.scope.write("MEASUREMENT:ADDMEAS BASE")
            self.scope.write("MEASUREMENT:MEAS8:SOURCE CH2")
            
            return True, "电压测量项添加成功"
        except Exception as e:
            return False, f"添加电压测量项失败: {str(e)}"
    def get_voltage_measurements(self):
        """获取电压测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            
            # CH1电压参数
            results["CH1_Maximum"] = float(self.scope.query("MEASUREMENT:MEAS1:RESULT:ACTUAL?"))
            results["CH1_Minimum"] = float(self.scope.query("MEASUREMENT:MEAS2:RESULT:ACTUAL?"))
            results["CH1_Top"] = float(self.scope.query("MEASUREMENT:MEAS3:RESULT:ACTUAL?"))
            results["CH1_Base"] = float(self.scope.query("MEASUREMENT:MEAS4:RESULT:ACTUAL?"))
            
            # CH2电压参数
            results["CH2_Maximum"] = float(self.scope.query("MEASUREMENT:MEAS5:RESULT:ACTUAL?"))
            results["CH2_Minimum"] = float(self.scope.query("MEASUREMENT:MEAS6:RESULT:ACTUAL?"))
            results["CH2_Top"] = float(self.scope.query("MEASUREMENT:MEAS7:RESULT:ACTUAL?"))
            results["CH2_Base"] = float(self.scope.query("MEASUREMENT:MEAS8:RESULT:ACTUAL?"))
            
            return True, "电压测量值获取成功", results
        except Exception as e:
            return False, f"获取电压测量值失败: {str(e)}", {}
    def add_frequency_measurements(self):
        """添加频率和脉宽测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        
        try:
            # 清除所有现有测量
            self.clear_measurements()
            
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE CH1")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 30")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 30")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 30")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 30")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 30")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 30")
            
            # 正脉宽 (从70%到70%)
            self.scope.write("MEASUREMENT:ADDMEAS PWIDTH")
            self.scope.write("MEASUREMENT:MEAS2:SOURCE CH1")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS2:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS2:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS2:REFLevels:PERCent:RISEHigh 70")
            self.scope.write("MEASUREMENT:MEAS2:REFLevels:PERCent:RISEMid 70")
            self.scope.write("MEASUREMENT:MEAS2:REFLevels:PERCent:RISELow 70")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS2:REFLevels:PERCent:FALLHigh 70")
            self.scope.write("MEASUREMENT:MEAS2:REFLevels:PERCent:FALLMid 70")
            self.scope.write("MEASUREMENT:MEAS2:REFLevels:PERCent:FALLLow 70")
            
            # 负脉宽 (从30%到30%)
            self.scope.write("MEASUREMENT:ADDMEAS NWIDTH")
            self.scope.write("MEASUREMENT:MEAS3:SOURCE CH1")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS3:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS3:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS3:REFLevels:PERCent:RISEHigh 30")
            self.scope.write("MEASUREMENT:MEAS3:REFLevels:PERCent:RISEMid 30")
            self.scope.write("MEASUREMENT:MEAS3:REFLevels:PERCent:RISELow 30")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS3:REFLevels:PERCent:FALLHigh 30")
            self.scope.write("MEASUREMENT:MEAS3:REFLevels:PERCent:FALLMid 30")
            self.scope.write("MEASUREMENT:MEAS3:REFLevels:PERCent:FALLLow 30")
            
            return True, "频率和脉宽测量项添加成功"
        except Exception as e:
            return False, f"添加频率和脉宽测量项失败: {str(e)}"
    
    def get_frequency_measurements(self):
        """获取频率和脉宽测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            
            results["Frequency"] = float(self.scope.query("MEASUREMENT:MEAS1:RESULT:ACTUAL?"))
            results["Positive_Pulse_Width"] = float(self.scope.query("MEASUREMENT:MEAS2:RESULT:ACTUAL?"))
            results["Negative_Pulse_Width"] = float(self.scope.query("MEASUREMENT:MEAS3:RESULT:ACTUAL?"))
            
            return True, "频率和脉宽测量值获取成功", results
        except Exception as e:
            return False, f"获取频率和脉宽测量值失败: {str(e)}", {}
    
    def get_delay_measurement(self):
        """获取Setup Time测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            
            results["Delay_Rise_Edge_70_30"] = float(self.scope.query("MEASUREMENT:MEAS1:RESULT:ACTUAL?"))
            
            return True, "Setup Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Setup Time测量值失败: {str(e)}", {}
    
    def add_delay_measurement(self):
        """添加延迟测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        
        try:
            # 清除所有现有测量
            self.clear_measurements()
            
            # setup time (从CH2上升沿70%到CH1上升沿30%)
            self.scope.write("MEASUREMENT:ADDMEAS DELAY")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE1 CH2")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE2 CH1")
            
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            
            # 设置CH2(SOURCE1)的参考电平为70%
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 90")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 70")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 10")
            
            # 设置CH1(SOURCE2)的参考电平为30%
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISE2High 90")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISE2Mid 30")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISE2Low 10")
            
            # 设置测量的边沿和方向
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION1 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION2 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE1 RISE")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE2 RISE")
            
            return True, "setup time测量项添加成功"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def get_delay_measurement(self):
        """获取Setup Time测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            
            results["Delay_Rise_Edge_70_30"] = float(self.scope.query("MEASUREMENT:MEAS1:RESULT:ACTUAL?"))
            
            return True, "Setup Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Setup Time测量值失败: {str(e)}", {}

controller=MSO64Controller()
suc, msg=controller.connect('169.254.103.178')
controller.add_voltage_measurements()
controller.scope.write('SAVe:EVENTtable:MEASUrement "C:\\ST\\hi.csv"')
controller.scope.write('SAVe:IMAGe "C:\\ST\\MyScreenFrame.png"')
controller.scope.query("*OPC?")
controller.scope.write('FILESYSTEM:READFILE "C:\\ST\\MyScreenFrame.png"')
raw_data=controller.scope.read_raw()
fid=open("MyScreenFrame.png", 'wb')
fid.write(raw_data)
fid.close()
controller.scope.write('FILESYSTEM:READFILE "C:\\ST\\hi.csv"')
data=controller.scope.read()
#print(data)
lines = data.strip().splitlines()
print(lines[-1])
