import sys
import os
import time
import pandas as pd
from datetime import datetime
import pyvisa
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
import math

def round_sig(x):
    return round(float(x), 3)


def adapt_time(x):
    if x < 1e-6:  # less than 1 microsecond (1e-6)
        x *= 1e9  # Convert seconds to nanoseconds
        return str(x)
    elif x < 1:  # between 1 microsecond and 1 second
        x *= 1e6  # Convert seconds to microseconds
        return str(x)
    else:  # 1 second or more
        return str(x)
    
def adapt_frequency(x):
    x=float(x)
    if x < 1e3:  # Less than 1 kHz
        return str(x)
    elif x < 1e6:  # Between 1 kHz and 1 MHz
        x /= 1e3  # Convert Hz to kHz
        return str(x)
    elif x < 1e9:  # Between 1 MHz and 1 GHz
        x /= 1e6  # Convert Hz to MHz
        return str(x)
    else:  # 1 GHz or more
        x /= 1e9  # Convert Hz to GHz
        return str(x)


class MSO64Controller:
    def __init__(self):
        self.rm = pyvisa.ResourceManager()
        self.scope = None
        self.connected = False
        self.save_path = ""  # 基本保存路径
        self.image_path = None  # 图片保存路径
        self.excel_path = ""  # Excel保存路径
        self.time=time.strftime("%Y-%m-%d %H-%M-%S")
        self.folder_name = "Test " + self.time
        self.upper=None
        self.lower=None
        self.upper_read=None
        self.lower_read=None
        self.channels={}

    def connect(self, ip_address):
        """连接到示波器"""
        try:
            resource_string = f"TCPIP::{ip_address}::INSTR"
            self.scope = self.rm.open_resource(resource_string)
            self.scope.timeout = 5000  # 设置超时时间为10秒
            idn = self.scope.query("*IDN?")
            self.scope.write('FILESystem:MKDir "C:\\Test"')
            #self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 0")
            #self.scope.write('FILESYSTEM:CWD "C:\\ST\\hsd\\i2c"')
            #print(self.scope.query("FILESystem:CWD?"))
            
            if "MSO64" in idn:
                self.connected = True
                return True, idn
            else:
                return False, f"发现仪器但不是MSO64: {idn}"
        except Exception as e:
            return False, str(e)
    
    def disconnect(self):
        """断开与示波器的连接"""
        if self.scope is not None:
            self.scope.close()
            self.connected = False
            self.scope = None
    
    def clear_measurements(self):
        """清除所有测量项"""
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.scope.write("MEASUREMENT:DELETEALL")
            self.scope.write("SEARCH:DELETEALL")
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
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            
            self.scope.write("MEASUREMENT:ADDMEAS MINIMUM")
            self.scope.write(f"MEASUREMENT:MEAS2:SOURCE {self.channels['SCL']}")
            
            self.scope.write("MEASUREMENT:ADDMEAS TOP")
            self.scope.write(f"MEASUREMENT:MEAS3:SOURCE {self.channels['SCL']}")
            
            self.scope.write("MEASUREMENT:ADDMEAS BASE")
            self.scope.write(f"MEASUREMENT:MEAS4:SOURCE {self.channels['SCL']}")
            
            # 添加CH2的电压参数
            self.scope.write("MEASUREMENT:ADDMEAS MAXIMUM")
            self.scope.write(f"MEASUREMENT:MEAS5:SOURCE {self.channels['SDA']}")
            
            self.scope.write("MEASUREMENT:ADDMEAS MINIMUM")
            self.scope.write(f"MEASUREMENT:MEAS6:SOURCE {self.channels['SDA']}")
            
            self.scope.write("MEASUREMENT:ADDMEAS TOP")
            self.scope.write(f"MEASUREMENT:MEAS7:SOURCE {self.channels['SDA']}")
            
            self.scope.write("MEASUREMENT:ADDMEAS BASE")
            self.scope.write(f"MEASUREMENT:MEAS8:SOURCE {self.channels['SDA']}")
            self.scope.query("*OPC?")
            return True, "电压测量项添加成功"
        except Exception as e:
            return False, f"添加电压测量项失败: {str(e)}"
        
    def get_voltage_measurements(self):
        """获取电压测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            #self.scope.write(f'FILESYSTEM:READFILE "C:\\Test\\Voltage{self.time}.txt"')
            #data=self.scope.read()
            #fid=open(self.folder_name+"\\Parameters\\Voltage.txt", 'w')
            #fid.write(data)
            #fid.close()
            #print(self.scope.query("MEASUREMENT:MEAS8:VALUE?"))
            #lines = data.strip().splitlines()
            """
            results["CH2_Base"]=lines[-1].split(',')[4]
            results["CH2_Top"]=lines[-3].split(',')[4]
            results["CH2_Minimum"]=lines[-5].split(',')[4]
            results["CH2_Maximum"]=lines[-7].split(',')[4]
            results["CH1_Base"]=lines[-9].split(',')[4]
            results["CH1_Top"]=lines[-11].split(',')[4]
            results["CH1_Minimum"]=lines[-13].split(',')[4]
            results["CH1_Maximum"]=lines[-15].split(',')[4]
            """
            results["SCL_Maximum"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?")))
            results["SCL_Minimum"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS2:VALUE?")))
            results["SCL_Top"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS3:VALUE?")))
            results["SCL_Base"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS4:VALUE?")))
            
            # CH2电压参数
            results["SDA_Maximum"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS5:VALUE?")))
            results["SDA_Minimum"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS6:VALUE?")))
            results["SDA_Top"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS7:VALUE?")))
            results["SDA_Base"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS8:VALUE?")))
            return True, "电压测量值获取成功", results
        except Exception as e:
            return False, f"获取电压测量值失败: {str(e)}", {}
        

    def voltage_measurement_by_channel(self, channel):
        #Channel is self.
        """添加电压测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        
        try:
            # 清除所有现有测量
            self.clear_measurements()
            # 添加电压参数
            self.scope.write("MEASUREMENT:ADDMEAS MAXIMUM")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {channel}")
            
            self.scope.write("MEASUREMENT:ADDMEAS MINIMUM")
            self.scope.write(f"MEASUREMENT:MEAS2:SOURCE {channel}")
            
            self.scope.write("MEASUREMENT:ADDMEAS TOP")
            self.scope.write(f"MEASUREMENT:MEAS3:SOURCE {channel}")
            
            self.scope.write("MEASUREMENT:ADDMEAS BASE")
            self.scope.write(f"MEASUREMENT:MEAS4:SOURCE {channel}")
            self.scope.query("*OPC?")
            #find channel name
            channel_name=""
            for keys in self.channels.keys():
                if self.channels[keys]==channel:
                    channel_name=keys
            results = {}
            results[f"{channel_name}_Maximum"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?")))
            results[f"{channel_name}_Minimum"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS2:VALUE?")))
            results[f"{channel_name}_Top"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS3:VALUE?")))
            results[f"{channel_name}_Base"] = round_sig(float(self.scope.query("MEASUREMENT:MEAS4:VALUE?")))
            print(results)
            return True, "电压测量项添加成功", results
        except Exception as e:
            return False, f"添加电压测量项失败: {str(e)}"
        
    
    
    def add_frequency_measurements(self, frequency=30, positive=70, negative=30, mode="Percentage"):
        """添加频率和脉宽测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        try:
            # 清除所有现有测量
            self.clear_measurements()
            method="PERCent" if mode=="Percentage" else "ABSolute"
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write(f"MEASUREMENT:MEAS1:REFLevels:METHod {method}")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write(f"MEASUREMENT:MEAS1:REFLevels:{method}:RISEHigh {frequency}")
            self.scope.write(f"MEASUREMENT:MEAS1:REFLevels:{method}:RISEMid {frequency}")
            self.scope.write(f"MEASUREMENT:MEAS1:REFLevels:{method}:RISELow {frequency}")
            # 设置下降沿参考电平
            self.scope.write(f"MEASUREMENT:MEAS1:REFLevels:{method}:FALLHigh {frequency}")
            self.scope.write(f"MEASUREMENT:MEAS1:REFLevels:{method}:FALLMid {frequency}")
            self.scope.write(f"MEASUREMENT:MEAS1:REFLevels:{method}:FALLLow {frequency}")
            
            # 正脉宽 (从70%到70%)
            self.scope.write("MEASUREMENT:ADDMEAS PWIDTH")
            self.scope.write(f"MEASUREMENT:MEAS2:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write(f"MEASUREMENT:MEAS2:REFLevels:METHod {method}")
            self.scope.write("MEASUrement:MEAS2:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write(f"MEASUREMENT:MEAS2:REFLevels:{method}:RISEHigh {positive}")
            self.scope.write(f"MEASUREMENT:MEAS2:REFLevels:{method}:RISEMid {positive}")
            self.scope.write(f"MEASUREMENT:MEAS2:REFLevels:{method}:RISELow {positive}")
            # 设置下降沿参考电平
            self.scope.write(f"MEASUREMENT:MEAS2:REFLevels:{method}:FALLHigh {positive}")
            self.scope.write(f"MEASUREMENT:MEAS2:REFLevels:{method}:FALLMid {positive}")
            self.scope.write(f"MEASUREMENT:MEAS2:REFLevels:{method}:FALLLow {positive}")
            
            # 负脉宽 (从30%到30%)
            self.scope.write("MEASUREMENT:ADDMEAS NWIDTH")
            self.scope.write(f"MEASUREMENT:MEAS3:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:METHod {method}")
            self.scope.write("MEASUrement:MEAS3:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:RISEHigh {negative}")
            self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:RISEMid {negative}")
            self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:RISELow {negative}")
            # 设置下降沿参考电平
            self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:FALLHigh {negative}")
            self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:FALLMid {negative}")
            self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:FALLLow {negative}")
            self.scope.query("*OPC?")
            return True, "频率和脉宽测量项添加成功"
        except Exception as e:
            return False, f"添加频率和脉宽测量项失败: {str(e)}"
        
    def get_frequency_measurements(self):
        """获取频率和脉宽测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        try:
            results = {}
            """
            self.scope.write(f'FILESYSTEM:READFILE "C:\\Test\\Frequency{self.time}.txt"')
            data=self.scope.read()
            #fid=open(self.folder_name+"\\Parameters\\Freq.txt", 'w')
            #fid.write(data)
            #fid.close()
            lines = data.strip().splitlines()
            results["Frequency"] = lines[-5].split(',')[4]
            results["Positive_Pulse_Width"] = lines[-3].split(',')[4]
            results["Negative_Pulse_Width"] = lines[-1].split(',')[4]
            """
            results["SCL Frequency"] = round_sig(adapt_frequency(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            results["SCL Positive_Pulse_Width"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))))
            results["SCL Negative_Pulse_Width"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))))
            return True, "频率和脉宽测量值获取成功", results
        except Exception as e:
            return False, f"获取频率和脉宽测量值失败: {str(e)}", {}
        
    def add_delay_measurement(self):
        """添加延迟测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        try:
            # 清除所有现有测量
            self.clear_measurements()
            
            # setup time (从CH2上升沿70%到CH1上升沿30%)
            self.scope.write("MEASUREMENT:ADDMEAS DELAY")
            self.scope.write("MEASUrement:GATing SCREEN")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE1 {self.channels['SDA']}")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE2 {self.channels['SCL']}")
            
            # 设置参考电平方法为百分比
            self.scope.write(f"MEASUREMENT:{self.channels['SDA']}:REFLevels:METHod PERCent")
            self.scope.write(f"MEASUREMENT:{self.channels['SCL']}:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:REFLevels:TYPE PerSource")
            #self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            
            # 设置CH2(SOURCE1)的参考电平为70%
            self.scope.write(f"MEASUrement:{self.channels['SCL']}:REFLevels:PERCent:RISEHigh 90")
            self.scope.write(f"MEASUrement:{self.channels['SCL']}:REFLevels:PERCent:RISEMid 30")
            self.scope.write(f"MEASUrement:{self.channels['SCL']}:REFLevels:PERCent:RISELow 10")
            # 设置CH1(SOURCE2)的参考电平为30%
            self.scope.write(f"MEASUrement:{self.channels['SDA']}:REFLevels:PERCent:RISEHigh 90")
            self.scope.write(f"MEASUrement:{self.channels['SDA']}:REFLevels:PERCent:RISEMid 70")
            self.scope.write(f"MEASUrement:{self.channels['SDA']}:REFLevels:PERCent:RISELow 10")
            # 设置测量的边沿和方向
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION1 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION2 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE1 RISE")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE2 RISE")
            
            
            #self.scope.write(f'SAVe:EVENTtable:MEASUrement "C:\\Test\\setup_time{self.time}.txt"')
            self.scope.query("*OPC?")
            self.scope.query("MEASUREMENT:MEAS1:VALUE?")
            return True, "setup time测量项添加成功"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def add_setup_measurement_manual(self, sda_mode, scl_mode, sda_edge, scl_edge, sda_edge_level, scl_edge_level):
        if not self.connected:
            return False, "未连接到示波器"
    
        try:
            # 清除现有测量
            self.clear_measurements()
            self.scope.write("MEASUREMENT:ADDMEAS DELAY")
            self.scope.write("MEASUrement:GATing SCREEN")
            self.scope.write("MEASUrement:REFLevels:TYPE PerSource")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE1 {self.channels['SDA']}")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE2 {self.channels['SCL']}")

            def configure_edge(channel, mode, edge, level, edge_label):
                """简化设置参考电平和边沿方向"""
                method = "PERCent" if mode == "Percentage" else "ABSolute"
                self.scope.write(f"MEASUrement:{channel}:REFLevels:METHod {method}")
                print(f"MEASUrement:{channel}:REFLevels:METHod {method}")
                edge_dir = "RISE" if edge == "Rise" else "FALL"
                self.scope.write(f"MEASUREMENT:MEAS1:DELAY:{edge_label} {edge_dir}")
                print(f"MEASUREMENT:MEAS1:DELAY:{edge_label} {edge_dir}")
                level_cmd = f"{method}:{edge_dir}Mid {level}"
                self.scope.write(f"MEASUrement:{channel}:REFLevels:{level_cmd}")
                print(f"MEASUrement:{channel}:REFLevels:{level_cmd}")

            # SDA = EDGE1
            configure_edge(self.channels['SDA'], sda_mode, sda_edge, sda_edge_level, "EDGE1")

            # SCL = EDGE2
            configure_edge(self.channels['SCL'], scl_mode, scl_edge, scl_edge_level, "EDGE2")

            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION1 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION2 FORWARDS")
            self.scope.query("*OPC?")
        
            return True, "setup time测量项添加成功"

        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def add_rise_fall_measurements(self):
        """添加电压测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        
        try:
            # 清除所有现有测量
            self.clear_measurements()
            self.scope.write("MEASUrement:REFLEVELS:TYPE GLOBAL")
            self.scope.write(f"MEASUREMENT:{self.channels['SDA']}:REFLevels:METHod PERCent")
            self.scope.write(f"MEASUREMENT:{self.channels['SCL']}:REFLevels:METHod PERCent")
            # 添加电压参数
            self.scope.write("MEASUREMENT:ADDMEAS RISETIME")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            
            self.scope.write("MEASUREMENT:ADDMEAS FALLTIME")
            self.scope.write(f"MEASUREMENT:MEAS2:SOURCE {self.channels['SCL']}")
            
            self.scope.write("MEASUREMENT:ADDMEAS RISETIME")
            self.scope.write(f"MEASUREMENT:MEAS3:SOURCE {self.channels['SDA']}")
            
            self.scope.write("MEASUREMENT:ADDMEAS FALLTIME")
            self.scope.write(f"MEASUREMENT:MEAS4:SOURCE {self.channels['SDA']}")

            self.scope.write("MEASUrement:REFLevels:PERCent:RISEHigh 70")
            self.scope.write("MEASUrement:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUrement:REFLevels:PERCent:RISELow 30")

            self.scope.write("MEASUrement:REFLevels:PERCent:FALLHigh 70")
            self.scope.write("MEASUrement:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUrement:REFLevels:PERCent:FALLLow 30")
            self.scope.query("*OPC?")

            return True, "Rise/Fall测量项添加成功"
        except Exception as e:
            print(str(e))
            return False, f"添加Rise/Fall测量项失败: {str(e)}"
        
    def get_rise_measurements(self):
        """添加电压测量参数"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            results["SCL_Rise"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            results["SDA_Rise"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))))
            print(results)
            return True, "Rise测量值获取成功", results
        except Exception as e:
            return False, f"获取Rise测量值失败: {str(e)}", {}
        
    def add_rise_measurements(self, high, low, mode):
        """添加电压测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        method="PERCent" if mode=="Percentage" else "ABSolute"
        print(method)
        try:
            # 清除所有现有测量
            self.clear_measurements()
            self.scope.write("MEASUrement:REFLEVELS:TYPE GLOBAL")
            self.scope.write(f"MEASUrement:REFLevels:METHod {method}")
        
            self.scope.write("MEASUREMENT:ADDMEAS RISETIME")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            
            self.scope.write("MEASUREMENT:ADDMEAS RISETIME")
            self.scope.write(f"MEASUREMENT:MEAS2:SOURCE {self.channels['SDA']}")
            
            self.scope.write(f"MEASUrement:REFLevels:{method}:RISEHigh {high}")
            self.scope.write(f"MEASUrement:REFLevels:{method}:RISELow {low}")
            self.scope.query("*OPC?")

            return True, "Rise测量项添加成功"
        except Exception as e:
            print(str(e))
            return False, f"添加Rise测量项失败: {str(e)}"
        
    def add_fall_measurements(self, high, low, mode):
        """添加电压测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        method="PERCent" if mode=="Percentage" else "ABSolute"
        print(method)
        try:
            # 清除所有现有测量
            self.clear_measurements()
            self.scope.write("MEASUrement:REFLEVELS:TYPE GLOBAL")
            self.scope.write(f"MEASUrement:REFLevels:METHod {method}")
        
            self.scope.write("MEASUREMENT:ADDMEAS FALLTIME")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            
            self.scope.write("MEASUREMENT:ADDMEAS FALLTIME")
            self.scope.write(f"MEASUREMENT:MEAS2:SOURCE {self.channels['SDA']}")
            
            self.scope.write(f"MEASUrement:REFLevels:{method}:FALLHigh {high}")
            self.scope.write(f"MEASUrement:REFLevels:{method}:FALLLow {low}")
            self.scope.query("*OPC?")

            return True, "FALL测量项添加成功"
        except Exception as e:
            print(str(e))
            return False, f"添加FALL测量项失败: {str(e)}"
        
    def get_fall_measurements(self):
        """添加电压测量参数"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            results["SCL_Fall"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            results["SDA_Fall"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))))
            return True, "Rise测量值获取成功", results
        except Exception as e:
            return False, f"获取Rise测量值失败: {str(e)}", {}
        
    def get_rise_fall_measurements(self):
        """获取电压测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            results["SCL_Rise"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            results["SCL_Fall"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))))
            results["SDA_Rise"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))))
            results["SDA_Fall"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS4:VALUE?"))))
            
            return True, "Rise/Fall测量值获取成功", results
        except Exception as e:
            return False, f"获取Rise/Fall测量值失败: {str(e)}", {}
    
    def add_hold_time_measurements(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            self.scope.write("MEASUREMENT:ADDMEAS DELAY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE1 {self.channels['SCL']}")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE2 {self.channels['SDA']}")
            self.scope.write(f"MEASUREMENT:{self.channels['SDA']}:REFLevels:METHod PERCent")
            self.scope.write(f"MEASUREMENT:{self.channels['SCL']}:REFLevels:METHod PERCent")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUrement:GATing SCREEN")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:REFLevels:TYPE PerSource")
            
            # 设置CH2(SOURCE1)的参考电平为70%
            self.scope.write(f"MEASUrement:{self.channels['SCL']}:REFLevels:PERCent:FALLMid 30")
            # 设置CH1(SOURCE2)的参考电平为30%
            self.scope.write(f"MEASUrement:{self.channels['SDA']}:REFLevels:PERCent:FALLMid 70")
            # 设置测量的边沿和方向
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION1 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION2 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE1 FALL")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE2 FALL")
            self.scope.query("*OPC?")

            return True, "Hold Time测量项添加成功"
        except Exception as e:
            print(str(e))
            return False, f"添加Hold Time测量项失败: {str(e)}"
        
    def add_hold_measurement_manual(self, sda_mode, scl_mode, sda_edge, scl_edge, sda_edge_level, scl_edge_level):
        if not self.connected:
            return False, "未连接到示波器"
    
        try:
            # 清除现有测量
            self.clear_measurements()
            self.scope.write("MEASUREMENT:ADDMEAS DELAY")
            self.scope.write("MEASUrement:GATing SCREEN")
            self.scope.write("MEASUrement:REFLevels:TYPE PerSource")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE1 {self.channels['SCL']}")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE2 {self.channels['SDA']}")

            def configure_edge(channel, mode, edge, level, edge_label):
                method = "PERCent" if mode == "Percentage" else "ABSolute"
                self.scope.write(f"MEASUrement:{channel}:REFLevels:METHod {method}")
                edge_dir = "RISE" if edge == "Rise" else "FALL"
                self.scope.write(f"MEASUREMENT:MEAS1:DELAY:{edge_label} {edge_dir}")
                level_cmd = f"{method}:{edge_dir}Mid {level}"
                self.scope.write(f"MEASUrement:{channel}:REFLevels:{level_cmd}")

            # SDA = EDGE1
            configure_edge(self.channels['SDA'], sda_mode, sda_edge, sda_edge_level, "EDGE2")

            # SCL = EDGE2
            configure_edge(self.channels['SCL'], scl_mode, scl_edge, scl_edge_level, "EDGE1")

            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION1 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION2 FORWARDS")
            self.scope.query("*OPC?")
        
            return True, "setup time测量项添加成功"

        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def get_delay_measurement(self):
        """获取Hold Time测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            results["SDA Setup Time"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            return True, "Setup Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Setup Time测量值失败: {str(e)}", {}
        
    def get_hold_measurement(self):
        """获取Setup Time测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            results["SDA Hold Time"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            return True, "Hold Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Hold Time测量值失败: {str(e)}", {}
        

    def add_stop_time_measurement(self):
        """添加延迟测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        try:
            # 清除所有现有测量
            self.clear_measurements()
            
            # setup time (从CH2上升沿70%到CH1上升沿30%)
            self.scope.write("MEASUREMENT:ADDMEAS DELAY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE1 {self.channels['SCL']}")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE2 {self.channels['SDA']}")
            self.scope.write("MEASUrement:GATing SCREEN")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            #self.scope.write("MEASUrement:REFLevels:TYPE PerSource")
            #self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            
            # 设置CH2(SOURCE1)的参考电平为70%
            self.scope.write(f"MEASUrement:{self.channels['SCL']}:REFLevels:PERCent:RISEHigh 90")
            self.scope.write(f"MEASUrement:{self.channels['SCL']}:REFLevels:PERCent:RISEMid 70")
            self.scope.write(f"MEASUrement:{self.channels['SCL']}:REFLevels:PERCent:RISELow 10")
            # 设置CH1(SOURCE2)的参考电平为30%
            self.scope.write(f"MEASUrement:{self.channels['SDA']}:REFLevels:PERCent:RISEHigh 90")
            self.scope.write(f"MEASUrement:{self.channels['SDA']}:REFLevels:PERCent:RISEMid 30")
            self.scope.write(f"MEASUrement:{self.channels['SDA']}:REFLevels:PERCent:RISELow 10")
            # 设置测量的边沿和方向
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION1 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION2 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE1 RISE")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE2 RISE")
            #self.scope.write(f'SAVe:EVENTtable:MEASUrement "C:\\Test\\setup_time{self.time}.txt"')
            self.scope.query("*OPC?")
            return True, "stop time测量项添加成功"
        except Exception as e:
            return False, f"添加stop time测量项失败: {str(e)}"
        
    def get_stop_time_measurement(self):
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            results["Stop Setup Time"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            return True, "Stop Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Stop Time测量值失败: {str(e)}", {}
        
    def save_screenshot(self, filename, foldername):
        """保存示波器截图到C:\ST\hsd\i2c目录"""
        if not self.connected:
            return False, "未连接到示波器"
        
        try:
            #no change in filepath but change in filename
            self.scope.write(f'SAVe:IMAGe "C:\\Test\\{filename+".png"}"')
            self.scope.query("*OPC?")
            self.scope.write(f'FILESYSTEM:READFILE "C:\\Test\\{filename+".png"}"')
            raw_data=self.scope.read_raw()
            #check if folder name is in path
            #print(self.image_path)
            full_path = os.path.join(self.image_path, foldername)
            if os.path.isdir(full_path):
                fid=open(self.image_path+f"\\{foldername}\\{filename}.png", 'wb')
                fid.write(raw_data)
                fid.close()
            else:
                os.chdir(self.image_path)
                os.mkdir(foldername)
                fid=open(self.image_path+f"\\{foldername}\\{filename}.png", 'wb')
                fid.write(raw_data)
                fid.close()
            if "Rising Edge" not in filename and "Falling Edge" not in filename:
                self.scope.write(f'SAVe:EVENTtable:MEASUrement "C:\\Test\\{filename+".csv"}"')
                self.scope.query("*OPC?")
                self.scope.write(f'FILESYSTEM:READFILE "C:\\Test\\{filename+".csv"}"')
                data=self.scope.read()
                fid=open(self.image_path+f"\\{foldername}\\{filename}.csv", 'w')
                fid.write(data)
                fid.close()
                # Step 1: Read all lines
                with open(self.image_path+f"\\{foldername}\\{filename}.csv", 'r') as file:
                    lines = file.readlines()
                lines=lines[8:]
                new_lines=[]
                for i in lines:
                    if i!='\n':
                        new_lines.append(i)

                # Step 3: Overwrite the file
                with open(self.image_path+f"\\{foldername}\\{filename}.csv", 'w') as file:
                    file.writelines(new_lines)

                df=pd.read_csv(self.image_path+f"\\{foldername}\\{filename}.csv")
                

            
            

            return True, f"截图已保存至 {self.image_path}\\{foldername}\\{filename}"+".png"
        except Exception as e:
            error_msg = f"保存截图失败: {str(e)}"
            #self.log_message(error_msg)
            return False, error_msg
        
        
    def save_data_to_excel(self, measurements, filename, foldername, images):
        """保存示波器截图到C:\ST\hsd\i2c目录"""
        if not self.connected:
            return False, "未连接到示波器"
        df = pd.DataFrame(list(measurements.items()), columns=['Test Item', 'Measured Data'])
        try:
            full_path = os.path.join(self.excel_path, "excels")
            if os.path.isdir(full_path):
                #test if the excel folder already exists
                with pd.ExcelWriter(self.excel_path+f"\\excels\\{foldername}.xlsx", engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False)

                    workbook  = writer.book
                    worksheet = writer.sheets['Sheet1']

                    # Autofit columns based on content
                    for i, col in enumerate(df.columns):
                        #column_len = max(df[col].astype(str).map(len).max(), len(col))
                        #print(column_len+2)
                        worksheet.set_column(i, i, 24)  # Add a little extra space
                    """
                    image_path = images['Voltage']
                    img = Image.open(image_path)
                    width, height = img.size
                    worksheet.set_column('C:C', width * 0.33 / 8.815)
                    worksheet.set_row(1, height * 0.33 / 1.665)
                    worksheet.insert_image('C2', image_path, {'x_scale': 0.33, 'y_scale': 0.33})
                    """
                    # Define scaling factors
                    x_scale = 0.33
                    y_scale = 0.33
                    Col='A'
                    index=0
                    row=len(df)+3
                    text_col='B'
                    for key in list(images.keys()):
                        if index%2==0:
                            Col='A'
                            text_col='B'
                        else:
                            Col='F'
                            text_col='I'
                        row=16*math.floor(index//2)+len(df)+5
                        image_path = images[key]
                        img = Image.open(image_path)
                        width, height = img.size
                    # Insert image at 'C2' with scaling
                        worksheet.insert_image(f'{Col}{str(row)}', image_path, {'x_scale': x_scale, 'y_scale': y_scale})
                        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                        worksheet.write(f'{text_col}{str(row-1)}', key, cell_format)
                        index+=1
            else:
                os.chdir(self.excel_path)
                os.mkdir("excels")
                #df.to_excel(self.excel_path+f"\\{foldername}\\{filename}.xlsx", index=False)
                full_path = os.path.join(self.excel_path, "excels")
                with pd.ExcelWriter(self.excel_path+f"\\excels\\{foldername}.xlsx", engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False)

                    workbook  = writer.book
                    worksheet = writer.sheets['Sheet1']

                    # Autofit columns based on content
                    for i, col in enumerate(df.columns):
                        #column_len = max(df[col].astype(str).map(len).max(), len(col))
                        worksheet.set_column(i, i, 24)  # Add a little extra space
                    x_scale = 0.33
                    y_scale = 0.33
                    Col='A'
                    index=0
                    row=len(df)+3
                    text_col='B'
                    for key in list(images.keys()):
                        if index%2==0:
                            Col='A'
                            text_col='B'
                        else:
                            Col='F'
                            text_col='I'
                        row=16*math.floor(index//2)+len(df)+5
                        image_path = images[key]
                        img = Image.open(image_path)
                    # Insert image at 'C2' with scaling
                        worksheet.insert_image(f'{Col}{str(row)}', image_path, {'x_scale': x_scale, 'y_scale': y_scale})
                        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                        worksheet.write(f'{text_col}{str(row-1)}', key, cell_format)
                        index+=1
                 
            return True, "Excel saved to "+f"{self.excel_path}"+f"\\{foldername}.xlsx"
        except Exception as e:
            error_msg = f"保存Excel失败: {str(e)}"
            #self.log_message(error_msg)
            return False, error_msg
    
    def add_delay2_measurement(self):
        """添加延迟测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        try:
            # 清除所有现有测量
            self.clear_measurements()
            
            # setup time (从CH2上升沿70%到CH1上升沿30%)
            self.scope.write("MEASUREMENT:ADDMEAS DELAY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE1 {self.channels['SDA']}")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE2 {self.channels['SCL']}")
            self.scope.write("MEASUrement:GATing SCREEN")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:REFLevels:TYPE PerSource")
            #self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            
            # 设置CH2(SOURCE1)的参考电平为70%
            self.scope.write(f"MEASUrement:{self.channels['SCL']}:REFLevels:PERCent:FALLHigh 90")
            self.scope.write(f"MEASUrement:{self.channels['SCL']}:REFLevels:PERCent:FALLMid 70")
            self.scope.write(f"MEASUrement:{self.channels['SCL']}:REFLevels:PERCent:FALLLow 10")
            # 设置CH1(SOURCE2)的参考电平为30%
            self.scope.write(f"MEASUrement:{self.channels['SDA']}:REFLevels:PERCent:FALLHigh 90")
            self.scope.write(f"MEASUrement:{self.channels['SDA']}:REFLevels:PERCent:FALLMid 30")
            self.scope.write(f"MEASUrement:{self.channels['SDA']}:REFLevels:PERCent:FALLLow 10")
            # 设置测量的边沿和方向
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION1 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION2 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE1 FALL")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE2 FALL")
            
            
            #self.scope.write(f'SAVe:EVENTtable:MEASUrement "C:\\Test\\setup_time{self.time}.txt"')
            self.scope.query("*OPC?")
            return True, "Start Hold Time测量项添加成功"
        except Exception as e:
            return False, f"添加Start Hold Time测量项失败: {str(e)}"
        
    def get_start_hold_time(self):
        """获取Setup Time测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            results["Start Hold Time"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            return True, "Start Hold Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Start Hold Time测量值失败: {str(e)}", {}
        
    def zoom_on_rising_edge(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            # Optionally, use search to locate the rising edge # Clear existing searches
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            # Wait for acquisition and search   
            self.scope.query("*OPC?")  # Ensure acquisition completes
            for i in range(1):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")
            
            self.scope.write("MEASUrement:REFLEVELS:TYPE GLOBAL")
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGGER:A:EDGE:SOURCE {self.channels["SCL"]}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            for i in range(1):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")
            # 添加电压参数
            self.scope.write("MEASUREMENT:ADDMEAS RISETIME")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            self.scope.write(f"MEASUrement:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GATing SCREEN")
            self.scope.write("MEASUrement:REFLevels:PERCent:RISEHigh 70")
            self.scope.write("MEASUrement:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUrement:REFLevels:PERCent:RISELow 30")
            self.scope.query("*OPC?")
            scale=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            print(scale)
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {scale/2}")
            #self.scope.write("SEARCH:DELETEALL")

            return True, "success"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def zoom_on_falling_edge(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:

            self.clear_measurements()
            # Optionally, use search to locate the rising edge
            self.scope.write("SEARCH:DELETEALL")  # Clear existing searches
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGGER:A:EDGE:SOURCE {self.channels["SCL"]}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
           
            
            # Wait for acquisition and search   
            self.scope.query("*OPC?")  # Ensure acquisition completes
            for i in range(4):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")
            #self.scope.write("MEASUrement:REFLEVELS:TYPE GLOBAL")
            # 添加电压参数
            
            self.scope.write("MEASUREMENT:ADDMEAS FALLTIME")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            self.scope.write(f"MEASUrement:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GATing NONE")
            self.scope.write("MEASUrement:REFLevels:PERCent:FALLHigh 70")
            self.scope.write("MEASUrement:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUrement:REFLevels:PERCent:FALLLow 30")
            self.scope.query("*OPC?")
            scale=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {scale/2}")
            self.scope.write("SEARCH:DELETEALL")
            

            return True, "success"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"

    def zoom_middle(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            self.scope.write("SEARCH:DELETEALL")
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/1.5}")
            self.clear_measurements()
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGGER:A:EDGE:SOURCE {self.channels["SDA"]}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
            self.scope.query("*OPC?")

            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGGER:A:EDGE:SOURCE {self.channels["SCL"]}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            for i in range(5):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength*8/10}")
           
            self.clear_measurements()
            self.scope.write("SEARCH:DELETEALL")
            
            return True, "success"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def zoom_middle_read(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            #self.clear_measurements()
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 100")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            #n=self.scope.query("SEARCH:SEARCH1:TOTAL?")
            #self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/1.35}")

            for i in range(6):
                self.scope.write("SEARCH:SEARCH1:NAVigate PREVious")
                self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/1.5}")
            self.clear_measurements()
            self.scope.write("SEARCH:DELETEALL")
            return True, "success"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def zoom_voltage(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            self.clear_measurements()
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.query("*OPC?")
            # Optionally, use search to locate the rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            #n=self.scope.query("SEARCH:SEARCH1:TOTAL?")
            #self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/1.35}")
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            for i in range(2):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/1.5}")
        
            self.clear_measurements()
            self.scope.write("SEARCH:DELETEALL")
            
            return True, "success"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def zoom_frequency(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            #self.clear_measurements()
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.query("*OPC?")
            # Optionally, use search to locate the rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            #n=self.scope.query("SEARCH:SEARCH1:TOTAL?")
            #self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/1.35}")
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            for i in range(2):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/2.5}")
        
            self.clear_measurements()
            self.scope.write("SEARCH:DELETEALL")
            
            return True, "success"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
    
    def zoom_rise(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            self.clear_measurements()
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.query("*OPC?")
            # Optionally, use search to locate the rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            #n=self.scope.query("SEARCH:SEARCH1:TOTAL?")
            #self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/1.35}")
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            for i in range(2):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/10}")
        
            self.clear_measurements()
            self.scope.write("SEARCH:DELETEALL")
            
            return True, "success"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def zoom_fall(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            self.clear_measurements()
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.query("*OPC?")
            # Optionally, use search to locate the rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            #n=self.scope.query("SEARCH:SEARCH1:TOTAL?")
            #self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/1.35}")
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            for i in range(2):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/10}")
        
            #self.clear_measurements()
            #self.scope.write("SEARCH:DELETEALL")
            
            return True, "success"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def zoom_hold_time(self, num_zoom=2):
        #find 2nd ch2 falling edge
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.query("*OPC?")
            # Optionally, use search to locate the rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            for i in range(num_zoom):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")

            self.clear_measurements()
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/10}")
            self.clear_measurements()
            
        except Exception as e:
            return False, e
        return True, ":)"
    
    def zoom_hold_time_read(self):
        #find 2nd to last ch2 falling edge
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/10}")
            self.clear_measurements()
            # Optionally, use search to locate the rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            n=self.scope.query("SEARCH:SEARCH1:TOTAL?")
            #print(n)
            
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 2000")
            for i in range(2):
                self.scope.write("SEARCH:SEARCH1:NAVigate PREVious")
                self.scope.query("*OPC?")
            self.clear_measurements()
            
        except Exception as e:
            return False, e
        return True, ":)"
    
    def zoom_setup_time(self, num_zoom=1):
        #find 2nd ch2 falling edge
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.query("*OPC?")
            # find 1st ch2 rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            for i in range(num_zoom):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")

            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SCL']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")

            self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
            self.scope.query("*OPC?")
            

            self.clear_measurements()
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/10}")
            self.clear_measurements()
            
        except Exception as e:
            return False, e
        return True, ":)"
    
    def zoom_setup_time_read(self):
        #find 3rd to last ch2 rising edge
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.query("*OPC?")
            # find 1st ch2 rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 100")
            for i in range(3):
                self.scope.write("SEARCH:SEARCH1:NAVigate PREVious")
                self.scope.query("*OPC?")
            
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SCL']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")

            self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
            self.scope.query("*OPC?")

            self.clear_measurements()
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/10}")
            self.clear_measurements()
            
        except Exception as e:
            return False, e
        return True, ":)"
    
    def zoom_start(self):
        #find 2nd ch2 falling edge
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.query("*OPC?")
            # find 1st ch2 rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            for i in range(1):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")
            
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/5}")
            self.clear_measurements()

            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SCL']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")

            for i in range(1):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")
            self.scope.write("SEARCH:DELETEALL")
            
        except Exception as e:
            return False, e
        return True, ":)"
    
    def zoom_stop(self):
        #find 2nd ch2 falling edge
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.query("*OPC?")
            # find 1st ch2 rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce {self.channels['SDA']}")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 2000")
            for i in range(1):
                self.scope.write("SEARCH:SEARCH1:NAVigate PREVious")
                self.scope.query("*OPC?")

            self.clear_measurements()
            # 频率 (从30%到30%)
            #self.scope.write("MEASUrement:MEAS?:GLOBalref 0")
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            # 设置上升沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 50")
            # 设置下降沿参考电平
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLHigh 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:FALLLow 50")
            self.scope.query("*OPC?")
            frequency=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            wavelength=1/frequency
            self.scope.query("*OPC?")
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {wavelength/5}")
            self.clear_measurements()
            return True, ":)"
        except Exception as e:
            return False, e


class MSO64ControllerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MSO64示波器控制器")
        self.root.geometry("1500x1000")
        self.controller = MSO64Controller()
        self.create_widgets()
        self.measured_voltage=False
        self.measured_frequency=False
        self.measured_setup=False
        self.res={}
        self.images={}#schema: name, file path
        self.measured_rise_fall=False
        self.measured_hold=False
        self.measured_start_hold=False
        self.measured_stop_setup=False
        self.measured_voltage_read=False
        self.measured_rise_fall_read=False
        self.measured_setup_read=False
        self.measured_hold_read=False
        self.res_read={}
        self.images_read={}
        self.controller.channels['SCL']=self.clock_var.get()
        self.controller.channels['SDA']=self.data_var.get()
        self.set_max_voltage=False
        
    def create_widgets(self):
        connection_frame = ttk.LabelFrame(self.root, text="连接设置")
        connection_frame.pack(fill="x", padx=10, pady=5)

        # IP地址输入
        ttk.Label(connection_frame, text="示波器IP地址:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.ip_entry = ttk.Entry(connection_frame, width=20)
        self.ip_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.ip_entry.insert(0, "169.254.123.124")  # 默认IP

        # 连接按钮
        self.connect_button = ttk.Button(connection_frame, text="连接", command=self.connect)
        self.connect_button.grid(row=0, column=2, padx=5, pady=5)

        # Keep these as instance attributes so they are accessible everywhere
        self.clock_options = ["REF1", "CH2", "CH3", "CH4"]
        self.data_options = ["REF2", "CH3", "CH4"]
        self.voltage_options=["1.8V", "3.3V"]
        
        self.clock_var = tk.StringVar(value=self.clock_options[0])
        self.data_var = tk.StringVar(value=self.data_options[0])
        self.voltage_var=tk.StringVar(value=self.voltage_options[0])
       
        # Set initial values for clock and data
        self.clock = self.clock_var.get()
        self.data = self.data_var.get()
        self.voltage=self.voltage_var.get()

        # Register trace callbacks to update clock/data values
        self.clock_var.trace_add("write", self.on_channel_change)
        self.data_var.trace_add("write", self.on_channel_change)
        self.voltage_var.trace_add("write", self.on_voltage_change)

        # Create dropdowns
        ttk.Label(connection_frame, text="CLK: ").grid(row=0, column=5, padx=5, pady=5, sticky="w")
        self.clock_dropdown = tk.OptionMenu(connection_frame, self.clock_var, *self.clock_options)
        self.clock_dropdown.grid(row=0, column=6, padx=10, pady=5)

        ttk.Label(connection_frame, text="DATA: ").grid(row=0, column=7, padx=5, pady=5, sticky="w")
        self.data_dropdown = tk.OptionMenu(connection_frame, self.data_var, *self.data_options)
        self.data_dropdown.grid(row=0, column=8, padx=10, pady=5)

        ttk.Label(connection_frame, text="Max Voltage: ").grid(row=0, column=9, padx=5, pady=5, sticky="w")
        self.voltage_dropdown = tk.OptionMenu(connection_frame, self.voltage_var, *self.voltage_options)
        self.voltage_dropdown.grid(row=0, column=10, padx=10, pady=5)

        # 状态标签
        ttk.Label(connection_frame, text="状态:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.status_label = ttk.Label(connection_frame, text="未连接")
        self.status_label.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="w")

        # 测试框架
        test_frame = ttk.LabelFrame(self.root, text="测试操作")
        test_frame.pack(fill="x", padx=10, pady=5)
        
        # 图片保存路径
        ttk.Label(test_frame, text="图片保存路径:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.image_path_entry = ttk.Entry(test_frame, width=35)
        self.image_path_entry.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        
        self.image_browse_button = ttk.Button(test_frame, text="浏览...", command=self.browse_image_path)
        self.image_browse_button.grid(row=0, column=3, padx=5, pady=5)

        # 保存到示波器选项
        """
        self.save_to_scope_image = tk.BooleanVar()
        self.save_to_scope_image_check = ttk.Checkbutton(test_frame, text="保存到示波器", 
                                                          variable=self.save_to_scope_image,
                                                          command=self.toggle_image_path)
        self.save_to_scope_image_check.grid(row=0, column=4, padx=5, pady=5)
        """
        # Excel保存路径
        ttk.Label(test_frame, text="Excel保存路径:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.excel_path_entry = ttk.Entry(test_frame, width=35)
        self.excel_path_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="w")

        
        self.excel_browse_button = ttk.Button(test_frame, text="浏览...", command=self.browse_excel_path)
        self.excel_browse_button.grid(row=1, column=3, padx=5, pady=5)
        
        # 保存到示波器选项
        """
        self.save_to_scope_excel = tk.BooleanVar()
        self.save_to_scope_excel_check = ttk.Checkbutton(test_frame, text="保存到示波器", 
                                                         variable=self.save_to_scope_excel,
                                                         command=self.toggle_excel_path)
        self.save_to_scope_excel_check.grid(row=1, column=4, padx=5, pady=5)
        """
        
        # 文件夹名称
        ttk.Label(test_frame, text="文件夹名称:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.folder_entry = ttk.Entry(test_frame, width=35)
        self.folder_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky="w")


        # 测量按钮
        i2c_frame = ttk.LabelFrame(self.root, text="I2C Tests")
        i2c_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.run_all_tests_button = ttk.Button(i2c_frame, text="Write Cycle", command=self.run_all_tests)
        self.run_all_tests_button.grid(row=0, column=0, padx=5, pady=5)

        ttk.Label(i2c_frame, text="Suffix Name:").grid(row=0, column=7, padx=5, pady=5, sticky="w")
        self.suffix_entry = ttk.Entry(i2c_frame, width=15)
        self.suffix_entry.grid(row=0, column=8, padx=5, pady=5, sticky="w")

        self.voltage_button = ttk.Button(i2c_frame, text="添加电压测量", command=self.write_voltage_group)
        self.voltage_button.grid(row=1, column=0, padx=5, pady=5)

        self.frequency_button = ttk.Button(i2c_frame, text="添加频率测量", command=self.write_frequency_group)
        self.frequency_button.grid(row=1, column=1, padx=5, pady=5)

        self.rise_fall_button = ttk.Button(i2c_frame, text="添加Rise/Fall测量", command=self.write_rise_fall_group)
        self.rise_fall_button.grid(row=1, column=2, padx=5, pady=5)

        self.delay_button = ttk.Button(i2c_frame, text="添加Setup Time测量", command=self.write_setup_group)
        self.delay_button.grid(row=1, column=3, padx=5, pady=5)

        self.hold_time_button = ttk.Button(i2c_frame, text="添加Hold Time测量", command=self.write_hold_group)
        self.hold_time_button.grid(row=1, column=4, padx=5, pady=5)

        self.start_hold_time_button = ttk.Button(i2c_frame, text="添加Start Hold Time测量", command=self.write_start_hold_group)
        self.start_hold_time_button.grid(row=1, column=5, padx=5, pady=5)

        self.stop_setup_time_button = ttk.Button(i2c_frame, text="添加Stop Setup Time测量", command=self.write_stop_setup_group)
        self.stop_setup_time_button.grid(row=1, column=6, padx=5, pady=5)

        self.run_read_tests_button = ttk.Button(i2c_frame, text="Read Cycle", command=self.run_read_tests)
        self.run_read_tests_button.grid(row=2, column=0, padx=5, pady=5)

        self.voltage_button_read = ttk.Button(i2c_frame, text="添加电压测量", command=self.read_voltage_group)
        self.voltage_button_read.grid(row=3, column=0, padx=5, pady=5)

        self.rise_fall_button_read = ttk.Button(i2c_frame, text="添加Rise Fall Time测量", command=self.read_rise_fall_group)
        self.rise_fall_button_read.grid(row=3, column=1, padx=5, pady=5)

        self.setup_button_read = ttk.Button(i2c_frame, text="添加Setup Time测量", command=self.read_setup_group)
        self.setup_button_read.grid(row=3, column=2, padx=5, pady=5)

        self.hold_button_read = ttk.Button(i2c_frame, text="添加Hold Time测量", command=self.read_hold_group)
        self.hold_button_read.grid(row=3, column=3, padx=5, pady=5)


        # 创建测量显示框架
        measurement_frame = ttk.LabelFrame(self.root, text="测量结果")
        measurement_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # 设置选项卡
        self.tab_control = ttk.Notebook(measurement_frame)
        self.tab_control.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 电压测量选项卡
        self.voltage_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.voltage_tab, text="电压参数")
        
        # 频率测量选项卡
        self.frequency_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.frequency_tab, text="频率参数")

        self.rise_fall_tab=ttk.Frame(self.tab_control)
        self.tab_control.add(self.rise_fall_tab, text="Rise/Fall参数")

        # Setup Time测量选项卡
        self.delay_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.delay_tab, text="Setup Time参数")

        self.hold_time_tab=ttk.Frame(self.tab_control)
        self.tab_control.add(self.hold_time_tab, text="Hold Time参数")
        """
        self.start_hold_time_tab=ttk.Frame(self.tab_control)
        self.tab_control.add(self.start_hold_time_tab, text="Start Hold Time参数")

        self.stop_setup_time_tab=ttk.Frame(self.tab_control)
        self.tab_control.add(self.stop_setup_time_tab, text="Stop Setup Time参数")
        """
        # 运行测试按钮
    
        self.run_test_button = ttk.Button(test_frame, text="Save to Excel", command=self.save_excel)
        self.run_test_button.grid(row=3, column=0, padx=5, pady=5)

        # 创建电压测量参数表
        self.create_voltage_tab()
        
        # 创建频率测量参数表
        self.create_frequency_tab()
        
        # 创建Setup Time测量参数表
        #self.create_start_hold_time_tab()
        self.create_delay_tab()

        self.create_rise_fall_tab()

        self.create_hold_time_tab()

        #self.create_start_hold_time_tab()

        #self.create_stop_setup_time_tab()

        # 创建状态面板
        status_frame = ttk.LabelFrame(self.root, text="保存状态")
        status_frame.pack(fill="x", padx=10, pady=5)
        
        # 创建状态标签
        """
        ttk.Label(status_frame, text="Voltage.png:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.voltage_status = ttk.Label(status_frame, text="未保存")
        self.voltage_status.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(status_frame, text="Freq.png:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.freq_status = ttk.Label(status_frame, text="未保存")
        self.freq_status.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        ttk.Label(status_frame, text="setup_time.png:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.setup_status = ttk.Label(status_frame, text="未保存")
        self.setup_status.grid(row=0, column=5, padx=5, pady=5, sticky="w")
        """

        #log
        log_frame = ttk.LabelFrame(self.root, text="日志")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.log_text = tk.Text(log_frame, height=8, width=70)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)

        self.update_button_states(False)

    def run_all_tests(self):
        if self.controller.excel_path=="":
            messagebox.showerror("错误", "File path has not been configured")
            return
        if self.controller.image_path is None:
            messagebox.showerror("错误", "Image path has not been configured!")
            return
        self.add_voltage_measurements(False, True)
        self.controller.scope.query("*OPC?")
        self.get_voltage_measurements(False)
        time.sleep(0.5)
        self.save_screenshot("Voltage", auto=True)
        self.controller.scope.query("*OPC?")
        self.add_frequency_measurements(False, True)
        self.controller.scope.query("*OPC?")
        self.get_frequency_measurements(False)
        time.sleep(0.5)
        self.save_screenshot("Frequency", auto=True)
        self.controller.scope.query("*OPC?")
        self.add_rise_fall_measurements(False, True)
        self.controller.scope.query("*OPC?")
        self.get_rise_fall_measurements(False)
        self.controller.scope.query("*OPC?")
        time.sleep(0.5)
        self.save_screenshot("Rise Fall", auto=True)
        self.controller.scope.query("*OPC?")
        self.add_delay_measurement(False, True)
        self.controller.scope.query("*OPC?")
        self.get_delay_measurement(False)
        self.controller.scope.query("*OPC?")
        time.sleep(0.5)
        self.save_screenshot("Setup Time", auto=True)
        self.controller.scope.query("*OPC?")
        self.add_hold_time_measurement(False, True)
        self.controller.scope.query("*OPC?")
        self.get_hold_time_measurement(False)
        self.controller.scope.query("*OPC?")
        time.sleep(0.5)
        self.save_screenshot("Hold Time", auto=True)
        self.controller.scope.query("*OPC?")
        self.add_start_hold_time_measurement(False, True)
        self.controller.scope.query("*OPC?")
        self.get_start_hold_time_measurement(False)
        self.controller.scope.query("*OPC?")
        time.sleep(0.5)
        self.save_screenshot("Start Hold Time", auto=True)
        self.controller.scope.query("*OPC?")
        self.add_stop_setup_time_measurement(False, True)
        self.controller.scope.query("*OPC?")
        self.get_stop_setup_time_measurement(False)
        self.controller.scope.query("*OPC?")
        time.sleep(0.5)
        self.save_screenshot("Stop Setup Time", auto=True)
        self.controller.scope.query("*OPC?")
        self.controller.zoom_on_rising_edge()
        self.controller.scope.query("*OPC?")
        time.sleep(0.5)
        self.save_screenshot("Rising Edge", auto=True)
        self.controller.scope.query("*OPC?")
        self.controller.zoom_on_falling_edge()
        self.controller.scope.query("*OPC?")
        time.sleep(0.5)
        self.save_screenshot("Falling Edge", end=True, auto=True)
        self.controller.scope.query("*OPC?")
        self.save_excel()
        self.folder_entry.delete(0, tk.END)

    
    def connect(self):
        ip_address = self.ip_entry.get().strip()
        if not ip_address:
            messagebox.showerror("错误", "请输入示波器IP地址")
            return
        self.log_message(f"正在连接到 {ip_address}...")

        # 如果已连接，先断开
        if self.controller.connected:
            self.controller.disconnect()
            self.status_label.config(text="未连接")
            self.connect_button.config(text="连接")
            self.update_button_states(False)
            self.log_message("已断开连接")
            return
        
        # 连接到示波器
        success, message = self.controller.connect(ip_address)
        if success:
            self.status_label.config(text="已连接")
            self.connect_button.config(text="断开")
            self.update_button_states(True)
            self.log_message(f"连接成功: {message}")
        else:
            messagebox.showerror("连接错误", message)
            self.log_message(f"连接失败: {message}")

    def log_message(self, message):
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)

    def browse_image_path(self):
        path = filedialog.askdirectory()
        newPath=path.replace('/', '\\')
        if path:
            self.image_path_entry.delete(0, tk.END)
            self.image_path_entry.insert(0, newPath)
            self.controller.image_path = newPath
            self.log_message(f"已选择图片保存路径: {newPath}")

    def add_voltage_measurements(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        self.update_voltage()
        self.log_message("正在添加电压测量项...")
        if move:
            self.controller.zoom_middle()
        success, message = self.controller.add_voltage_measurements()
        self.log_message(message)
        
        if success:
            self.measured_voltage=True
            if display:
                messagebox.showinfo("成功", "电压测量项添加成功\n")
                # 切换到电压选项卡
            self.tab_control.select(self.voltage_tab)
        else:
            messagebox.showerror("错误", message)

    def add_voltage_measurements_by_channel(self, channel, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        
        self.log_message("正在添加电压测量项...")
        if move:
            self.controller.zoom_voltage()
        success, message, measurements = self.controller.voltage_measurement_by_channel(channel)
        self.log_message(message)
        
        if success:
            self.res.update(measurements)
            self.measured_voltage=True
            channel_name=""
            for keys in self.controller.channels.keys():
                if self.controller.channels[keys]==channel:
                    channel_name=keys
            if channel_name=="SCL":
                self.scl_max_value.config(text=f"{measurements['SCL_Maximum']} V")
                self.scl_min_value.config(text=f"{measurements['SCL_Minimum']} V")
                self.scl_top_value.config(text=f"{measurements['SCL_Top']} V")
                self.scl_base_value.config(text=f"{measurements['SCL_Base']} V")
            if channel_name=="SDA":
                self.sda_max_value.config(text=f"{measurements['SDA_Maximum']} V")
                self.sda_min_value.config(text=f"{measurements['SDA_Minimum']} V")
                self.sda_top_value.config(text=f"{measurements['SDA_Top']} V")
                self.sda_base_value.config(text=f"{measurements['SDA_Base']} V")
            if display:
                messagebox.showinfo("成功", "电压测量值获取成功")
            self.save_screenshot(f"{channel_name} Voltage")
        else:
            messagebox.showerror("错误", message)


    def add_rise_fall_measurements(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Rise/Fall测量项...")
        if move:
            self.controller.zoom_middle()
        success, message = self.controller.add_rise_fall_measurements()
        self.log_message(message)
        
        if success:
            self.measured_rise_fall=True
            if display:
                messagebox.showinfo("成功", "Rise/Fall测量项添加成功\n")
            # 切换到电压选项卡
            self.tab_control.select(self.rise_fall_tab)
        else:
            messagebox.showerror("错误", message)

    def add_rise_fall_measurements_read(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Rise/Fall测量项...")
        if move:
            self.controller.zoom_middle_search()
        success, message = self.controller.add_rise_fall_measurements()
        self.log_message(message)
        
        if success:
            self.measured_rise_fall=True
            if display:
                messagebox.showinfo("成功", "Rise/Fall测量项添加成功\n")
            # 切换到电压选项卡
            self.tab_control.select(self.rise_fall_tab)
        else:
            messagebox.showerror("错误", message)

    def add_frequency_measurements(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        self.update_voltage()
        self.log_message("正在添加频率测量项...")
        if move:
            self.controller.zoom_middle()
        success, message = self.controller.add_frequency_measurements(frequency=self.wavelength_entry.get(), positive=self.pwd_entry.get(), negative=self.nwd_entry.get(), mode=self.frequency_mode.get())
        self.log_message(message)
        
        if success:
            self.measured_frequency=True
            if display:
                messagebox.showinfo("成功", "频率测量项添加成功\n")
            # 切换到频率选项卡
            self.tab_control.select(self.frequency_tab)
        else:
            messagebox.showerror("错误", message)

    def add_frequency_measurements_manual(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加频率测量项...")
        if move:
            self.controller.zoom_frequency()
        success, message = self.controller.add_frequency_measurements(frequency=self.wavelength_entry.get(), positive=self.pwd_entry.get(), negative=self.nwd_entry.get(), mode=self.frequency_mode.get())
        self.log_message(message)
        
        if success:
            self.measured_frequency=True
            if display:
                messagebox.showinfo("成功", "频率测量项添加成功\n")
            # 切换到频率选项卡
            self.tab_control.select(self.frequency_tab)
        else:
            messagebox.showerror("错误", message)

    def add_delay_measurement(self, display=True, move=False, mode="auto"):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Setup Time测量项...")
        numzoom=1 if mode=="auto" else 2
        if move:
            self.controller.zoom_setup_time(num_zoom=numzoom)
        success, message = self.controller.add_setup_measurement_manual(self.sda_mode.get(), self.scl_mode.get(), self.sda_edge_var.get(), self.scl_edge_var.get(), self.setup_ch2_entry.get(), self.setup_ch1_entry.get())
        self.log_message(message)
        
        if success:
            self.measured_setup=True
            if display:
                messagebox.showinfo("成功", "Setup Time测量项添加成功\n")
            # 切换到延迟选项卡
            self.tab_control.select(self.delay_tab)
        else:
            messagebox.showerror("错误", message)

    def add_hold_time_measurement(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Hold Time测量项...")
        if move:
            self.controller.zoom_hold_time()
        print(self.hold_ch2_entry.get())
        print(self.hold_ch1_entry.get())
        success, message = self.controller.add_hold_measurement_manual(self.sda_mode_hold.get(), self.scl_mode_hold.get(), self.sda_edge_var_hold.get(), self.scl_edge_var_hold.get(), self.hold_ch2_entry.get(), self.hold_ch1_entry.get())
        self.log_message(message)
        
        if success:
            self.measured_hold=True
            if display:
                messagebox.showinfo("成功", "Hold Time测量项添加成功\n")
                # 切换到延迟选项卡
            self.tab_control.select(self.hold_time_tab)
        else:
            messagebox.showerror("错误", message)

    def add_start_hold_time_measurement(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Start Hold Time测量项...")
        if move:
            self.controller.zoom_start()
        success, message = self.controller.add_delay2_measurement()
        self.log_message(message)
        
        if success:
            self.measured_start_hold=True
            if display:
                messagebox.showinfo("成功", "Start Hold Time测量项添加成功\n")
            # 切换到延迟选项卡
            #self.tab_control.select(self.start_hold_time_tab)
        else:
            messagebox.showerror("错误", message)

    def add_stop_setup_time_measurement(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Start Hold Time测量项...")
        if move:
            self.controller.zoom_stop()
        success, message = self.controller.add_stop_time_measurement()
        self.log_message(message)
        
        if success:
            self.measured_stop_setup=True
            if display: 
                messagebox.showinfo("成功", "Stop Setup Time测量项添加成功\n")
            # 切换到延迟选项卡
            #self.tab_control.select(self.stop_setup_time_tab)
        else:
            messagebox.showerror("错误", message)
    """
    def toggle_excel_path(self):
        if self.save_to_scope_excel.get():
            self.excel_path_entry.config(state="disabled")
            self.controller.excel_path = "oscilloscope"
            self.log_message("Excel将保存到示波器")
        else:
            self.excel_path_entry.config(state="normal")
            self.controller.excel_path = self.excel_path_entry.get()
            self.log_message(f"Excel将保存到本地路径: {self.controller.excel_path}")
    """
    def browse_excel_path(self):
        path = filedialog.askdirectory()
        newPath=path.replace('/', '\\')
        if path:
            self.excel_path_entry.delete(0, tk.END)
            self.excel_path_entry.insert(0, newPath)
            self.controller.excel_path = newPath
            self.log_message(f"已选择Excel保存路径: {newPath}")
    

    def create_voltage_tab(self):
        """创建电压测量参数表"""
        # CH1电压参数
        ttk.Label(self.voltage_tab, text="SCL 电压参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Maximum:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.scl_max_value = ttk.Label(self.voltage_tab, text="--")
        self.scl_max_value.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Minimum:").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.scl_min_value = ttk.Label(self.voltage_tab, text="--")
        self.scl_min_value.grid(row=2, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Top:").grid(row=3, column=0, padx=5, pady=2, sticky="w")
        self.scl_top_value = ttk.Label(self.voltage_tab, text="--")
        self.scl_top_value.grid(row=3, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Base:").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.scl_base_value = ttk.Label(self.voltage_tab, text="--")
        self.scl_base_value.grid(row=4, column=1, padx=5, pady=2, sticky="w")
        
        # CH2电压参数
        ttk.Label(self.voltage_tab, text="SDA 电压参数", font=("Arial", 10, "bold")).grid(row=0, column=2, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Maximum:").grid(row=1, column=2, padx=5, pady=2, sticky="w")
        self.sda_max_value = ttk.Label(self.voltage_tab, text="--")
        self.sda_max_value.grid(row=1, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Minimum:").grid(row=2, column=2, padx=5, pady=2, sticky="w")
        self.sda_min_value = ttk.Label(self.voltage_tab, text="--")
        self.sda_min_value.grid(row=2, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Top:").grid(row=3, column=2, padx=5, pady=2, sticky="w")
        self.sda_top_value = ttk.Label(self.voltage_tab, text="--")
        self.sda_top_value.grid(row=3, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Base:").grid(row=4, column=2, padx=5, pady=2, sticky="w")
        self.sda_base_value = ttk.Label(self.voltage_tab, text="--")
        self.sda_base_value.grid(row=4, column=3, padx=5, pady=2, sticky="w")
        
    

        self.get_voltage_scl_button = ttk.Button(self.voltage_tab, text="获取SCL电压测量值", command=lambda: self.add_voltage_measurements_by_channel(self.controller.channels['SCL'], True, True))
        self.get_voltage_scl_button.grid(row=5, column=0, padx=5, pady=10)

        self.get_voltage_sda_button = ttk.Button(self.voltage_tab, text="获取SDA电压测量值", command=lambda: self.add_voltage_measurements_by_channel(self.controller.channels['SDA'], True, True))
        self.get_voltage_sda_button.grid(row=5, column=2, padx=5, pady=10)

    
    def create_frequency_tab(self):
        """创建频率测量参数表"""

        # Section Title
        ttk.Label(self.frequency_tab, text="SCL 频率参数", font=("Arial", 10, "bold")).grid(
        row=0, column=0, columnspan=2, padx=5, pady=(10, 5), sticky="w"
        )

        # SDA Mode Selection
        ttk.Label(self.frequency_tab, text="Mode:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.frequency_mode = tk.StringVar()
        self.frequency_mode_combobox = ttk.Combobox(
            self.frequency_tab,
            textvariable=self.frequency_mode,
            values=["Percentage", "Absolute"],
            state="readonly",
            width=12
        )
        self.frequency_mode_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.frequency_mode_combobox.current(0)

        # Frequency Input Parameters
        ttk.Label(self.frequency_tab, text="SCL Frequency (Rise-Fall):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.wavelength_entry = ttk.Entry(self.frequency_tab, width=10)
        self.wavelength_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.wavelength_entry.insert(0, "30")

        ttk.Label(self.frequency_tab, text="SCL PWIDTH (Rise-Fall):").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.pwd_entry = ttk.Entry(self.frequency_tab, width=10)
        self.pwd_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.pwd_entry.insert(0, "70")

        ttk.Label(self.frequency_tab, text="SCL NWIDTH (Rise-Fall):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.nwd_entry = ttk.Entry(self.frequency_tab, width=10)
        self.nwd_entry.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        self.nwd_entry.insert(0, "30")

        # Measurement Results
        ttk.Label(self.frequency_tab, text="Frequency:").grid(row=5, column=0, padx=5, pady=2, sticky="w")
        self.frequency_value = ttk.Label(self.frequency_tab, text="--")
        self.frequency_value.grid(row=5, column=1, padx=5, pady=2, sticky="w")

        ttk.Label(self.frequency_tab, text="Positive Pulse Width:").grid(row=6, column=0, padx=5, pady=2, sticky="w")
        self.pos_width_value = ttk.Label(self.frequency_tab, text="--")
        self.pos_width_value.grid(row=6, column=1, padx=5, pady=2, sticky="w")

        ttk.Label(self.frequency_tab, text="Negative Pulse Width:").grid(row=7, column=0, padx=5, pady=2, sticky="w")
        self.neg_width_value = ttk.Label(self.frequency_tab, text="--")
        self.neg_width_value.grid(row=7, column=1, padx=5, pady=2, sticky="w")

        # Action Button
        self.get_frequency_button = ttk.Button(
            self.frequency_tab,
            text="获取频率测量值",
            command=self.frequency_group_manual
        )
        self.get_frequency_button.grid(row=8, column=0, padx=5, pady=(10, 10))
        
    def create_delay_tab(self):
        """创建setup time测量参数表"""
        ttk.Label(self.delay_tab, text="Setup Time参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")

        # Mode Dropdown
        ttk.Label(self.delay_tab, text="SDA Mode: ").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sda_mode = tk.StringVar()
        self.sda_mode_combobox = ttk.Combobox(self.delay_tab, textvariable=self.sda_mode, values=["Percentage", "Absolute"], state="readonly", width=10)
        self.sda_mode_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.sda_mode_combobox.current(0)  # default to "Percentage"

        ttk.Label(self.delay_tab, text="SCL Mode: ").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.scl_mode = tk.StringVar()
        self.scl_mode_combobox = ttk.Combobox(self.delay_tab, textvariable=self.scl_mode, values=["Percentage", "Absolute"], state="readonly", width=10)
        self.scl_mode_combobox.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        self.scl_mode_combobox.current(0)  # default to "Percentage"

        ttk.Label(self.delay_tab, text="SDA Edge: ").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.sda_edge_var = tk.StringVar()
        self.sda_edge_combobox = ttk.Combobox(self.delay_tab, textvariable=self.sda_edge_var, values=["Rise", "Fall"], state="readonly", width=8)
        self.sda_edge_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.sda_edge_combobox.current(0)  

        ttk.Label(self.delay_tab, text="SCL Edge: ").grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.scl_edge_var = tk.StringVar()
        self.scl_edge_combobox = ttk.Combobox(self.delay_tab, textvariable=self.scl_edge_var, values=["Rise", "Fall"], state="readonly", width=8)
        self.scl_edge_combobox.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        self.scl_edge_combobox.current(0)  # default to "rise"

        ttk.Label(self.delay_tab, text="SDA: Edge Mid").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.setup_ch2_entry = ttk.Entry(self.delay_tab, width=10)
        self.setup_ch2_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.setup_ch2_entry.insert(0, "70")  # 默认
        
        ttk.Label(self.delay_tab, text="SCL: Edge Mid").grid(row=3, column=2, padx=5, pady=5, sticky="w")
        self.setup_ch1_entry = ttk.Entry(self.delay_tab, width=10)
        self.setup_ch1_entry.grid(row=3, column=3, padx=5, pady=5, sticky="w")
        self.setup_ch1_entry.insert(0, "30")  # 默认
        
        ttk.Label(self.delay_tab, text="Setup Time:").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.delay_value = ttk.Label(self.delay_tab, text="--")
        self.delay_value.grid(row=4, column=2, padx=5, pady=2, sticky="w")

        # 获取setup time测量值按钮
        self.get_delay_button = ttk.Button(self.delay_tab, text="获取Setup Time测量值", command=self.setup_group_manual)
        self.get_delay_button.grid(row=5, column=0, padx=5, pady=5)
    
    def create_hold_time_tab(self):
        """创建hold time测量参数表"""
        ttk.Label(self.hold_time_tab, text="Hold Time参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")

        # Mode Dropdown
        ttk.Label(self.hold_time_tab, text="SDA Mode: ").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.sda_mode_hold = tk.StringVar()
        self.sda_mode_hold_combobox = ttk.Combobox(self.hold_time_tab, textvariable=self.sda_mode_hold, values=["Percentage", "Absolute"], state="readonly", width=10)
        self.sda_mode_hold_combobox.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        self.sda_mode_hold_combobox.current(0)  # default to "Percentage"

        ttk.Label(self.hold_time_tab, text="SCL Mode: ").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.scl_mode_hold = tk.StringVar()
        self.scl_mode_hold_combobox = ttk.Combobox(self.hold_time_tab, textvariable=self.scl_mode_hold, values=["Percentage", "Absolute"], state="readonly", width=10)
        self.scl_mode_hold_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.scl_mode_hold_combobox.current(0)  # default to "Percentage"

        ttk.Label(self.hold_time_tab, text="SDA Edge: ").grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.sda_edge_var_hold = tk.StringVar()
        self.sda_edge_hold_combobox = ttk.Combobox(self.hold_time_tab, textvariable=self.sda_edge_var_hold, values=["Rise", "Fall"], state="readonly", width=8)
        self.sda_edge_hold_combobox.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        self.sda_edge_hold_combobox.current(1)  

        ttk.Label(self.hold_time_tab, text="SCL Edge: ").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.scl_edge_var_hold = tk.StringVar()
        self.scl_edge_hold_combobox = ttk.Combobox(self.hold_time_tab, textvariable=self.scl_edge_var_hold, values=["Rise", "Fall"], state="readonly", width=8)
        self.scl_edge_hold_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.scl_edge_hold_combobox.current(1)  # default to "rise"

        ttk.Label(self.hold_time_tab, text="SDA: Edge Mid").grid(row=3, column=2, padx=5, pady=5, sticky="w")
        self.hold_ch2_entry = ttk.Entry(self.hold_time_tab, width=10)
        self.hold_ch2_entry.grid(row=3, column=3, padx=5, pady=5, sticky="w")
        self.hold_ch2_entry.insert(0, "70")  # 默认
        
        ttk.Label(self.hold_time_tab, text="SCL: Edge Mid").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.hold_ch1_entry = ttk.Entry(self.hold_time_tab, width=10)
        self.hold_ch1_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.hold_ch1_entry.insert(0, "30")  # 默认
        
        ttk.Label(self.hold_time_tab, text="Hold Time:").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.hold_time_value = ttk.Label(self.hold_time_tab, text="--")
        self.hold_time_value.grid(row=4, column=1, padx=5, pady=2, sticky="w")
        
        # 获取setup time测量值按钮
        self.get_hold_time_button = ttk.Button(self.hold_time_tab, text="获取Hold Time测量值", command=self.write_hold_group)
        self.get_hold_time_button.grid(row=5, column=0, columnspan=1, padx=5, pady=5)
    """
    def create_start_hold_time_tab(self):
        
        ttk.Label(self.start_hold_time_tab, text="Start Hold Time参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.start_hold_time_tab, text="Start Hold Time (SDA下降沿70% 到 SCL下降沿30%):").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.start_hold_time_value = ttk.Label(self.start_hold_time_tab, text="--")
        self.start_hold_time_value.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        # 获取setup time测量值按钮
        self.get_start_hold_time_button = ttk.Button(self.start_hold_time_tab, text="获取Start Hold Time测量值", command=self.get_start_hold_time_measurement)
        self.get_start_hold_time_button.grid(row=2, column=1, columnspan=2, padx=5, pady=5)

        # Screenshot
        self.start_hold_time_screenshot_button = ttk.Button(self.start_hold_time_tab, text="Save Screenshot", command=lambda: self.save_screenshot("Start Hold Time"))
        self.start_hold_time_screenshot_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5)
    
    def create_stop_setup_time_tab(self):
        
        ttk.Label(self.stop_setup_time_tab, text="Stop Setup Time参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.stop_setup_time_tab, text="Stop Setup Time (SC:上升沿30% 到 SDA上升沿70%):").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.stop_setup_time_value = ttk.Label(self.stop_setup_time_tab, text="--")
        self.stop_setup_time_value.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        # 获取setup time测量值按钮
        self.get_stop_setup_time_button = ttk.Button(self.stop_setup_time_tab, text="获取Stop Setup Time测量值", command=self.get_stop_setup_time_measurement)
        self.get_stop_setup_time_button.grid(row=2, column=1, columnspan=2, padx=5, pady=5)

        # Screenshot
        self.stop_setup_time_screenshot_button = ttk.Button(self.stop_setup_time_tab, text="Save Screenshot", command=lambda: self.save_screenshot("Stop Setup Time"))
        self.stop_setup_time_screenshot_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5)
    """
    def create_rise_fall_tab(self):
        """创建rise/fall测量参数表"""
        # CH1电压参数
        ttk.Label(self.rise_fall_tab, text="Rise/Fall Times", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")

        ttk.Label(self.rise_fall_tab, text="Rise Mode:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.rise_mode = tk.StringVar()
        self.rise_mode_combobox = ttk.Combobox(
            self.rise_fall_tab,
            textvariable=self.rise_mode,
            values=["Percentage", "Absolute"],
            state="readonly",
            width=12
        )
        self.rise_mode_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.rise_mode_combobox.current(0)

        ttk.Label(self.rise_fall_tab, text="Fall Mode:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.fall_mode = tk.StringVar()
        self.fall_mode_combobox = ttk.Combobox(
            self.rise_fall_tab,
            textvariable=self.fall_mode,
            values=["Percentage", "Absolute"],
            state="readonly",
            width=12
        )
        self.fall_mode_combobox.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        self.fall_mode_combobox.current(0)

        ttk.Label(self.rise_fall_tab, text="Rising Edge: Low").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.rising_edge_low_entry = ttk.Entry(self.rise_fall_tab, width=10)
        self.rising_edge_low_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.rising_edge_low_entry.insert(0, "30")  # 默认

        ttk.Label(self.rise_fall_tab, text="Rising Edge: High").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.rising_edge_high_entry = ttk.Entry(self.rise_fall_tab, width=10)
        self.rising_edge_high_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.rising_edge_high_entry.insert(0, "70")  # 默认

        ttk.Label(self.rise_fall_tab, text="Falling Edge: High").grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.falling_edge_high_entry = ttk.Entry(self.rise_fall_tab, width=10)
        self.falling_edge_high_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        self.falling_edge_high_entry.insert(0, "70")  # 默认

        ttk.Label(self.rise_fall_tab, text="Falling Edge: Low").grid(row=3, column=2, padx=5, pady=5, sticky="w")
        self.falling_edge_low_entry = ttk.Entry(self.rise_fall_tab, width=10)
        self.falling_edge_low_entry.grid(row=3, column=3, padx=5, pady=5, sticky="w")
        self.falling_edge_low_entry.insert(0, "30")  # 默认
        
        ttk.Label(self.rise_fall_tab, text="SCL Risetime: ").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.ch1_risetime = ttk.Label(self.rise_fall_tab, text="--")
        self.ch1_risetime.grid(row=4, column=1, padx=5, pady=2, sticky="w")

        ttk.Label(self.rise_fall_tab, text="SDA Risetime: ").grid(row=5, column=0, padx=5, pady=2, sticky="w")
        self.ch2_risetime = ttk.Label(self.rise_fall_tab, text="--")
        self.ch2_risetime.grid(row=5, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.rise_fall_tab, text="SCL Falltime: ").grid(row=4, column=2, padx=5, pady=2, sticky="w")
        self.ch1_falltime = ttk.Label(self.rise_fall_tab, text="--")
        self.ch1_falltime.grid(row=4, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.rise_fall_tab, text="SDA Falltime: ").grid(row=5, column=2, padx=5, pady=2, sticky="w")
        self.ch2_falltime = ttk.Label(self.rise_fall_tab, text="--")
        self.ch2_falltime.grid(row=5, column=3, padx=5, pady=2, sticky="w")
        
        # 获取电压测量值按钮
        self.get_rise_button = ttk.Button(self.rise_fall_tab, text="获取Rise测量值", command=self.rise_button)
        self.get_rise_button.grid(row=6, column=0, padx=5, pady=10)

        # 获取电压测量值按钮
        self.get_fall_button = ttk.Button(self.rise_fall_tab, text="获取Fall测量值", command=self.fall_button)
        self.get_fall_button.grid(row=6, column=2, padx=5, pady=10)

        # Screenshot
        #self.get_rise_fall_screenshot_button = ttk.Button(self.rise_fall_tab, text="Save Screenshot", command=lambda: self.save_screenshot("Rise Fall"))
        #self.get_rise_fall_screenshot_button.grid(row=6, column=0, columnspan=4, padx=10, pady=10)

    def get_voltage_measurements(self, display=True):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.measured_voltage==False:
            messagebox.showerror("错误", "Voltage has not been measured")
            return
        self.log_message("正在获取电压测量值...")
        success, message, measurements = self.controller.get_voltage_measurements()
        self.res.update(measurements)
        self.log_message(message)
        
        if success:
            # 更新电压测量值显示
            self.scl_max_value.config(text=f"{measurements['SCL_Maximum']} V")
            self.scl_min_value.config(text=f"{measurements['SCL_Minimum']} V")
            self.scl_top_value.config(text=f"{measurements['SCL_Top']} V")
            self.scl_base_value.config(text=f"{measurements['SCL_Base']} V")
            
            self.sda_max_value.config(text=f"{measurements['SDA_Maximum']} V")
            self.sda_min_value.config(text=f"{measurements['SDA_Minimum']} V")
            self.sda_top_value.config(text=f"{measurements['SDA_Top']} V")
            self.sda_base_value.config(text=f"{measurements['SDA_Base']} V")
            if display:
                messagebox.showinfo("成功", "电压测量值获取成功")
        else:
            messagebox.showerror("错误", message)
    
    def get_rise_fall_measurements(self, display=True):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.measured_rise_fall==False:
            messagebox.showerror("错误", "Rise/Fall has not been measured")
            return
        self.log_message("正在获取Rise/Fall测量值...")
        success, message, measurements = self.controller.get_rise_fall_measurements()
        self.res.update(measurements)
        self.log_message(message)
        
        if success:
            # 更新电压测量值显示
            self.ch1_risetime.config(text=f"{measurements['SCL_Rise']} ns")
            self.ch1_falltime.config(text=f"{measurements['SCL_Fall']} ns")
            self.ch2_risetime.config(text=f"{measurements['SDA_Rise']} ns")
            self.ch2_falltime.config(text=f"{measurements['SDA_Fall']} ns")
            if display:
                messagebox.showinfo("成功", "Rise/Fall测量值获取成功")
        else:
            messagebox.showerror("错误", message)
    
    def get_frequency_measurements(self, display=True):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.measured_frequency==False:
            messagebox.showerror("错误", "Frequency has not been measured")
            return
        self.log_message("正在获取频率测量值...")
        success, message, measurements = self.controller.get_frequency_measurements()
        self.res.update(measurements)
        self.log_message(message)
        
        if success:
            # 更新频率测量值显示
            self.frequency_value.config(text=f"{measurements['SCL Frequency']} KHz")
            self.pos_width_value.config(text=f"{measurements['SCL Positive_Pulse_Width']} us")
            self.neg_width_value.config(text=f"{measurements['SCL Negative_Pulse_Width']} us")
            if display:
                messagebox.showinfo("成功", "频率测量值获取成功")
        else:
            messagebox.showerror("错误", message)
    
    def get_delay_measurement(self, display=True):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.measured_setup==False:
            messagebox.showerror("错误", "Setup Time has not been measured")
            return
        self.log_message("正在获取Setup Time测量值...")
        success, message, measurements = self.controller.get_delay_measurement()
        self.res.update(measurements)
        self.log_message(message)
        
        if success:
            # 更新延迟测量值显示
            self.delay_value.config(text=f"{measurements['SDA Setup Time']} us")
            if display:
                messagebox.showinfo("成功", "Setup Time测量值获取成功")
        else:
            messagebox.showerror("错误", message)

    def get_hold_time_measurement(self, display=True):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.measured_hold==False:
            messagebox.showerror("错误", "Hold Time has not been measured")
            return
        self.log_message("正在获取Hold Time测量值...")
        success, message, measurements = self.controller.get_hold_measurement()
        self.res.update(measurements)
        self.log_message(message)
        
        if success:
            # 更新延迟测量值显示
            self.hold_time_value.config(text=f"{measurements['SDA Hold Time']} us")
            if display:
                messagebox.showinfo("成功", "Hold Time测量值获取成功")
        else:
            messagebox.showerror("错误", message)

    def get_start_hold_time_measurement(self, display=True):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.measured_start_hold==False:
            messagebox.showerror("错误", "Start Hold Time has not been measured")
            return
        self.log_message("正在获取Start Hold Time测量值...")
        success, message, measurements = self.controller.get_start_hold_time()
        self.res.update(measurements)
        self.log_message(message)
        
        if success:
            # 更新延迟测量值显示
            #self.start_hold_time_value.config(text=f"{measurements['Start Hold Time']} us")
            if display:
                messagebox.showinfo("成功", "Start Hold Time测量值获取成功")
        else:
            messagebox.showerror("错误", message)

    def get_stop_setup_time_measurement(self, display=True):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.measured_stop_setup==False:
            messagebox.showerror("错误", "Stop Setup Time has not been measured")
            return
        self.log_message("正在获取Stop Setup Time测量值...")
        success, message, measurements = self.controller.get_stop_time_measurement()
        self.res.update(measurements)
        self.log_message(message)
        
        if success:
            # 更新延迟测量值显示
            #self.stop_setup_time_value.config(text=f"{measurements['Stop Setup Time']} us")
            if display:
                messagebox.showinfo("成功", "Stop Setup Time测量值获取成功")
        else:
            messagebox.showerror("错误", message)

    def save_screenshot(self, filename, end=False, auto=False):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        filenameCopy=filename
        foldername=self.folder_entry.get()
        suffix=self.suffix_entry.get()
        if suffix=="" and auto==False:
            filename+=f"_{datetime.now().strftime("%H-%M-%S")}"
        elif len(suffix)>0:
            filename+="_"+suffix
        if foldername=="":
            current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            foldername=f"Test{current_time}"
            self.folder_entry.insert(0, f"Test{current_time}")
        suc, msg=self.controller.save_screenshot(filename, foldername)
        if suc:
            if foldername not in self.controller.image_path:
                self.images[filenameCopy]=self.controller.image_path+f"\\{foldername}\\{filename}.png"
            else:
                self.images[filenameCopy]=self.controller.image_path+f"\\{filename}.png"
            if end==True:
                self.suffix_entry.delete(0, tk.END)
        self.log_message(msg)

    def save_excel(self):
        print(self.controller.excel_path)
        if self.controller.excel_path=="":
            messagebox.showerror("错误", "File path has not been configured")
            return
        foldername=""
        if self.folder_entry.get()!="":
            foldername=self.folder_entry.get()
        else:
            current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            time = datetime.now().strftime("%H-%M-%S")
            foldername=f"Test{current_time}"
            self.folder_entry.insert(0, f"Test{current_time}")
        time = datetime.now().strftime("%H-%M-%S")
        suc, msg=self.controller.save_data_to_excel(self.res, time, foldername, self.images)
        self.images={}
        self.res={}
        self.set_max_voltage=False
        self.log_message(msg)

    def update_button_states(self, connected):
        state = "normal" if connected else "disabled"
        self.voltage_button.config(state=state)
        self.frequency_button.config(state=state)
        self.delay_button.config(state=state)
        self.get_frequency_button.config(state=state)
        self.get_delay_button.config(state=state)
        self.rise_fall_button.config(state=state)
        #self.get_rise_fall_button.config(state=state)
        #self.get_rise_fall_screenshot_button.config(state=state)
        self.hold_time_button.config(state=state)
        self.get_hold_time_button.config(state=state)
        self.start_hold_time_button.config(state=state)
        #self.get_stop_setup_time_button.config(state=state)
        #self.stop_setup_time_screenshot_button.config(state=state)
        self.stop_setup_time_button.config(state=state)
        #self.get_start_hold_time_button.config(state=state)
        #self.start_hold_time_screenshot_button.config(state=state)
        self.run_all_tests_button.config(state=state)
        self.run_read_tests_button.config(state=state)
        self.voltage_button_read.config(state=state)
        self.rise_fall_button_read.config(state=state)
        self.setup_button_read.config(state=state)
        self.hold_button_read.config(state=state)
        self.get_rise_button.config(state=state)
        self.get_fall_button.config(state=state)
        self.get_voltage_sda_button.config(state=state)
        self.get_voltage_scl_button.config(state=state)

    def on_channel_change(self, *args):
        clock = self.clock_var.get()
        data = self.data_var.get()

        # Prevent clock and data from being the same
        if clock == data:
            # Revert the most recent change
            if args[0] == str(self.clock_var):
                # clock_var just changed → revert it
                for ch in self.clock_options:
                    if ch != data:
                        self.clock_var.set(ch)
                        break
            elif args[0] == str(self.data_var):
                # data_var just changed → revert it
                for ch in self.data_options:
                    if ch != clock:
                        self.data_var.set(ch)
                        break

        # Update internal state
        self.clock = self.clock_var.get()
        self.data = self.data_var.get()

        self.controller.channels['SCL']=self.clock_var.get()
        self.controller.channels['SDA']=self.data_var.get()

    def add_voltage_measurements_read(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加电压测量项...")
        if move:
            self.controller.zoom_middle_read()
        success, message = self.controller.add_voltage_measurements()
        self.log_message(message)
        
        if success:
            self.measured_voltage_read=True
            if display:
                messagebox.showinfo("成功", "电压测量项添加成功\n")
                # 切换到电压选项卡
            self.tab_control.select(self.voltage_tab)
        else:
            messagebox.showerror("错误", message)

    def get_voltage_measurements_read(self, display=True):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.measured_voltage_read==False:
            messagebox.showerror("错误", "Voltage has not been measured")
            return
        self.log_message("正在获取电压测量值...")
        success, message, measurements = self.controller.get_voltage_measurements()
        self.res_read.update(measurements)
        self.log_message(message)
        
        if success:
            # 更新电压测量值显示
            self.scl_max_value.config(text=f"{measurements['SCL_Maximum']} V")
            self.scl_min_value.config(text=f"{measurements['SCL_Minimum']} V")
            self.scl_top_value.config(text=f"{measurements['SCL_Top']} V")
            self.scl_base_value.config(text=f"{measurements['SCL_Base']} V")
            
            self.sda_max_value.config(text=f"{measurements['SDA_Maximum']} V")
            self.sda_min_value.config(text=f"{measurements['SDA_Minimum']} V")
            self.sda_top_value.config(text=f"{measurements['SDA_Top']} V")
            self.sda_base_value.config(text=f"{measurements['SDA_Base']} V")
            if display:
                messagebox.showinfo("成功", "电压测量值获取成功")
        else:
            messagebox.showerror("错误", message)
    
    def add_rise_fall_measurements_read(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Rise/Fall测量项...")
        if move:
            self.controller.zoom_middle_read()
        success, message = self.controller.add_rise_fall_measurements()
        self.log_message(message)
        
        if success:
            self.measured_rise_fall_read=True
            if display:
                messagebox.showinfo("成功", "Rise/Fall测量项添加成功\n")
                # 切换到电压选项卡
            self.tab_control.select(self.rise_fall_tab)
        else:
            messagebox.showerror("错误", message)

    def get_rise_fall_measurements_read(self, display=True):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.measured_rise_fall_read==False:
            messagebox.showerror("错误", "Rise/Fall has not been measured")
            return
        self.log_message("正在获取Rise/Fall测量值...")
        success, message, measurements = self.controller.get_rise_fall_measurements()
        self.res_read.update(measurements)
        self.log_message(message)
        
        if success:
            # 更新电压测量值显示
            self.ch1_risetime.config(text=f"{measurements['SCL_Rise']} ns")
            self.ch1_falltime.config(text=f"{measurements['SCL_Fall']} ns")
            self.ch2_risetime.config(text=f"{measurements['SDA_Rise']} ns")
            self.ch2_falltime.config(text=f"{measurements['SDA_Fall']} ns")
            if display:
                messagebox.showinfo("成功", "Rise/Fall测量值获取成功")
        else:
            messagebox.showerror("错误", message)
    
    def add_setup_measurements_read(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Setup Time测量项...")
        if move:
            self.controller.zoom_setup_time_read()
        success, message = self.controller.add_delay_measurement()
        self.log_message(message)
        
        if success:
            self.measured_setup_read=True
            if display:
                messagebox.showinfo("成功", "Setup Time测量项添加成功\n")
                # 切换到电压选项卡
            self.tab_control.select(self.delay_tab)
        else:
            messagebox.showerror("错误", message)
            return
        
    def get_setup_measurement_read(self, display=True):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.measured_setup_read==False:
            messagebox.showerror("错误", "Setup Time has not been measured")
            return
        self.log_message("正在获取Setup Time测量值...")
        success, message, measurements = self.controller.get_delay_measurement()
        self.res_read.update(measurements)
        self.log_message(message)
        
        if success:
            # 更新延迟测量值显示
            self.delay_value.config(text=f"{measurements['SDA Setup Time']} ns")
            if display:
                messagebox.showinfo("成功", "Setup Time测量值获取成功")
        else:
            messagebox.showerror("错误", message)
    
    def add_hold_measurement_read(self, display=True, move=False):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Hold Time测量项...")
        if move:
            self.controller.zoom_hold_time_read()
        success, message = self.controller.add_hold_time_measurements()
        self.log_message(message)
        
        if success:
            self.measured_hold_read=True
            if display:
                messagebox.showinfo("成功", "电压测量项添加成功\n")
                # 切换到电压选项卡
            self.tab_control.select(self.hold_time_tab)
        else:
            messagebox.showerror("错误", message)
            return
    
    def get_hold_time_measurement_read(self, display=True):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        if self.measured_hold_read==False:
            messagebox.showerror("错误", "Hold Time has not been measured")
            return
        self.log_message("正在获取Hold Time测量值...")
        success, message, measurements = self.controller.get_hold_measurement()
        self.res_read.update(measurements)
        self.log_message(message)
        
        if success:
            # 更新延迟测量值显示
            self.hold_time_value.config(text=f"{measurements['SDA Hold Time']} s")
            if display:
                messagebox.showinfo("成功", "Hold Time测量值获取成功")
        else:
            messagebox.showerror("错误", message)

    def save_screenshot_read(self, filename, end=False, auto=False):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        filenameCopy=filename
        suffix=self.suffix_entry.get()
        if suffix=="" and auto==False:
            filename+=f"_{datetime.now().strftime("%H-%M-%S")}"
        elif len(suffix)>0:
            filename+="_"+suffix
        foldername=self.folder_entry.get()
        if foldername=="":
            current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            foldername=f"Test{current_time}"
            self.folder_entry.insert(0, f"Test{current_time}")
        if auto==False:
            suc, msg=self.controller.save_screenshot(filename, foldername)
        else:
            suc, msg=self.controller.save_screenshot(filename, foldername)
        if suc:
            if foldername not in self.controller.image_path:
                self.images_read[filenameCopy]=self.controller.image_path+f"\\{foldername}\\{filename}.png"
            else:
                self.images_read[filenameCopy]=self.controller.image_path+f"\\{filename}.png"
            if end:
                self.suffix_entry.delete(0, tk.END)
        #print(self.images)
        self.log_message(msg)

    def save_excel_read(self):
        if self.controller.excel_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        foldername=""
        if self.folder_entry.get()!="":
            foldername=self.folder_entry.get()
        else:
            current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            time = datetime.now().strftime("%H-%M-%S")
            foldername=f"Test{current_time}"
            self.folder_entry.insert(0, f"Test{current_time}")
        time = datetime.now().strftime("%H-%M-%S")
        suc, msg=self.controller.save_data_to_excel(self.res_read, time, foldername, self.images_read)
        self.res_read={}
        self.images_read={}
        self.log_message(msg)

    
    def run_read_tests(self):
        if self.controller.excel_path is None:
            messagebox.showerror("错误", "File path has not been configured!")
            return
        if self.controller.image_path is None:
            messagebox.showerror("错误", "Image path has not been configured!")
            return
        self.add_voltage_measurements_read(False, True)
        self.controller.scope.query("*OPC?")
        self.get_voltage_measurements_read(False)
        self.controller.scope.query("*OPC?")
        time.sleep(0.5)
        self.save_screenshot_read("Voltage", auto=True)
        self.controller.scope.query("*OPC?")
        self.add_rise_fall_measurements_read(False, True)
        self.controller.scope.query("*OPC?")
        self.get_rise_fall_measurements_read(False)
        self.controller.scope.query("*OPC?")
        time.sleep(0.5)
        self.save_screenshot_read("Rise Fall", auto=True)
        self.controller.scope.query("*OPC?")
        self.add_setup_measurements_read(False, True)
        self.controller.scope.query("*OPC?")
        self.get_setup_measurement_read(False)
        self.controller.scope.query("*OPC?")
        time.sleep(0.5)
        self.save_screenshot_read("Setup Time", auto=True)
        self.controller.scope.query("*OPC?")
        self.add_hold_measurement_read(False, True)
        self.controller.scope.query("*OPC?")
        self.get_hold_time_measurement_read(False)
        self.controller.scope.query("*OPC?")
        time.sleep(0.5)
        self.save_screenshot_read("Hold Time", end=True, auto=True)
        self.controller.scope.query("*OPC?")
        self.save_excel_read()
        self.folder_entry.delete(0, tk.END)
    
    def write_voltage_group(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_voltage_measurements(False, True)
        self.controller.scope.query("*OPC?")
        self.get_voltage_measurements()
        self.controller.scope.query("*OPC?")
        self.save_screenshot("Voltage")

    def read_voltage_group(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_voltage_measurements_read(False, True)
        self.controller.scope.query("*OPC?")
        self.get_voltage_measurements_read()
        self.controller.scope.query("*OPC?")
        self.save_screenshot_read("Voltage")

    def write_frequency_group(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_frequency_measurements(False, True)
        self.controller.scope.query("*OPC?")
        self.get_frequency_measurements()
        self.controller.scope.query("*OPC?")
        self.save_screenshot("Frequency")

    def write_rise_fall_group(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_rise_fall_measurements(False, True)
        self.controller.scope.query("*OPC?")
        self.get_rise_fall_measurements()
        self.controller.scope.query("*OPC?")
        self.save_screenshot("Rise Fall")

    def read_rise_fall_group(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_rise_fall_measurements_read(False, True)
        self.controller.scope.query("*OPC?")
        self.get_rise_fall_measurements_read()
        self.controller.scope.query("*OPC?")
        self.save_screenshot_read("Rise Fall")

    def write_setup_group(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_delay_measurement(False, True)
        self.controller.scope.query("*OPC?")
        self.get_delay_measurement()
        self.controller.scope.query("*OPC?")
        self.save_screenshot("Setup Time")
    
    def setup_group_manual(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_delay_measurement(False, True, "manual")
        self.controller.scope.query("*OPC?")
        self.get_delay_measurement()
        self.controller.scope.query("*OPC?")
        self.save_screenshot("Setup Time")

    def read_setup_group(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_setup_measurements_read(False, True)
        self.controller.scope.query("*OPC?")
        self.get_setup_measurement_read()
        self.controller.scope.query("*OPC?")
        self.save_screenshot_read("Setup Time")

    def write_hold_group(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_hold_time_measurement(False, True)
        self.controller.scope.query("*OPC?")
        self.get_hold_time_measurement()
        self.controller.scope.query("*OPC?")
        self.save_screenshot("Hold Time")

    def read_hold_group(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_hold_measurement_read(False, True)
        self.controller.scope.query("*OPC?")
        self.get_hold_time_measurement_read()
        self.controller.scope.query("*OPC?")
        self.save_screenshot_read("Hold Time")

    def write_start_hold_group(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_start_hold_time_measurement(False, True)
        self.controller.scope.query("*OPC?")
        self.get_start_hold_time_measurement()
        self.controller.scope.query("*OPC?")
        self.save_screenshot("Start Hold Time")

    def write_stop_setup_group(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        self.add_stop_setup_time_measurement(False, True)
        self.controller.scope.query("*OPC?")
        self.get_stop_setup_time_measurement()
        self.controller.scope.query("*OPC?")
        self.save_screenshot("Stop Setup Time")

    def frequency_group_manual(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "Image file path has not been configured")
            return
        self.add_frequency_measurements_manual(False, True)
        self.controller.scope.query("*OPC?")
        self.get_frequency_measurements()
        self.controller.scope.query("*OPC?")
        self.save_screenshot("Frequency")

    def rise_button(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "Image file path has not been configured")
            return
        self.controller.zoom_rise()
        suc, msg=self.controller.add_rise_measurements(self.rising_edge_high_entry.get(), self.rising_edge_low_entry.get(), self.rise_mode.get())
        if suc:
            success, message, measurements = self.controller.get_rise_measurements()
            print(measurements)
            self.res.update(measurements)
            self.log_message(message)
            if success:
                # 更新电压测量值显示
                self.ch1_risetime.config(text=f"{measurements['SCL_Rise']} ns")
                self.ch2_risetime.config(text=f"{measurements['SDA_Rise']} ns")
                messagebox.showinfo("成功", "Rise测量值获取成功")
                self.save_screenshot("Rising Edge")
            else:
                messagebox.showerror("错误", message)
        else:
            messagebox.showerror("错误", msg)

    def fall_button(self):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "Image file path has not been configured")
            return
        self.controller.zoom_fall()
        suc, msg=self.controller.add_fall_measurements(self.falling_edge_high_entry.get(), self.falling_edge_low_entry.get(), self.fall_mode.get())
        if suc:
            success, message, measurements = self.controller.get_fall_measurements()
            self.res.update(measurements)
            self.log_message(message)
            if success:
                # 更新电压测量值显示
                self.ch1_falltime.config(text=f"{measurements['SCL_Fall']} ns")
                self.ch2_falltime.config(text=f"{measurements['SDA_Fall']} ns")
                messagebox.showinfo("成功", "Fall测量值获取成功")
                self.save_screenshot("Falling Edge")
            else:
                messagebox.showerror("错误", message)
        else:
            messagebox.showerror("错误", msg)
        
    def on_voltage_change(self, *args):
        # Update internal state
        self.voltage = self.voltage_var.get()
        self.set_max_voltage=True
        print("changed")
    
    def update_voltage(self):
        if self.set_max_voltage==False:
            self.controller.zoom_middle()
            #add one measurement
            self.controller.scope.write("MEASUREMENT:ADDMEAS MAXIMUM")
            self.controller.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.clock}")
            self.controller.scope.query("*OPC?")
            #get the measurement now
            voltage=float(self.controller.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            #set acceptable range as +-0.5V
            if voltage>=2.8:
                self.voltage_var.set("3.3V")
            else:
                self.voltage_var.set("1.8V")
            self.set_max_voltage=True

if __name__ == "__main__":
    root = tk.Tk()
    app = MSO64ControllerGUI(root)
    root.mainloop() 

"""
controller=MSO64Controller()
controller.connect('169.254.103.178')
controller.channels['SDA']="REF2"
controller.channels['SCL']="REF1"
suc, msg=controller.zoom_on_falling_edge()
"""
