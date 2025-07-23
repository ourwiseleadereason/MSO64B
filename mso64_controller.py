import os
import time
import pandas as pd
from datetime import datetime
import pyvisa
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import math
from PIL import Image

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
        self.channels={}
        self.wavelength=None

    def connect(self, ip_address):
        """连接到示波器"""
        try:
            resource_string = f"TCPIP::{ip_address}::INSTR"
            self.scope = self.rm.open_resource(resource_string)
            self.scope.timeout = 5000  # 设置超时时间为10秒
            idn = self.scope.query("*IDN?")
            self.scope.write('FILESystem:MKDir "C:\\Test"')
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
        
    def measure_voltage(self):
        """添加电压测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        
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
        results={}
        try:
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
        
    def meausre_voltage_by_channel(self, channel):
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
            return True, "电压测量项添加成功", results
        except Exception as e:
            return False, f"添加电压测量项失败: {str(e)}", {}
        
    def measure_frequency(self, frequency=30, positive=70, negative=30, mode="Percentage"):
        """添加频率和脉宽测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        
        # 清除所有现有测量
        self.clear_measurements()
        method="PERCent" if mode=="Percentage" else "ABSOLUTE"
        self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
        self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
        self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
        self.scope.write(f"MEASUREMENT:MEAS1:REFLEVELS1:METHOD {method}")
        self.scope.write("MEASUrement:GATing SCREEN")
        self.scope.write("MEASUrement:MEAS1:GATing SCREEN")

        self.scope.write(f"MEASUREMENT:MEAS1:REFLevels1:{method}:RISEHigh {frequency}")
        self.scope.write(f"MEASUREMENT:MEAS1:REFLevels1:{method}:RISEMid {frequency}")
        self.scope.write(f"MEASUREMENT:MEAS1:REFLevels1:{method}:RISELow {frequency}")

        # 正脉宽 (从70%到70%)
        self.scope.write("MEASUREMENT:ADDMEAS PWIDTH")
        self.scope.write(f"MEASUREMENT:MEAS2:SOURCE {self.channels['SCL']}")
        # 设置参考电平方法为百分比
        self.scope.write(f"MEASUREMENT:MEAS2:REFLevels:METHod {method}")
        self.scope.write("MEASUrement:MEAS2:GLOBalref 0")
        self.scope.write("MEASUrement:MEAS2:GATing SCREEN")
        # 设置上升沿参考电平
        self.scope.write(f"MEASUREMENT:MEAS2:REFLevels1:{method}:RISEHigh {positive}")
        self.scope.write(f"MEASUREMENT:MEAS2:REFLevels1:{method}:RISEMid {positive}")
        self.scope.write(f"MEASUREMENT:MEAS2:REFLevels1:{method}:RISELow {positive}")
        # 设置下降沿参考电平
        self.scope.write(f"MEASUREMENT:MEAS2:REFLevels1:{method}:FALLHigh {positive}")
        self.scope.write(f"MEASUREMENT:MEAS2:REFLevels1:{method}:FALLMid {positive}")
        self.scope.write(f"MEASUREMENT:MEAS2:REFLevels1:{method}:FALLLow {positive}")
        
        # 负脉宽 (从30%到30%)
        self.scope.write("MEASUREMENT:ADDMEAS NWIDTH")
        self.scope.write(f"MEASUREMENT:MEAS3:SOURCE {self.channels['SCL']}")
        # 设置参考电平方法为百分比
        self.scope.write(f"MEASUREMENT:MEAS3:REFLevels1:METHod {method}")
        self.scope.write("MEASUrement:MEAS3:GLOBalref 0")
        self.scope.write("MEASUrement:MEAS3:GATing SCREEN")
        # 设置上升沿参考电平
        self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:RISEHigh {negative}")
        self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:RISEMid {negative}")
        self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:RISELow {negative}")
        # 设置下降沿参考电平
        self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:FALLHigh {negative}")
        self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:FALLMid {negative}")
        self.scope.write(f"MEASUREMENT:MEAS3:REFLevels:{method}:FALLLow {negative}")
        self.scope.query("*OPC?")
        results={}
        try:
            results["SCL Frequency"] = round_sig(adapt_frequency(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            results["SCL Positive_Pulse_Width"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))))
            results["SCL Negative_Pulse_Width"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))))
            return True, "频率和脉宽测量项添加成功", results
        except Exception as e:
            return False, f"添加频率和脉宽测量项失败: {str(e)}", results
        
    def measure_setup_time(self, sda_mode, scl_mode, sda_edge, scl_edge, sda_edge_level, scl_edge_level):
        if not self.connected:
            return False, "未连接到示波器"
    
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
            edge_dir = "RISE" if edge == "Rise" else "FALL"
            self.scope.write(f"MEASUREMENT:MEAS1:DELAY:{edge_label} {edge_dir}")
            level_cmd = f"{method}:{edge_dir}Mid {level}"
            self.scope.write(f"MEASUrement:{channel}:REFLevels:{level_cmd}")
        
        # SDA = EDGE1
        configure_edge(self.channels['SDA'], sda_mode, sda_edge, sda_edge_level, "EDGE1")

        # SCL = EDGE2
        configure_edge(self.channels['SCL'], scl_mode, scl_edge, scl_edge_level, "EDGE2")

        self.scope.write("MEASUrement:MEAS1:TOEDGESEARCHDIRect FORWARD")
        self.scope.query("*OPC?")

        try:
            results = {}
            results["SDA Setup Time"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            return True, "Setup Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Setup Time测量值失败: {str(e)}", {}
        
    def measure_hold_time(self, sda_mode, scl_mode, sda_edge, scl_edge, sda_edge_level, scl_edge_level):
        if not self.connected:
            return False, "未连接到示波器"
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
        # SDA = EDGE2
        configure_edge(self.channels['SDA'], sda_mode, sda_edge, sda_edge_level, "EDGE2")
        # SCL = EDGE1
        configure_edge(self.channels['SCL'], scl_mode, scl_edge, scl_edge_level, "EDGE1")

        self.scope.write("MEASUrement:MEAS1:TOEDGESEARCHDIRect FORWARD")
        self.scope.query("*OPC?")
        try:
            results = {}
            results["SDA Hold Time"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            return True, "Hold Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Hold Time测量值失败: {str(e)}", {}
        
    def measure_start_hold_time(self, sda_mode, scl_mode, sda_edge, scl_edge, sda_edge_level, scl_edge_level):
        if not self.connected:
            return False, "未连接到示波器"
        # 清除现有测量
        self.clear_measurements()
        self.scope.write("MEASUREMENT:ADDMEAS DELAY")
        self.scope.write("MEASUrement:GATing SCREEN")
        self.scope.write("MEASUrement:REFLevels:TYPE PerSource")
        self.scope.write(f"MEASUREMENT:MEAS1:SOURCE1 {self.channels['SDA']}")
        self.scope.write(f"MEASUREMENT:MEAS1:SOURCE2 {self.channels['SCL']}")

        def configure_edge(channel, mode, edge, level, edge_label):
            method = "PERCent" if mode == "Percentage" else "ABSolute"
            self.scope.write(f"MEASUrement:{channel}:REFLevels:METHod {method}")
            edge_dir = "RISE" if edge == "Rise" else "FALL"
            self.scope.write(f"MEASUREMENT:MEAS1:DELAY:{edge_label} {edge_dir}")
            level_cmd = f"{method}:{edge_dir}Mid {level}"
            self.scope.write(f"MEASUrement:{channel}:REFLevels:{level_cmd}")
        # SDA = EDGE1
        configure_edge(self.channels['SDA'], sda_mode, sda_edge, sda_edge_level, "EDGE1")
        # SCL = EDGE2
        configure_edge(self.channels['SCL'], scl_mode, scl_edge, scl_edge_level, "EDGE2")

        self.scope.write("MEASUrement:MEAS1:TOEDGESEARCHDIRect FORWARD")
        self.scope.query("*OPC?")
        try:
            results = {}
            results["Start Hold Time"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            return True, "Start Hold Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Start Hold Time测量值失败: {str(e)}", {}
        

    def measure_stop_setup_time(self, sda_mode, scl_mode, sda_edge, scl_edge, sda_edge_level, scl_edge_level):
        if not self.connected:
            return False, "未连接到示波器"
    
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

        self.scope.write("MEASUrement:MEAS1:TOEDGESEARCHDIRect FORWARD")
        self.scope.query("*OPC?")

        try:
            results = {}
            results["Stop Setup Time"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            return True, "Stop Setup Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Stop Setup Time测量值失败: {str(e)}", {}
        
    def measure_rise_fall(self, rise_high, rise_low, rise_mode, fall_high, fall_low, fall_mode):
        """添加电压测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        method="PERCent" if rise_mode=="Percentage" else "ABSolute"
        fall_method="PERCent" if fall_mode=="Percentage" else "ABSolute"
        # 清除所有现有测量
        self.clear_measurements()
        self.scope.write("MEASUrement:GATing SCREEN")
        self.scope.write("MEASUREMENT:ADDMEAS RISETIME")
        self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
        self.scope.write(f"MEASUREMENT:MEAS1:SOURCE {self.channels['SCL']}")
        self.scope.write(f"MEASUREMENT:MEAS1:REFLEVELS1:METHOD {method}")
        self.scope.write(f"MEASUrement:MEAS1:REFLevels1:{method}:RISELow {rise_low}")
        self.scope.write(f"MEASUrement:MEAS1:REFLevels1:{method}:RISEHigh {rise_high}")

        self.scope.write("MEASUREMENT:ADDMEAS FALLTIME")
        self.scope.write("MEASUrement:MEAS2:GLOBalref 0")
        self.scope.write(f"MEASUREMENT:MEAS2:SOURCE {self.channels['SCL']}")
        self.scope.write(f"MEASUREMENT:MEAS2:REFLEVELS1:METHOD {fall_method}")
        self.scope.write(f"MEASUREMENT:MEAS2:REFLEVELS1:{fall_method}:RISELow {fall_low}")
        self.scope.write(f"MEASUREMENT:MEAS2:REFLEVELS1:{fall_method}:RISEHigh {fall_high}")

        self.scope.write("MEASUREMENT:ADDMEAS RISETIME")
        self.scope.write("MEASUrement:MEAS3:GLOBalref 0")
        self.scope.write(f"MEASUREMENT:MEAS3:SOURCE {self.channels['SDA']}")
        self.scope.write(f"MEASUREMENT:MEAS3:REFLEVELS1:METHOD {method}")
        self.scope.write(f"MEASUrement:MEAS3:REFLevels1:{method}:RISELow {rise_low}")
        self.scope.write(f"MEASUrement:MEAS3:REFLevels1:{method}:RISEHigh {rise_high}")

        self.scope.write("MEASUREMENT:ADDMEAS FALLTIME")
        self.scope.write("MEASUrement:MEAS4:GLOBalref 0")
        self.scope.write(f"MEASUREMENT:MEAS4:SOURCE {self.channels['SDA']}")
        self.scope.write(f"MEASUREMENT:MEAS4:REFLEVELS1:METHOD {fall_method}")
        self.scope.write(f"MEASUREMENT:MEAS4:REFLEVELS1:{fall_method}:RISELow {fall_low}")
        self.scope.write(f"MEASUREMENT:MEAS4:REFLEVELS1:{fall_method}:RISEHigh {fall_high}")

        self.scope.query("*OPC?")
        try:
            results = {}
            results["SCL_Rise"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))))
            results["SCL_Fall"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))))
            results["SDA_Rise"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))))
            results["SDA_Fall"] = round_sig(adapt_time(float(self.scope.query("MEASUREMENT:MEAS4:VALUE?"))))
            
            return True, "Rise/Fall测量值获取成功", results
        except Exception as e:
            return False, f"获取Rise/Fall测量值失败: {str(e)}", {}

    def zoom(self, channel, edge_type, count, periods):
        if not self.connected:
            return False, "未连接到示波器"
        self.clear_measurements()
        #zoom out
        self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
        self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
        #measure wavelength of SCL if it has not been measured
        if self.wavelength is None:
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
            self.wavelength=1/frequency
            self.clear_measurements()
        #start search
        self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {self.wavelength/10*periods}")
        initial_position=0 if count>0 else 100
        direction="NEXT" if count>0 else "PREVious"
        #set the initial position
        self.scope.write(f"DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION {initial_position}")
        self.scope.write("SEARCH:ADDNew SEARCH1")
        self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
        self.scope.write(f"SEARCH:SEARCH1:TRIGGER:A:EDGE:SOURCE {channel}")
        self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe {edge_type}")
        self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
        self.scope.write("SEARCH:SEARCH1:STATE ON")
        self.scope.query("*OPC?")

        #move to specified edge
        for i in range(abs(count)):
            self.scope.write(f"SEARCH:SEARCH1:NAVigate {direction}")
            self.scope.query("*OPC?")
        #zoom 
        self.clear_measurements()
        return
    
    def zoom_composite(self, channel1, edge_type1, channel2, edge_type2, count1, count2, periods):
        if not self.connected:
            return False, "未连接到示波器"
        self.clear_measurements()
        #zoom out
        self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
        self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
        #measure wavelength of SCL if it has not been measured
        if self.wavelength is None:
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
            self.wavelength=1/frequency
            self.clear_measurements()
        self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {self.wavelength/10*periods}")
        #start search
        initial_position=0 if count1>0 else 100
        direction="NEXT" if count1>0 else "PREVious"
        #set the initial position
        self.scope.write(f"DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION {initial_position}")
        self.scope.write("SEARCH:ADDNew SEARCH1")
        self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
        self.scope.write(f"SEARCH:SEARCH1:TRIGGER:A:EDGE:SOURCE {channel1}")
        self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe {edge_type1}")
        self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
        self.scope.write("SEARCH:SEARCH1:STATE ON")
        self.scope.query("*OPC?")

        #move to specified edge
        for i in range(abs(count1)):
            self.scope.write(f"SEARCH:SEARCH1:NAVigate {direction}")
            self.scope.query("*OPC?")
        self.clear_measurements()

        direction="NEXT" if count2>0 else "PREVious"
        #here comes the composite part
        #search the other edge
        self.scope.write("SEARCH:ADDNew SEARCH1")
        self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
        self.scope.write(f"SEARCH:SEARCH1:TRIGGER:A:EDGE:SOURCE {channel2}")
        self.scope.write(f"SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe {edge_type2}")
        self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
        self.scope.write("SEARCH:SEARCH1:STATE ON")
        self.scope.query("*OPC?")
        #move to the specified edge
        for i in range(abs(count2)):
            self.scope.write(f"SEARCH:SEARCH1:NAVigate {direction}")
            self.scope.query("*OPC?")
        self.clear_measurements()
        return
    
    def save_screenshot(self, filename, filepath):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.scope.query("MEASUREMENT:MEAS1:VALUE?")
            self.scope.write(f'SAVe:IMAGe "C:\\Test\\{filename+".png"}"')
            self.scope.query("*OPC?")
            self.scope.write(f'FILESYSTEM:READFILE "C:\\Test\\{filename+".png"}"')
            raw_data=self.scope.read_raw()
            fid=open(filepath+f"\\{filename}.png", 'wb')
            fid.write(raw_data)
            fid.close()
            return True, f"截图已保存至 {filepath}\\{filename}"+".png"
        except Exception as e:
            error_msg = f"保存截图失败: {str(e)}"
            #self.log_message(error_msg)
            return False, error_msg
        
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
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {scale/2}")
            self.scope.write("SEARCH:DELETEALL")

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
    
class MSO64ControllerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MSO64示波器控制器")
        self.root.geometry("1500x1000")
        self.controller = MSO64Controller()
        self.create_widgets()
        self.res={}
        self.images={}#schema: name, file path
        self.controller.channels['SCL']=self.clock_var.get()
        self.controller.channels['SDA']=self.data_var.get()
        self.excel_path=None
        self.image_path=None

    def create_widgets(self):
        connection_frame = ttk.LabelFrame(self.root, text="连接设置")
        connection_frame.pack(fill="x", padx=10, pady=5)

        # IP地址输入
        ttk.Label(connection_frame, text="示波器IP地址:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.ip_entry = ttk.Entry(connection_frame, width=20)
        self.ip_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.ip_entry.insert(0, "169.254.103.178")  # 默认IP

        # 连接按钮
        self.connect_button = ttk.Button(connection_frame, text="连接", command=self.connect)
        self.connect_button.grid(row=0, column=2, padx=5, pady=5)

        # Keep these as instance attributes so they are accessible everywhere
        self.clock_options = ["CH1", "CH2", "CH3", "CH4", "REF1", "REF2"]
        self.data_options = ["CH2", "CH3", "CH4", "REF1", "REF2"]
        self.voltage_options=["1.8V", "3.3V"]
        
        self.clock_var = tk.StringVar(value=self.clock_options[4])
        self.data_var = tk.StringVar(value=self.data_options[4])
        self.voltage_var=tk.StringVar(value=self.voltage_options[1])
       
        # Set initial values for clock and data
        self.clock = self.clock_var.get()
        self.data = self.data_var.get()
        self.voltage=self.voltage_var.get()

        # Register trace callbacks to update clock/data values
        self.clock_var.trace_add("write", self.temp)
        self.data_var.trace_add("write", self.temp)
        #self.voltage_var.trace_add("write", self.on_voltage_change)

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

        self.run_all_tests_button = ttk.Button(i2c_frame, text="Write Cycle", command=self.write_cycle)
        self.run_all_tests_button.grid(row=0, column=0, padx=5, pady=5)

        ttk.Label(i2c_frame, text="Suffix Name:").grid(row=0, column=7, padx=5, pady=5, sticky="w")
        self.suffix_entry = ttk.Entry(i2c_frame, width=15)
        self.suffix_entry.grid(row=0, column=8, padx=5, pady=5, sticky="w")

        self.voltage_button = ttk.Button(i2c_frame, text="添加电压测量", command=self.write_voltage)
        self.voltage_button.grid(row=1, column=0, padx=5, pady=5)

        self.frequency_button = ttk.Button(i2c_frame, text="添加频率测量", command=self.write_frequency)
        self.frequency_button.grid(row=1, column=1, padx=5, pady=5)

        self.rise_fall_button = ttk.Button(i2c_frame, text="添加Rise/Fall测量", command=self.write_rise_fall)
        self.rise_fall_button.grid(row=1, column=2, padx=5, pady=5)

        self.delay_button = ttk.Button(i2c_frame, text="添加Setup Time测量", command=self.write_setup)
        self.delay_button.grid(row=1, column=3, padx=5, pady=5)

        self.hold_time_button = ttk.Button(i2c_frame, text="添加Hold Time测量", command=self.write_hold)
        self.hold_time_button.grid(row=1, column=4, padx=5, pady=5)

        self.start_hold_time_button = ttk.Button(i2c_frame, text="添加Start Hold Time测量", command=self.write_start_hold)
        self.start_hold_time_button.grid(row=1, column=5, padx=5, pady=5)

        self.stop_setup_time_button = ttk.Button(i2c_frame, text="添加Stop Setup Time测量", command=self.write_stop_setup)
        self.stop_setup_time_button.grid(row=1, column=6, padx=5, pady=5)

        self.run_read_tests_button = ttk.Button(i2c_frame, text="Read Cycle", command=self.temp)
        self.run_read_tests_button.grid(row=2, column=0, padx=5, pady=5)

        self.voltage_button_read = ttk.Button(i2c_frame, text="添加电压测量", command=self.temp)
        self.voltage_button_read.grid(row=3, column=0, padx=5, pady=5)

        self.rise_fall_button_read = ttk.Button(i2c_frame, text="添加Rise Fall Time测量", command=self.temp)
        self.rise_fall_button_read.grid(row=3, column=1, padx=5, pady=5)

        self.setup_button_read = ttk.Button(i2c_frame, text="添加Setup Time测量", command=self.temp)
        self.setup_button_read.grid(row=3, column=2, padx=5, pady=5)

        self.hold_button_read = ttk.Button(i2c_frame, text="添加Hold Time测量", command=self.temp)
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
        self.setup_time_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.setup_time_tab, text="Setup Time参数")

        self.hold_time_tab=ttk.Frame(self.tab_control)
        self.tab_control.add(self.hold_time_tab, text="Hold Time参数")
        """
        self.start_hold_time_tab=ttk.Frame(self.tab_control)
        self.tab_control.add(self.start_hold_time_tab, text="Start Hold Time参数")

        self.stop_setup_time_tab=ttk.Frame(self.tab_control)
        self.tab_control.add(self.stop_setup_time_tab, text="Stop Setup Time参数")
        """
        # 运行测试按钮
    
        self.run_test_button = ttk.Button(test_frame, text="Save to Excel", command=self.save_data_to_excel)
        self.run_test_button.grid(row=3, column=0, padx=5, pady=5)

        # 创建电压测量参数表
        self.create_voltage_tab()
        
        # 创建频率测量参数表
        self.create_frequency_tab()
        
        # 创建Setup Time测量参数表
        #self.create_start_hold_time_tab()
        self.create_setup_time_tab()

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

    def browse_excel_path(self):
        path = filedialog.askdirectory()
        if path:
            self.excel_path_entry.delete(0, tk.END)
            base_dir=path.replace('/', '\\')
            # Define folder structure
            test_data_path = os.path.join(base_dir, "test data")
            i2c_path = os.path.join(test_data_path, "I2C")
            waveform_path = os.path.join(i2c_path, "excel")

            # Create folders if they don't exist
            os.makedirs(waveform_path, exist_ok=True)
            self.excel_path = waveform_path
            self.excel_path_entry.insert(0, self.excel_path)
            
            self.log_message(f"已选择图片保存路径: {self.excel_path}")

    def browse_image_path(self):
        path = filedialog.askdirectory()
        if path:
            self.image_path_entry.delete(0, tk.END)
            base_dir=path.replace('/', '\\')
            # Define folder structure
            test_data_path = os.path.join(base_dir, "test data")
            i2c_path = os.path.join(test_data_path, "I2C")
            waveform_path = os.path.join(i2c_path, "waveform")

            # Create folders if they don't exist
            os.makedirs(waveform_path, exist_ok=True)
            self.image_path = waveform_path
            self.image_path_entry.insert(0, self.image_path)
            
            self.log_message(f"已选择图片保存路径: {self.image_path}")

    def log_message(self, message):
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)

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
    
    def connect(self):
        ip_address = self.ip_entry.get().strip()
        if not ip_address:
            messagebox.showerror("错误", "请输入示波器IP地址")
            return
        if self.controller.connected==False:
            self.log_message(f"正在连接到 {ip_address}...")

        # 如果已连接，先断开
        if self.controller.connected:
            self.controller.disconnect()
            self.status_label.config(text="未连接")
            self.connect_button.config(text="连接")
            #self.update_button_states(False)
            self.log_message("已断开连接")
            return
        
        # 连接到示波器
        success, message = self.controller.connect(ip_address)
        if success:
            self.status_label.config(text="已连接")
            self.connect_button.config(text="断开")
            #self.update_button_states(True)
            self.log_message(f"连接成功: {message}")
        else:
            messagebox.showerror("连接错误", message)
            self.log_message(f"连接失败: {message}")

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
        
        self.get_voltage_scl_button = ttk.Button(self.voltage_tab, text="获取SCL电压测量值", command=self.temp)
        self.get_voltage_scl_button.grid(row=5, column=0, padx=5, pady=10)

        self.get_voltage_sda_button = ttk.Button(self.voltage_tab, text="获取SDA电压测量值", command=self.temp)
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
            command=self.temp
        )
        self.get_frequency_button.grid(row=8, column=0, padx=5, pady=(10, 10))

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
        self.get_rise_button = ttk.Button(self.rise_fall_tab, text="获取Rise测量值", command=self.temp)
        self.get_rise_button.grid(row=6, column=0, padx=5, pady=10)

        # 获取电压测量值按钮
        self.get_fall_button = ttk.Button(self.rise_fall_tab, text="获取Fall测量值", command=self.temp)
        self.get_fall_button.grid(row=6, column=2, padx=5, pady=10)

        # Screenshot
        #self.get_rise_fall_screenshot_button = ttk.Button(self.rise_fall_tab, text="Save Screenshot", command=lambda: self.save_screenshot("Rise Fall"))
        #self.get_rise_fall_screenshot_button.grid(row=6, column=0, columnspan=4, padx=10, pady=10)

    def create_setup_time_tab(self):
        """创建setup time测量参数表"""
        ttk.Label(self.setup_time_tab, text="Setup Time参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")

        # Mode Dropdown
        ttk.Label(self.setup_time_tab, text="SDA Mode: ").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sda_mode = tk.StringVar()
        self.sda_mode_combobox = ttk.Combobox(self.setup_time_tab, textvariable=self.sda_mode, values=["Percentage", "Absolute"], state="readonly", width=10)
        self.sda_mode_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.sda_mode_combobox.current(0)  # default to "Percentage"

        ttk.Label(self.setup_time_tab, text="SCL Mode: ").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.scl_mode = tk.StringVar()
        self.scl_mode_combobox = ttk.Combobox(self.setup_time_tab, textvariable=self.scl_mode, values=["Percentage", "Absolute"], state="readonly", width=10)
        self.scl_mode_combobox.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        self.scl_mode_combobox.current(0)  # default to "Percentage"

        ttk.Label(self.setup_time_tab, text="SDA Edge: ").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.sda_edge_var = tk.StringVar()
        self.sda_edge_combobox = ttk.Combobox(self.setup_time_tab, textvariable=self.sda_edge_var, values=["Rise", "Fall"], state="readonly", width=8)
        self.sda_edge_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.sda_edge_combobox.current(0)  

        ttk.Label(self.setup_time_tab, text="SCL Edge: ").grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.scl_edge_var = tk.StringVar()
        self.scl_edge_combobox = ttk.Combobox(self.setup_time_tab, textvariable=self.scl_edge_var, values=["Rise", "Fall"], state="readonly", width=8)
        self.scl_edge_combobox.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        self.scl_edge_combobox.current(0)  # default to "rise"

        ttk.Label(self.setup_time_tab, text="SDA: Edge Mid").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.setup_ch2_entry = ttk.Entry(self.setup_time_tab, width=10)
        self.setup_ch2_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.setup_ch2_entry.insert(0, "70")  # 默认
        
        ttk.Label(self.setup_time_tab, text="SCL: Edge Mid").grid(row=3, column=2, padx=5, pady=5, sticky="w")
        self.setup_ch1_entry = ttk.Entry(self.setup_time_tab, width=10)
        self.setup_ch1_entry.grid(row=3, column=3, padx=5, pady=5, sticky="w")
        self.setup_ch1_entry.insert(0, "30")  # 默认
        
        ttk.Label(self.setup_time_tab, text="Setup Time:").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.delay_value = ttk.Label(self.setup_time_tab, text="--")
        self.delay_value.grid(row=4, column=2, padx=5, pady=2, sticky="w")

        # 获取setup time测量值按钮
        self.get_delay_button = ttk.Button(self.setup_time_tab, text="获取Setup Time测量值", command=self.temp)
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
        self.get_hold_time_button = ttk.Button(self.hold_time_tab, text="获取Hold Time测量值", command=self.temp)
        self.get_hold_time_button.grid(row=5, column=0, columnspan=1, padx=5, pady=5)

    def temp(self):
        return
    
    def write_voltage(self, display=True):
        if self.image_path==None:
            messagebox.showerror("错误", "Image file path has not been configured")
            return
        if self.folder_entry.get()=="":
            messagebox.showerror("错误", "Folder name has not been configured")
            return
        self.log_message("正在获取电压测量值...")
        self.controller.zoom_composite(self.data, "FALL", self.clock, "FALL", 1, 5, 8)
        self.controller.scope.write("*WAI")
        suc, msg, res=self.controller.measure_voltage()
        self.res.update(res)
        if suc:
            # 更新电压测量值显示
            self.update_voltage()
            if display:
                messagebox.showinfo("成功", "电压测量值获取成功")
            self.log_message(msg)
        else:
            messagebox.showerror("错误", msg)
        self.controller.scope.write("*WAI")
        self.save_screenshot("Voltage", "v")

    def write_frequency(self, display=True):
        #self.insert_frequency_parameters()
        if self.image_path==None:
            messagebox.showerror("错误", "Image file path has not been configured")
            return
        if self.folder_entry.get()=="":
            messagebox.showerror("错误", "Folder name has not been configured")
            return
        self.log_message("正在获取频率测量值...")
        self.controller.zoom_composite(self.data, "FALL", self.clock, "FALL", 1, 5, 8)
        self.controller.scope.query("*OPC?")
        suc, msg, res=self.controller.measure_frequency(self.wavelength_entry.get(), self.pwd_entry.get(), self.nwd_entry.get(), self.frequency_mode.get())
        self.res.update(res)
        if suc:
            # 更新电压测量值显示
            self.update_frequency()
            if display:
                messagebox.showinfo("成功", "频率测量值获取成功")
            self.log_message(msg)
        else:
            messagebox.showerror("错误", msg)
        self.controller.scope.write("*WAI")
        self.save_screenshot("Frequency", "freq")

    def write_rise_fall(self, display=True):
        if self.image_path==None:
            messagebox.showerror("错误", "Image file path has not been configured")
            return
        if self.folder_entry.get()=="":
            messagebox.showerror("错误", "Folder name has not been configured")
            return
        self.log_message("正在获取Rise Fall测量值...")
        self.controller.zoom_composite(self.data, "RISE", self.clock, "RISE", 2, 2, 3)
        self.controller.scope.write("*WAI")
        suc, msg, res=self.controller.measure_rise_fall(self.rising_edge_high_entry.get(), self.rising_edge_low_entry.get(), self.rise_mode.get(), 
                                                        self.falling_edge_high_entry.get(), self.falling_edge_low_entry.get(), self.fall_mode.get())
        self.res.update(res)
        if suc:
            # 更新电压测量值显示
            self.update_rise_fall()
            if display:
                messagebox.showinfo("成功", "Rise Fall测量值获取成功")
            self.log_message(msg)
        else:
            messagebox.showerror("错误", msg)
        self.controller.scope.write("*WAI")
        self.save_screenshot("Rise Fall Time", "rf")

    def write_setup(self, display=True):
        if self.image_path==None:
            messagebox.showerror("错误", "Image file path has not been configured")
            return
        if self.folder_entry.get()=="":
            messagebox.showerror("错误", "Folder name has not been configured")
            return
        self.log_message("正在获取Setup Time测量值...")
        self.controller.zoom_composite(self.data, "RISE", self.clock, "RISE", 2, 1, 1.00)
        self.controller.scope.write("*WAI")
        suc, msg, res=self.controller.measure_setup_time(self.sda_mode.get(), self.scl_mode.get(), self.sda_edge_var.get(), self.scl_edge_var.get(),
                                                         self.setup_ch2_entry.get(), self.setup_ch1_entry.get())
        self.res.update(res)
        if suc:
            self.delay_value.config(text=f"{self.res['SDA Setup Time']} us")
            self.tab_control.select(self.setup_time_tab)
            if display:
                messagebox.showinfo("成功", "Setup Time测量值获取成功")
            self.log_message(msg)
        else:
            messagebox.showerror("错误", msg)
        self.controller.scope.write("*WAI")
        self.save_screenshot("Setup Time", "tsu")

    def write_hold(self, display=True):
        if self.image_path==None:
            messagebox.showerror("错误", "Image file path has not been configured")
            return
        if self.folder_entry.get()=="":
            messagebox.showerror("错误", "Folder name has not been configured")
            return
        self.log_message("正在获取Hold Time测量值...")
        self.controller.zoom_composite(self.data, "FALL", self.clock, "FALL", 3, -1, 1)
        self.controller.scope.write("*WAI")
        suc, msg, res=self.controller.measure_hold_time(self.sda_mode_hold.get(), self.scl_mode_hold.get(), self.sda_edge_var_hold.get(), 
                                                        self.scl_edge_var_hold.get(), self.hold_ch2_entry.get(), self.hold_ch1_entry.get())
        self.res.update(res)
        if suc:
            self.delay_value.config(text=f"{self.res['SDA Hold Time']} us")
            self.tab_control.select(self.hold_time_tab)
            if display:
                messagebox.showinfo("成功", "Hold Time测量值获取成功")
            self.log_message(msg)
        else:
            messagebox.showerror("错误", msg)
        self.controller.scope.write("*WAI")
        self.save_screenshot("Hold Time", "thd")

    def write_start_hold(self, display=True):
        if self.image_path==None:
            messagebox.showerror("错误", "Image file path has not been configured")
            return
        if self.folder_entry.get()=="":
            messagebox.showerror("错误", "Folder name has not been configured")
            return
        self.log_message("正在获取Start Hold Time测量值...")
        self.controller.zoom_composite(self.data, "FALL", self.clock, "FALL", 1, 0, 1)
        self.controller.scope.write("*WAI")
        suc, msg, res=self.controller.measure_start_hold_time(self.sda_mode_hold.get(), self.scl_mode_hold.get(), self.sda_edge_var_hold.get(), 
                                                        self.scl_edge_var_hold.get(), self.hold_ch1_entry.get(), self.hold_ch2_entry.get())
        self.res.update(res)
        if suc:
            if display:
                messagebox.showinfo("成功", "Start Hold Time测量值获取成功")
            self.log_message(msg)
        else:
            messagebox.showerror("错误", msg)
        self.controller.scope.write("*WAI")
        self.save_screenshot("Start Hold Time", "sta")

    def write_stop_setup(self, display=True):
        if self.image_path==None:
            messagebox.showerror("错误", "Image file path has not been configured")
            return
        if self.folder_entry.get()=="":
            messagebox.showerror("错误", "Folder name has not been configured")
            return
        self.log_message("正在获取Stop Setup Time测量值...")
        self.controller.zoom_composite(self.data, "RISE", self.clock, "RISE", -1, 0, 1)
        self.controller.scope.write("*WAI")
        suc, msg, res=self.controller.measure_stop_setup_time(self.sda_mode.get(), self.scl_mode.get(), self.sda_edge_var.get(), 
                                                              self.scl_edge_var.get(), self.setup_ch1_entry.get(), self.setup_ch2_entry.get())
        self.res.update(res)
        if suc:
            if display:
                messagebox.showinfo("成功", "Start Hold Time测量值获取成功")
            self.log_message(msg)
        else:
            messagebox.showerror("错误", msg)
        self.controller.scope.write("*WAI")
        self.save_screenshot("Stop Setup Time", "sto")

    def write_cycle(self):
        self.write_voltage(False)
        self.write_frequency(False)
        self.write_rise_fall(False)
        self.write_setup(False)
        self.write_hold(False)
        self.write_start_hold(False)
        self.write_stop_setup(False)
        self.controller.zoom_on_rising_edge()
        self.controller.scope.write("*WAI")
        time.sleep(0.5)
        self.save_screenshot("Rise Time", "r")
        self.controller.zoom_on_falling_edge()
        self.save_screenshot("Fall Time", "f")
        self.save_data_to_excel()
        

    def save_screenshot(self, name, filename):
        if self.folder_entry==None:
            messagebox.showerror("错误", "Folder name has not been configured")
            return
        foldername=self.folder_entry.get()
        path=os.path.join(self.image_path, foldername)
        os.makedirs(path, exist_ok=True)
        time.sleep(0.5)
        suc, msg=self.controller.save_screenshot(filename, path)
        if suc:
            self.images[name]=self.image_path+f"\\{self.folder_entry.get()}\\{filename}.png"
        self.log_message(msg)

    def update_voltage(self):
        self.scl_max_value.config(text=f"{self.res['SCL_Maximum']} V")
        self.scl_min_value.config(text=f"{self.res['SCL_Minimum']} V")
        self.scl_top_value.config(text=f"{self.res['SCL_Top']} V")
        self.scl_base_value.config(text=f"{self.res['SCL_Base']} V")
        
        self.sda_max_value.config(text=f"{self.res['SDA_Maximum']} V")
        self.sda_min_value.config(text=f"{self.res['SDA_Minimum']} V")
        self.sda_top_value.config(text=f"{self.res['SDA_Top']} V")
        self.sda_base_value.config(text=f"{self.res['SDA_Base']} V")
        self.tab_control.select(self.voltage_tab)
    
    def update_frequency(self):
        self.frequency_value.config(text=f"{self.res['SCL Frequency']} KHz")
        self.pos_width_value.config(text=f"{self.res['SCL Positive_Pulse_Width']} us")
        self.neg_width_value.config(text=f"{self.res['SCL Negative_Pulse_Width']} us")
        self.tab_control.select(self.frequency_tab)

    def update_rise_fall(self):
        self.ch1_risetime.config(text=f"{self.res['SCL_Rise']} ns")
        self.ch1_falltime.config(text=f"{self.res['SCL_Fall']} ns")
        self.ch2_risetime.config(text=f"{self.res['SDA_Rise']} ns")
        self.ch2_falltime.config(text=f"{self.res['SDA_Fall']} ns")
        self.tab_control.select(self.rise_fall_tab)

    def insert_frequency_parameters(self):
        voltage=3.3 if self.voltage_var.get()=="3.3V" else 1.8
        self.frequency_mode_combobox.current(1)
        self.wavelength_entry.insert(0, str(round_sig(voltage*0.3)))
        self.pwd_entry.insert(0, str(round_sig(voltage*0.7)))
        self.nwd_entry.insert(0, str(round_sig(voltage*0.3)))
    
    def save_data_to_excel(self):
        """保存示波器截图到C:\ST\hsd\i2c目录"""
        if self.excel_path is None:
            messagebox.showerror("错误", "Excel file path has not been configured")
            return
        df = pd.DataFrame(list(self.res.items()), columns=['Test Item', 'Measured Data'])
        try:
            with pd.ExcelWriter(self.excel_path+f"\\{self.folder_entry.get()}.xlsx", engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)
                workbook  = writer.book
                worksheet = writer.sheets['Sheet1']
                # Autofit columns based on content
                for i, col in enumerate(df.columns):
                    #column_len = max(df[col].astype(str).map(len).max(), len(col))
                    #print(column_len+2)
                    worksheet.set_column(i, i, 24)  # Add a little extra space
                # Define scaling factors
                x_scale = 0.33
                y_scale = 0.33
                Col='A'
                index=0
                row=len(df)+3
                text_col='B'
                for key in list(self.images.keys()):
                    if index%2==0:
                        Col='A'
                        text_col='B'
                    else:
                        Col='F'
                        text_col='I'
                    row=16*math.floor(index//2)+len(df)+5
                    image_path = self.images[key]
                    img = Image.open(image_path)
                    width, height = img.size
                # Insert image at 'C2' with scaling
                    worksheet.insert_image(f'{Col}{str(row)}', image_path, {'x_scale': x_scale, 'y_scale': y_scale})
                    cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                    worksheet.write(f'{text_col}{str(row-1)}', key, cell_format)
                    index+=1
                self.res={}
                self.images={}
                self.controller.wavelength=None
                self.log_message("Excel saved to "+f"{self.excel_path}"+f"\\{self.folder_entry.get()}.xlsx")
                self.folder_entry.delete(0, tk.END)
            return True
        except Exception as e:
            error_msg = f"保存Excel失败: {str(e)}"
            self.log_message(error_msg)
            return False
        
if __name__ == "__main__":
    root = tk.Tk()
    app = MSO64ControllerGUI(root)
    root.mainloop() 

"""
controller=MSO64Controller()
controller.connect("169.254.103.178")
controller.channels['SCL']="REF1"
controller.channels['SDA']="REF2"
controller.zoom_composite("REF2", "RISE", "REF1", "RISE", 2, 1, 1.00)
controller.measure_setup_time("Percentage", "Percentage", "Rise", "Rise", 30, 70)
controller.save_screenshot("test", "C:\\Users\\eason\\Documents\\MSO64B")
"""