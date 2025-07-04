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

    def connect(self, ip_address):
        """连接到示波器"""
        try:
            resource_string = f"TCPIP::{ip_address}::INSTR"
            self.scope = self.rm.open_resource(resource_string)
            self.scope.timeout = 5000  # 设置超时时间为10秒
            idn = self.scope.query("*IDN?")
            self.scope.write('FILESystem:MKDir "C:\\Test"')
            self.scope.write(" DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:SCALe 1")
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
            print(str(e))
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
            results["CH1_Maximum"] = float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            results["CH1_Minimum"] = float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))
            results["CH1_Top"] = float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))
            results["CH1_Base"] = float(self.scope.query("MEASUREMENT:MEAS4:VALUE?"))
            
            # CH2电压参数
            results["CH2_Maximum"] = float(self.scope.query("MEASUREMENT:MEAS5:VALUE?"))
            results["CH2_Minimum"] = float(self.scope.query("MEASUREMENT:MEAS6:VALUE?"))
            results["CH2_Top"] = float(self.scope.query("MEASUREMENT:MEAS7:VALUE?"))
            results["CH2_Base"] = float(self.scope.query("MEASUREMENT:MEAS8:VALUE?"))
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
            results["Frequency"] = float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            results["Positive_Pulse_Width"] = float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))
            results["Negative_Pulse_Width"] = float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))
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
            self.scope.write("MEASUrement:MEAS1:GATing SCREEN")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE1 CH2")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE2 CH1")
            
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:REFLevels:TYPE PerSource")
            #self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            
            # 设置CH2(SOURCE1)的参考电平为70%
            self.scope.write("MEASUrement:CH1:REFLevels:PERCent:RISEHigh 90")
            self.scope.write("MEASUrement:CH1:REFLevels:PERCent:RISEMid 30")
            self.scope.write("MEASUrement:CH1:REFLevels:PERCent:RISELow 10")
            # 设置CH1(SOURCE2)的参考电平为30%
            self.scope.write("MEASUrement:CH2:REFLevels:PERCent:RISEHigh 90")
            self.scope.write("MEASUrement:CH2:REFLevels:PERCent:RISEMid 70")
            self.scope.write("MEASUrement:CH2:REFLevels:PERCent:RISELow 10")
            # 设置测量的边沿和方向
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION1 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION2 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE1 RISE")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE2 RISE")
            
            
            #self.scope.write(f'SAVe:EVENTtable:MEASUrement "C:\\Test\\setup_time{self.time}.txt"')
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
            # 添加电压参数
            self.scope.write("MEASUREMENT:ADDMEAS RISETIME")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE CH1")
            
            self.scope.write("MEASUREMENT:ADDMEAS FALLTIME")
            self.scope.write("MEASUREMENT:MEAS2:SOURCE CH1")
            
            self.scope.write("MEASUREMENT:ADDMEAS RISETIME")
            self.scope.write("MEASUREMENT:MEAS3:SOURCE CH2")
            
            self.scope.write("MEASUREMENT:ADDMEAS FALLTIME")
            self.scope.write("MEASUREMENT:MEAS4:SOURCE CH2")

            self.scope.write("MEASUrement:REFLevels:PERCent:RISEHigh 70")
            self.scope.write("MEASUrement:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUrement:REFLevels:PERCent:RISELow 30")

            self.scope.write("MEASUrement:REFLevels:PERCent:FALLHigh 70")
            self.scope.write("MEASUrement:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUrement:REFLevels:PERCent:FALLLow 30")
            self.scope.query("*OPC?")

            return True, "电压测量项添加成功"
        except Exception as e:
            print(str(e))
            return False, f"添加电压测量项失败: {str(e)}"
        
    def get_rise_fall_measurements(self):
        """获取电压测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            results["CH1_Rise"] = float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            results["CH1_Fall"] = float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))
            results["CH2_Rise"] = float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))
            results["CH2_Fall"] = float(self.scope.query("MEASUREMENT:MEAS4:VALUE?"))
            
            return True, "Rise/Fall测量值获取成功", results
        except Exception as e:
            return False, f"获取Rise/Fall测量值失败: {str(e)}", {}
    
    def add_hold_time_measurements(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            self.scope.write("MEASUREMENT:ADDMEAS DELAY")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE1 CH1")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE2 CH2")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUrement:MEAS1:GATing SCREEN")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:REFLevels:TYPE PerSource")
            
            # 设置CH2(SOURCE1)的参考电平为70%
            self.scope.write("MEASUrement:CH1:REFLevels:PERCent:FALLMid 30")
            # 设置CH1(SOURCE2)的参考电平为30%
            self.scope.write("MEASUrement:CH2:REFLevels:PERCent:FALLMid 70")
            # 设置测量的边沿和方向
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION1 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION2 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE1 FALL")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE2 FALL")
            self.scope.write("OPC")

            return True, "Hold Time测量项添加成功"
        except Exception as e:
            print(str(e))
            return False, f"添加Hold Time测量项失败: {str(e)}"
        
    def get_delay_measurement(self):
        """获取Hold Time测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            results["Delay_Rise_Edge_70_30"] = float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            return True, "Setup Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Setup Time测量值失败: {str(e)}", {}
        
    def get_hold_measurement(self):
        """获取Setup Time测量值"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            results = {}
            results["Delay_Fall_Edge_30_70"] = float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
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
            self.scope.write("MEASUREMENT:MEAS1:SOURCE1 CH1")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE2 CH2")
            self.scope.write("MEASUrement:MEAS1:GATing SCREEN")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            #self.scope.write("MEASUrement:REFLevels:TYPE PerSource")
            #self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            
            # 设置CH2(SOURCE1)的参考电平为70%
            self.scope.write("MEASUrement:CH1:REFLevels:PERCent:RISEHigh 90")
            self.scope.write("MEASUrement:CH1:REFLevels:PERCent:RISEMid 70")
            self.scope.write("MEASUrement:CH1:REFLevels:PERCent:RISELow 10")
            # 设置CH1(SOURCE2)的参考电平为30%
            self.scope.write("MEASUrement:CH2:REFLevels:PERCent:RISEHigh 90")
            self.scope.write("MEASUrement:CH2:REFLevels:PERCent:RISEMid 30")
            self.scope.write("MEASUrement:CH2:REFLevels:PERCent:RISELow 10")
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
            results["Stop Time"] = float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
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
            """
            # 固定保存路径为C:\ST\hsd\i2c
            file_path = f"C:\\ST\\hsd\\i2c\\{filename}"
            
            # 设置图片格式为PNG
            self.scope.write("SAVE:IMAGE:FILEFORMAT PNG")
            #self.log_message(f"设置图片格式为PNG")
            
            # 发送SAVE:IMAGE命令，使用两个单引号包裹路径
            cmd = f"SAVE:IMAGE ''{file_path}''"
            self.scope.write(cmd)
            self.log_message(f"发送命令: {cmd}")
            """
            # 等待保存完成
            #time.sleep(2)
            
            # 检查命令是否执行成功
            return True, f"截图已保存至 {self.image_path}"+".png"
        except Exception as e:
            error_msg = f"保存截图失败: {str(e)}"
            #self.log_message(error_msg)
            return False, error_msg
        
    def get_measurements(self):
        """获取所有测量值（旧版本，保留用于兼容）"""
        if not self.connected:
            return False, "未连接到示波器", {}
        
        try:
            # 添加所有测量项
            success, message = self.add_voltage_measurements()
            if not success:
                return False, message, {}
            
            success, message, voltage_results = self.get_voltage_measurements()
            if not success:
                return False, message, {}
            
            success, message = self.add_frequency_measurements()
            if not success:
                return False, message, {}
            
            success, message, frequency_results = self.get_frequency_measurements()
            if not success:
                return False, message, {}
            
            success, message = self.add_delay_measurement()
            if not success:
                return False, message, {}
            
            success, message, delay_results = self.get_delay_measurement()
            if not success:
                return False, message, {}
            
            # 合并所有结果
            results = {}
            results.update(voltage_results)
            results.update(frequency_results)
            results.update(delay_results)
            
            return True, "所有测量值获取成功", results
        except Exception as e:
            return False, f"获取测量值失败: {str(e)}", {}
        
    def save_data_to_excel(self, measurements, filename, foldername, images):
        """保存示波器截图到C:\ST\hsd\i2c目录"""
        if not self.connected:
            return False, "未连接到示波器"
        df = pd.DataFrame(list(measurements.items()), columns=['Test Item', 'Measured Data'])
        try:
            full_path = os.path.join(self.excel_path, foldername)
            if os.path.isdir(full_path):
                #df.to_excel(self.excel_path+f"\\{filename}.xlsx", index=False)
                with pd.ExcelWriter(self.excel_path+f"\\{foldername}\\{filename}.xlsx", engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False)

                    workbook  = writer.book
                    worksheet = writer.sheets['Sheet1']

                    # Autofit columns based on content
                    for i, col in enumerate(df.columns):
                        column_len = max(df[col].astype(str).map(len).max(), len(col))
                        worksheet.set_column(i, i, column_len + 2)  # Add a little extra space
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
                        row=16*math.floor(index//2)+len(df)+3
                        image_path = images[key]
                        img = Image.open(image_path)
                        width, height = img.size
                    # Insert image at 'C2' with scaling
                        worksheet.insert_image(f'{Col}{str(row)}', image_path, {'x_scale': x_scale, 'y_scale': y_scale})
                        worksheet.write(f'{text_col}{str(row-1)}', key)
                        index+=1
            else:
                os.chdir(self.excel_path)
                os.mkdir(foldername)
                #df.to_excel(self.excel_path+f"\\{foldername}\\{filename}.xlsx", index=False)
                with pd.ExcelWriter(self.excel_path+f"\\{foldername}\\{filename}.xlsx", engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False)

                    workbook  = writer.book
                    worksheet = writer.sheets['Sheet1']

                    # Autofit columns based on content
                    for i, col in enumerate(df.columns):
                        column_len = max(df[col].astype(str).map(len).max(), len(col))
                        worksheet.set_column(i, i, column_len + 2)  # Add a little extra space
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
                        row=16*math.floor(index//2)+len(df)+3
                        image_path = images[key]
                        img = Image.open(image_path)
                        width, height = img.size
                    # Insert image at 'C2' with scaling
                        worksheet.insert_image(f'{Col}{str(row)}', image_path, {'x_scale': x_scale, 'y_scale': y_scale})
                        worksheet.write(f'{text_col}{str(row-1)}', key)
                        index+=1
                 
            return True, "Excel saved to "+f"\\{foldername}\\{filename}.xlsx"
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
            self.scope.write("MEASUREMENT:MEAS1:SOURCE1 CH2")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE2 CH1")
            self.scope.write("MEASUrement:MEAS1:GATing SCREEN")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
            self.scope.write("MEASUrement:REFLevels:TYPE PerSource")
            #self.scope.write("MEASUrement:MEAS1:GLOBalref 0")
            
            # 设置CH2(SOURCE1)的参考电平为70%
            self.scope.write("MEASUrement:CH1:REFLevels:PERCent:FALLHigh 90")
            self.scope.write("MEASUrement:CH1:REFLevels:PERCent:FALLMid 70")
            self.scope.write("MEASUrement:CH1:REFLevels:PERCent:FALLLow 10")
            # 设置CH1(SOURCE2)的参考电平为30%
            self.scope.write("MEASUrement:CH2:REFLevels:PERCent:FALLHigh 90")
            self.scope.write("MEASUrement:CH2:REFLevels:PERCent:FALLMid 30")
            self.scope.write("MEASUrement:CH2:REFLevels:PERCent:FALLLow 10")
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
            results["Start Hold Time"] = float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            return True, "Start Hold Time测量值获取成功", results
        except Exception as e:
            return False, f"获取Start Hold Time测量值失败: {str(e)}", {}
        
    def zoom_on_rising_edge(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # Optionally, use search to locate the rising edge # Clear existing searches
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:SOUrce CH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            # Wait for acquisition and search   
            self.scope.query("*OPC?")  # Ensure acquisition completes
            self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
            self.scope.write("MEASUrement:REFLEVELS:TYPE GLOBAL")
            # 添加电压参数
            self.scope.write("MEASUREMENT:ADDMEAS RISETIME")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE CH1")
            self.scope.write("MEASUrement:MEAS1:GATing NONE")
            self.scope.write("MEASUrement:REFLevels:PERCent:RISEHigh 70")
            self.scope.write("MEASUrement:REFLevels:PERCent:RISEMid 50")
            self.scope.write("MEASUrement:REFLevels:PERCent:RISELow 30")
            self.scope.query("*OPC?")
            scale=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {scale/2}")
            #current

            """
            # Enable zoom and center on the edge
            self.scope.write("HORizontal:ZOOM:STATE ON")
            self.scope.write("HORizontal:ZOOM:SCALe 10E-9")  # 10 ns/div (adjust for edge duration, e.g., 100 ns wide)
            self.scope.write(f"HORizontal:ZOOM:POSition {edge_time}")  # Center zoom on edge
            """
            return True, "success"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def zoom_on_falling_edge(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:

            self.clear_measurements()
            # Optionally, use search to locate the rising edge
            self.scope.write("SEARCH:CLEAr")  # Clear existing searches
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:SOUrce CH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")
            # Wait for acquisition and search   
            self.scope.query("*OPC?")  # Ensure acquisition completes
            self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
            #self.scope.write("MEASUrement:REFLEVELS:TYPE GLOBAL")
            # 添加电压参数
            self.scope.write("MEASUREMENT:ADDMEAS FALLTIME")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE CH1")
            self.scope.write("MEASUrement:MEAS1:GATing NONE")
            self.scope.write("MEASUrement:REFLevels:PERCent:FALLHigh 70")
            self.scope.write("MEASUrement:REFLevels:PERCent:FALLMid 50")
            self.scope.write("MEASUrement:REFLevels:PERCent:FALLLow 30")
            self.scope.query("*OPC?")
            scale=float(self.scope.query("MEASUREMENT:MEAS1:VALUE?"))
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {scale/4}")
            

            return True, "success"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"

    def zoom_middle(self):
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # Optionally, use search to locate the rising edge
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:SOUrce CH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")

            n=self.scope.query("SEARCH:SEARCH1:TOTAL?")
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 0")

            for i in range(4):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")

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
            self.scope.query("*OPC?")
            if self.upper is None:
                self.upper=float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))
            if self.lower is None:
                self.lower=float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))
            self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {(self.upper+self.lower)/1.25}")
            self.clear_measurements()
            self.scope.write("SEARCH:DELETEALL")
            return True, "success"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
        
    def zoom_hold_time(self):
        #find 2nd ch2 falling edge
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # Optionally, use search to locate the rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce CH2")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            n=self.scope.query("SEARCH:SEARCH1:TOTAL?")
            #print(n)
            
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION -2000")
            for i in range(2):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")

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
            self.scope.query("*OPC?")

            if self.upper is None:
                self.upper=float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))
                self.scope.query("*OPC?")
            if self.lower is None:
                self.lower=float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))
                self.scope.query("*OPC?")
            current_scale=float(self.scope.query("DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:WINSCALe?"))
            self.scope.query("*OPC?")
            print(current_scale)
            if (self.upper+self.lower)/10!=current_scale:
                self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {(self.upper+self.lower)/10}")
            #print(current_position)
            #self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 30")
            
        except Exception as e:
            return False, e
        return True, ":)"
    
    def zoom_setup_time(self):
        #find 2nd ch2 falling edge
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # find 1st ch2 rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce CH2")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION -2000")
            for i in range(1):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")

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
            self.scope.query("*OPC?")

            if self.upper is None:
                self.upper=float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))
                self.scope.query("*OPC?")
            if self.lower is None:
                self.lower=float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))
                self.scope.query("*OPC?")
            current_scale=float(self.scope.query("DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:WINSCALe?"))
            self.scope.query("*OPC?")
            print(current_scale)
            if (self.upper+self.lower)/10!=current_scale:
                self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {(self.upper+self.lower)/10}")
            #print(current_position)
            #self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 30")
            
        except Exception as e:
            return False, e
        return True, ":)"
    
    def zoom_start(self):
        #find 2nd ch2 falling edge
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # find 1st ch2 rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce CH2")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe FALL")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION -2000")
            for i in range(1):
                self.scope.write("SEARCH:SEARCH1:NAVigate NEXT")
                self.scope.query("*OPC?")

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
            self.scope.query("*OPC?")

            if self.upper is None:
                self.upper=float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))
                self.scope.query("*OPC?")
            if self.lower is None:
                self.lower=float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))
                self.scope.query("*OPC?")
            current_scale=float(self.scope.query("DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:WINSCALe?"))
            self.scope.query("*OPC?")
            print(current_scale)
            if (self.upper+self.lower)/10!=current_scale:
                self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {(self.upper+self.lower)/5}")
            #print(current_position)
            #self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 30")
            
        except Exception as e:
            return False, e
        return True, ":)"
    
    def zoom_stop(self):
        #find 2nd ch2 falling edge
        if not self.connected:
            return False, "未连接到示波器"
        try:
            self.clear_measurements()
            # find 1st ch2 rising edge
            self.scope.write("SEARCH:DELETEALL")
            self.scope.write("SEARCH:ADDNew SEARCH1")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:TYPe EDGE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SOUrce CH2")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:SLOPe RISE")
            self.scope.write("SEARCH:SEARCH1:TRIGger:A:EDGE:THReshold 1.5")
            self.scope.write("SEARCH:SEARCH1:STATE ON")
            self.scope.query("*OPC?")
            
            
            self.scope.write("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION 2000")
            for i in range(1):
                self.scope.write("SEARCH:SEARCH1:NAVigate PREVious")
                self.scope.query("*OPC?")

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
            self.scope.query("*OPC?")

            if self.upper is None:
                self.upper=float(self.scope.query("MEASUREMENT:MEAS2:VALUE?"))
                self.scope.query("*OPC?")
            if self.lower is None:
                self.lower=float(self.scope.query("MEASUREMENT:MEAS3:VALUE?"))
                self.scope.query("*OPC?")
            current_scale=float(self.scope.query("DISplay:WAVEView1:ZOOM:ZOOM1:HORizontal:WINSCALe?"))
            self.scope.query("*OPC?")
            print(current_scale)
            if (self.upper+self.lower)/10!=current_scale:
                self.scope.write(f":DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {(self.upper+self.lower)/5}")
            return True, ":)"
        except Exception as e:
            return False, e


class MSO64ControllerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MSO64示波器控制器")
        self.root.geometry("1200x700")
        self.controller = MSO64Controller()
        self.create_widgets()
        self.measured_voltage=False
        self.measured_frequency=False
        self.measured_setup=False
        self.res={}
        self.images={}#schema: name, file path
        self.excels=1
        self.measured_rise_fall=False
        self.measured_hold=False
        self.measured_start_hold=False
        self.measured_stop_setup=False
        
        
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
        self.voltage_button = ttk.Button(test_frame, text="添加电压测量", command=self.add_voltage_measurements)
        self.voltage_button.grid(row=3, column=0, padx=5, pady=5)
        
        self.frequency_button = ttk.Button(test_frame, text="添加频率测量", command=self.add_frequency_measurements)
        self.frequency_button.grid(row=3, column=1, padx=5, pady=5)
        
        self.delay_button = ttk.Button(test_frame, text="添加Setup Time测量", command=self.add_delay_measurement)
        self.delay_button.grid(row=3, column=2, padx=5, pady=5)

        self.rise_fall_button = ttk.Button(test_frame, text="添加Rise/Fall测量", command=self.add_rise_fall_measurements)
        self.rise_fall_button.grid(row=3, column=3, padx=5, pady=5)

        self.hold_time_button = ttk.Button(test_frame, text="添加Hold Time测量", command=self.add_hold_time_measurement)
        self.hold_time_button.grid(row=3, column=4, padx=5, pady=5)

        self.start_hold_time_button = ttk.Button(test_frame, text="添加Start Hold Time测量", command=self.add_start_hold_time_measurement)
        self.start_hold_time_button.grid(row=3, column=5, padx=5, pady=5)

        self.stop_setup_time_button = ttk.Button(test_frame, text="添加Stop Setup Time测量", command=self.add_stop_setup_time_measurement)
        self.stop_setup_time_button.grid(row=3, column=6, padx=5, pady=5)

        self.run_all_tests_button = ttk.Button(test_frame, text="Write Cycle", command=self.run_all_tests)
        self.run_all_tests_button.grid(row=4, column=0, padx=5, pady=5)

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
        
        # Setup Time测量选项卡
        self.delay_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.delay_tab, text="Setup Time参数")

        self.rise_fall_tab=ttk.Frame(self.tab_control)
        self.tab_control.add(self.rise_fall_tab, text="Rise/Fall参数")

        self.hold_time_tab=ttk.Frame(self.tab_control)
        self.tab_control.add(self.hold_time_tab, text="Hold Time参数")

        self.start_hold_time_tab=ttk.Frame(self.tab_control)
        self.tab_control.add(self.start_hold_time_tab, text="Start Hold Time参数")

        self.stop_setup_time_tab=ttk.Frame(self.tab_control)
        self.tab_control.add(self.stop_setup_time_tab, text="Stop Setup Time参数")

        # 运行测试按钮
        self.run_test_button = ttk.Button(test_frame, text="Save to Excel", command=self.save_excel)
        self.run_test_button.grid(row=3, column=7, padx=5, pady=5)

        # 创建电压测量参数表
        self.create_voltage_tab()
        
        # 创建频率测量参数表
        self.create_frequency_tab()
        
        # 创建Setup Time测量参数表
        self.create_delay_tab()

        self.create_rise_fall_tab()

        self.create_hold_time_tab()

        self.create_start_hold_time_tab()

        self.create_stop_setup_time_tab()

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
        if self.controller.excel_path is None:
            messagebox.showerror("错误", "File path has not been configured!")
            return
        if self.controller.image_path is None:
            messagebox.showerror("错误", "Image path has not been configured!")
            return
        self.add_voltage_measurements()
        self.get_voltage_measurements()
        self.save_screenshot("Voltage")
        self.add_frequency_measurements()
        self.get_frequency_measurements()
        self.save_screenshot("Frequency")
        self.add_delay_measurement()
        self.get_delay_measurement()
        self.save_screenshot("Setup Time")
        self.add_rise_fall_measurements()
        self.get_rise_fall_measurements()
        self.save_screenshot("Rise Fall")
        self.add_hold_time_measurement()
        self.get_hold_time_measurement()
        self.save_screenshot("Hold Time")
        self.add_start_hold_time_measurement()
        self.get_start_hold_time_measurement()
        self.save_screenshot("Start Hold Time")
        self.add_stop_setup_time_measurement()
        self.get_stop_setup_time_measurement()
        self.save_screenshot("Stop Setup Time")
        self.controller.zoom_on_rising_edge()
        self.save_screenshot("Rising Edge")
        self.controller.zoom_on_falling_edge()
        self.save_screenshot("Falling Edge")
        self.save_excel()

    
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

    def toggle_image_path(self):
        """切换图片保存位置为示波器或本地路径"""
        if self.save_to_scope_image.get():
            self.image_path_entry.config(state="disabled")
            self.controller.image_path = "oscilloscope"
            self.log_message("图片将保存到示波器")
        else:
            self.image_path_entry.config(state="normal")
            self.controller.image_path = self.image_path_entry.get()
            self.log_message(f"图片将保存到本地路径: {self.controller.image_path}")


    def add_voltage_measurements(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加电压测量项...")
        self.controller.zoom_middle()
        success, message = self.controller.add_voltage_measurements()
        self.log_message(message)
        
        if success:
            self.measured_voltage=True
            messagebox.showinfo("成功", "电压测量项添加成功\n")
            # 切换到电压选项卡
            self.tab_control.select(self.voltage_tab)
        else:
            messagebox.showerror("错误", message)

    def add_rise_fall_measurements(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Rise/Fall测量项...")
        self.controller.zoom_middle()
        success, message = self.controller.add_rise_fall_measurements()
        self.log_message(message)
        
        if success:
            self.measured_rise_fall=True
            messagebox.showinfo("成功", "Rise/Fall测量项添加成功\n")
            # 切换到电压选项卡
            self.tab_control.select(self.rise_fall_tab)
        else:
            messagebox.showerror("错误", message)

    def add_frequency_measurements(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加频率测量项...")
        self.controller.zoom_middle()
        success, message = self.controller.add_frequency_measurements()
        self.log_message(message)
        
        if success:
            self.measured_frequency=True
            messagebox.showinfo("成功", "频率测量项添加成功\n")
            # 切换到频率选项卡
            self.tab_control.select(self.frequency_tab)
        else:
            messagebox.showerror("错误", message)

    def add_delay_measurement(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Setup Time测量项...")
        self.controller.zoom_setup_time()
        success, message = self.controller.add_delay_measurement()
        self.log_message(message)
        
        if success:
            self.measured_setup=True
            messagebox.showinfo("成功", "Setup Time测量项添加成功\n")
            # 切换到延迟选项卡
            self.tab_control.select(self.delay_tab)
        else:
            messagebox.showerror("错误", message)

    def add_hold_time_measurement(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Hold Time测量项...")
        self.controller.zoom_hold_time()
        success, message = self.controller.add_hold_time_measurements()
        self.log_message(message)
        
        if success:
            self.measured_hold=True
            messagebox.showinfo("成功", "Hold Time测量项添加成功\n")
            # 切换到延迟选项卡
            self.tab_control.select(self.hold_time_tab)
        else:
            messagebox.showerror("错误", message)

    def add_start_hold_time_measurement(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Start Hold Time测量项...")
        self.controller.zoom_start()
        success, message = self.controller.add_delay2_measurement()
        self.log_message(message)
        
        if success:
            self.measured_start_hold=True
            messagebox.showinfo("成功", "Start Hold Time测量项添加成功\n")
            # 切换到延迟选项卡
            self.tab_control.select(self.start_hold_time_tab)
        else:
            messagebox.showerror("错误", message)

    def add_stop_setup_time_measurement(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Start Hold Time测量项...")
        self.controller.zoom_stop()
        success, message = self.controller.add_stop_time_measurement()
        self.log_message(message)
        
        if success:
            self.measured_stop_setup=True
            messagebox.showinfo("成功", "Stop Setup Time测量项添加成功\n")
            # 切换到延迟选项卡
            self.tab_control.select(self.stop_setup_time_tab)
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
        ttk.Label(self.voltage_tab, text="CH1 电压参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Maximum:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.ch1_max_value = ttk.Label(self.voltage_tab, text="--")
        self.ch1_max_value.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Minimum:").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.ch1_min_value = ttk.Label(self.voltage_tab, text="--")
        self.ch1_min_value.grid(row=2, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Top:").grid(row=3, column=0, padx=5, pady=2, sticky="w")
        self.ch1_top_value = ttk.Label(self.voltage_tab, text="--")
        self.ch1_top_value.grid(row=3, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Base:").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.ch1_base_value = ttk.Label(self.voltage_tab, text="--")
        self.ch1_base_value.grid(row=4, column=1, padx=5, pady=2, sticky="w")
        
        # CH2电压参数
        ttk.Label(self.voltage_tab, text="CH2 电压参数", font=("Arial", 10, "bold")).grid(row=0, column=2, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Maximum:").grid(row=1, column=2, padx=5, pady=2, sticky="w")
        self.ch2_max_value = ttk.Label(self.voltage_tab, text="--")
        self.ch2_max_value.grid(row=1, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Minimum:").grid(row=2, column=2, padx=5, pady=2, sticky="w")
        self.ch2_min_value = ttk.Label(self.voltage_tab, text="--")
        self.ch2_min_value.grid(row=2, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Top:").grid(row=3, column=2, padx=5, pady=2, sticky="w")
        self.ch2_top_value = ttk.Label(self.voltage_tab, text="--")
        self.ch2_top_value.grid(row=3, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.voltage_tab, text="Base:").grid(row=4, column=2, padx=5, pady=2, sticky="w")
        self.ch2_base_value = ttk.Label(self.voltage_tab, text="--")
        self.ch2_base_value.grid(row=4, column=3, padx=5, pady=2, sticky="w")
        
        # 获取电压测量值按钮
        self.get_voltage_button = ttk.Button(self.voltage_tab, text="获取电压测量值", command=self.get_voltage_measurements)
        self.get_voltage_button.grid(row=5, column=0, columnspan=4, padx=5, pady=10)

        # Screenshot
        self.get_screenshot_button = ttk.Button(self.voltage_tab, text="Save Screenshot", command=lambda: self.save_screenshot("Voltage"))
        self.get_screenshot_button.grid(row=6, column=0, columnspan=4, padx=10, pady=10)

    
    def create_frequency_tab(self):
        """创建频率测量参数表"""
        ttk.Label(self.frequency_tab, text="CH1 频率参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.frequency_tab, text="Frequency (30% to 30%):").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.frequency_value = ttk.Label(self.frequency_tab, text="--")
        self.frequency_value.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.frequency_tab, text="Positive Pulse Width (70% to 70%):").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.pos_width_value = ttk.Label(self.frequency_tab, text="--")
        self.pos_width_value.grid(row=2, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.frequency_tab, text="Negative Pulse Width (30% to 30%):").grid(row=3, column=0, padx=5, pady=2, sticky="w")
        self.neg_width_value = ttk.Label(self.frequency_tab, text="--")
        self.neg_width_value.grid(row=3, column=1, padx=5, pady=2, sticky="w")
        
        # 获取频率测量值按钮
        self.get_frequency_button = ttk.Button(self.frequency_tab, text="获取频率测量值", command=self.get_frequency_measurements)
        self.get_frequency_button.grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        # Screenshot
        self.frequency_screenshot_button = ttk.Button(self.frequency_tab, text="Save Screenshot", command=lambda: self.save_screenshot("Frequency"))
        self.frequency_screenshot_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)
    
    def create_delay_tab(self):
        """创建setup time测量参数表"""
        ttk.Label(self.delay_tab, text="Setup Time参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.delay_tab, text="Setup Time (CH2上升沿70% 到 CH1上升沿30%):").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.delay_value = ttk.Label(self.delay_tab, text="--")
        self.delay_value.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        # 获取setup time测量值按钮
        self.get_delay_button = ttk.Button(self.delay_tab, text="获取Setup Time测量值", command=self.get_delay_measurement)
        self.get_delay_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

        # Screenshot
        self.setup_screenshot_button = ttk.Button(self.delay_tab, text="Save Screenshot", command=lambda: self.save_screenshot("Setup Time"))
        self.setup_screenshot_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5)
    
    def create_hold_time_tab(self):
        """创建hold time测量参数表"""
        ttk.Label(self.hold_time_tab, text="Hold Time参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.hold_time_tab, text="Hold Time (CH1下降沿30% 到 CH2下降沿70%):").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.hold_time_value = ttk.Label(self.hold_time_tab, text="--")
        self.hold_time_value.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        # 获取setup time测量值按钮
        self.get_hold_time_button = ttk.Button(self.hold_time_tab, text="获取Hold Time测量值", command=self.get_hold_time_measurement)
        self.get_hold_time_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

        # Screenshot
        self.hold_time_screenshot_button = ttk.Button(self.hold_time_tab, text="Save Screenshot", command=lambda: self.save_screenshot("Hold Time"))
        self.hold_time_screenshot_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5)
    
    def create_start_hold_time_tab(self):
        """创建hold time测量参数表"""
        ttk.Label(self.start_hold_time_tab, text="Start Hold Time参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.start_hold_time_tab, text="Start Hold Time (CH2下降沿70% 到 CH1下降沿30%):").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.start_hold_time_value = ttk.Label(self.start_hold_time_tab, text="--")
        self.start_hold_time_value.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        # 获取setup time测量值按钮
        self.get_start_hold_time_button = ttk.Button(self.start_hold_time_tab, text="获取Start Hold Time测量值", command=self.get_start_hold_time_measurement)
        self.get_start_hold_time_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

        # Screenshot
        self.start_hold_time_screenshot_button = ttk.Button(self.start_hold_time_tab, text="Save Screenshot", command=lambda: self.save_screenshot("Start Hold Time"))
        self.start_hold_time_screenshot_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5)
    
    def create_stop_setup_time_tab(self):
        """创建hold time测量参数表"""
        ttk.Label(self.stop_setup_time_tab, text="Stop Setup Time参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.stop_setup_time_tab, text="Stop Setup Time (CH1上升沿30% 到 CH2上升沿70%):").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.stop_setup_time_value = ttk.Label(self.stop_setup_time_tab, text="--")
        self.stop_setup_time_value.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        # 获取setup time测量值按钮
        self.get_stop_setup_time_button = ttk.Button(self.stop_setup_time_tab, text="获取Stop Setup Time测量值", command=self.get_stop_setup_time_measurement)
        self.get_stop_setup_time_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

        # Screenshot
        self.stop_setup_time_screenshot_button = ttk.Button(self.stop_setup_time_tab, text="Save Screenshot", command=lambda: self.save_screenshot("Stop Setup Time"))
        self.stop_setup_time_screenshot_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

    def create_rise_fall_tab(self):
        """创建rise/fall测量参数表"""
        # CH1电压参数
        ttk.Label(self.rise_fall_tab, text="Rise/Fall Times", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.rise_fall_tab, text="CH1 Risetime: ").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.ch1_risetime = ttk.Label(self.rise_fall_tab, text="--")
        self.ch1_risetime.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.rise_fall_tab, text="CH1 Falltime: ").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.ch1_falltime = ttk.Label(self.rise_fall_tab, text="--")
        self.ch1_falltime.grid(row=2, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.rise_fall_tab, text="CH2 Risetime: ").grid(row=3, column=0, padx=5, pady=2, sticky="w")
        self.ch2_risetime = ttk.Label(self.rise_fall_tab, text="--")
        self.ch2_risetime.grid(row=3, column=1, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.rise_fall_tab, text="CH2 Falltime: ").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.ch2_falltime = ttk.Label(self.rise_fall_tab, text="--")
        self.ch2_falltime.grid(row=4, column=1, padx=5, pady=2, sticky="w")
        
        
        # 获取电压测量值按钮
        self.get_rise_fall_button = ttk.Button(self.rise_fall_tab, text="获取Rise/Fall测量值", command=self.get_rise_fall_measurements)
        self.get_rise_fall_button.grid(row=5, column=0, columnspan=4, padx=5, pady=10)

        # Screenshot
        self.get_rise_fall_screenshot_button = ttk.Button(self.rise_fall_tab, text="Save Screenshot", command=lambda: self.save_screenshot("RiseFall"))
        self.get_rise_fall_screenshot_button.grid(row=6, column=0, columnspan=4, padx=10, pady=10)

    def get_voltage_measurements(self):
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
            self.ch1_max_value.config(text=f"{measurements['CH1_Maximum']} V")
            self.ch1_min_value.config(text=f"{measurements['CH1_Minimum']} V")
            self.ch1_top_value.config(text=f"{measurements['CH1_Top']} V")
            self.ch1_base_value.config(text=f"{measurements['CH1_Base']} V")
            
            self.ch2_max_value.config(text=f"{measurements['CH2_Maximum']} V")
            self.ch2_min_value.config(text=f"{measurements['CH2_Minimum']} V")
            self.ch2_top_value.config(text=f"{measurements['CH2_Top']} V")
            self.ch2_base_value.config(text=f"{measurements['CH2_Base']} V")
            
            messagebox.showinfo("成功", "电压测量值获取成功")
        else:
            messagebox.showerror("错误", message)
    
    def get_rise_fall_measurements(self):
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
            self.ch1_risetime.config(text=f"{measurements['CH1_Rise']} s")
            self.ch1_falltime.config(text=f"{measurements['CH1_Fall']} s")
            self.ch2_risetime.config(text=f"{measurements['CH2_Rise']} s")
            self.ch2_falltime.config(text=f"{measurements['CH2_Fall']} s")
            
            messagebox.showinfo("成功", "Rise/Fall测量值获取成功")
        else:
            messagebox.showerror("错误", message)
    
    def get_frequency_measurements(self):
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
            self.frequency_value.config(text=f"{measurements['Frequency']} Hz")
            self.pos_width_value.config(text=f"{measurements['Positive_Pulse_Width']} s")
            self.neg_width_value.config(text=f"{measurements['Negative_Pulse_Width']} s")
            
            messagebox.showinfo("成功", "频率测量值获取成功")
        else:
            messagebox.showerror("错误", message)
    
    def get_delay_measurement(self):
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
            self.delay_value.config(text=f"{measurements['Delay_Rise_Edge_70_30']} s")
            
            messagebox.showinfo("成功", "Setup Time测量值获取成功")
        else:
            messagebox.showerror("错误", message)

    def get_hold_time_measurement(self):
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
            self.hold_time_value.config(text=f"{measurements['Delay_Fall_Edge_30_70']} s")
            
            messagebox.showinfo("成功", "Hold Time测量值获取成功")
        else:
            messagebox.showerror("错误", message)

    def get_start_hold_time_measurement(self):
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
            self.start_hold_time_value.config(text=f"{measurements['Start Hold Time']} s")
            
            messagebox.showinfo("成功", "Start Hold Time测量值获取成功")
        else:
            messagebox.showerror("错误", message)

    def get_stop_setup_time_measurement(self):
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
            self.stop_setup_time_value.config(text=f"{measurements['Stop Time']} s")
            
            messagebox.showinfo("成功", "Stop Setup Time测量值获取成功")
        else:
            messagebox.showerror("错误", message)


    def run_test(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        # 检查图片保存路径
        if not self.save_to_scope_image.get():
            image_path = self.image_path_entry.get().strip()
            if not image_path:
                messagebox.showerror("错误", "请选择图片保存路径")
                return
            self.controller.image_path = image_path
        
        # 检查Excel保存路径
        """
        if not self.save_to_scope_excel.get():
            excel_path = self.excel_path_entry.get().strip()
            if not excel_path:
                messagebox.showerror("错误", "请选择Excel保存路径")
                return
            self.controller.excel_path = excel_path
        """

        folder_name = self.folder_entry.get().strip()
        if not folder_name:
            messagebox.showerror("错误", "请输入文件夹名称")
            return
        
        # 重置状态标签
        self.voltage_status.config(text="处理中...")
        self.freq_status.config(text="处理中...")
        self.setup_status.config(text="处理中...")
        
        # 更新UI
        self.root.update()
        
        # 调用运行测试方法
        self.log_message(f"正在运行测试，文件夹名称: {folder_name}")
        success, message = self.controller.run_test(folder_name)
        self.log_message(message)
        
        if success:
            messagebox.showinfo("成功", "测试运行完成")
            
            # 更新保存状态
            if "Voltage.png: 成功" in message:
                self.voltage_status.config(text="成功")
            else:
                self.voltage_status.config(text="失败")
                
            if "Freq.png: 成功" in message:
                self.freq_status.config(text="成功")
            else:
                self.freq_status.config(text="失败")
                
            if "setup_time.png: 成功" in message:
                self.setup_status.config(text="成功")
            else:
                self.setup_status.config(text="失败")
        else:
            messagebox.showerror("错误", message)
            self.voltage_status.config(text="失败")
            self.freq_status.config(text="失败")
            self.setup_status.config(text="失败")

    def save_screenshot(self, filename):
        if self.controller.image_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        foldername=""
        if self.folder_entry.get()!="":
            foldername=self.folder_entry.get()
        else:
            current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            foldername=f"Test{current_time}"
            self.folder_entry.insert(0, f"Test{current_time}")
        suc, msg=self.controller.save_screenshot(filename, foldername)
        if foldername not in self.controller.image_path:
            self.images[filename]=self.controller.image_path+f"\\{foldername}\\{filename}.png"
        else:
            self.images[filename]=self.controller.image_path+f"\\{filename}.png"
        print(self.images)
        self.log_message(msg)

    def save_excel(self):
        if self.controller.excel_path is None:
            messagebox.showerror("错误", "File path has not been configured")
            return
        foldername=""
        if self.folder_entry.get()!="":
            foldername=self.folder_entry.get()
        else:
            current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            foldername=f"Test{current_time}"
            self.folder_entry.insert(0, f"Test{current_time}")
        suc, msg=self.controller.save_data_to_excel(self.res, "results"+str(self.excels), foldername, self.images)
        self.excels+=1
        self.log_message(msg)

    def update_button_states(self, connected):
        state = "normal" if connected else "disabled"
        self.voltage_button.config(state=state)
        self.frequency_button.config(state=state)
        self.delay_button.config(state=state)
        self.run_test_button.config(state=state)
        self.get_voltage_button.config(state=state)
        self.get_frequency_button.config(state=state)
        self.get_delay_button.config(state=state)
        self.get_screenshot_button.config(state=state)
        self.frequency_screenshot_button.config(state=state)
        self.setup_screenshot_button.config(state=state)
        self.rise_fall_button.config(state=state)
        self.get_rise_fall_button.config(state=state)
        self.get_rise_fall_screenshot_button.config(state=state)
        self.hold_time_button.config(state=state)
        self.get_hold_time_button.config(state=state)
        self.hold_time_screenshot_button.config(state=state)
        self.start_hold_time_button.config(state=state)
        self.get_stop_setup_time_button.config(state=state)
        self.stop_setup_time_screenshot_button.config(state=state)
        self.stop_setup_time_button.config(state=state)
        self.get_start_hold_time_button.config(state=state)
        self.start_hold_time_screenshot_button.config(state=state)
        self.run_all_tests_button.config(state=state)

    

    

if __name__ == "__main__":
    root = tk.Tk()
    app = MSO64ControllerGUI(root)
    root.mainloop() 
    
"""
controller=MSO64Controller()
controller.connect('169.254.123.124')
suc, msg=controller.zoom_stop()
controller.add_stop_time_measurement()
"""

