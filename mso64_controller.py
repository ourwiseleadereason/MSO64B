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
        self.test_folder_path=""
        
    def connect(self, ip_address):
        """连接到示波器"""
        try:
            resource_string = f"TCPIP::{ip_address}::INSTR"
            self.scope = self.rm.open_resource(resource_string)
            self.scope.timeout = 10000  # 设置超时时间为10秒
            idn = self.scope.query("*IDN?")
            if "MSO64" in idn:
                self.connected = True
                #self.scope.write('FILESYSTEM:CWD "C:\\ST\\hsd\\i2c"')
                #print(self.scope.query("FILESystem:CWD?"))
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
    
    def add_frequency_measurements(self):
        """添加频率和脉宽测量参数"""
        if not self.connected:
            return False, "未连接到示波器"
        
        try:
            # 清除所有现有测量
            self.clear_measurements()
            
            # 频率 (从30%到30%)
            self.scope.write("MEASUREMENT:ADDMEAS FREQUENCY")
            self.scope.write("MEASUREMENT:MEAS1:SOURCE CH1")
            # 设置参考电平方法为百分比
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:METHod PERCent")
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
            
            # 设置CH2(SOURCE1)的参考电平为70%
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEHigh 90")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISEMid 70")
            self.scope.write("MEASUREMENT:MEAS1:REFLevels:PERCent:RISELow 10")
            
            # 设置CH1(SOURCE2)的参考电平为30%
            self.scope.write("MEASUREMENT:CH1:REFLevels:PERCent:RISE2High 90")
            self.scope.write("MEASUREMENT:CH1:REFLevels:PERCent:RISE2Mid 30")
            self.scope.write("MEASUREMENT:CH1:REFLevels:PERCent:RISE2Low 10")
            
            # 设置测量的边沿和方向
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION1 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:DIRECTION2 FORWARDS")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE1 RISE")
            self.scope.write("MEASUREMENT:MEAS1:DELAY:EDGE2 RISE")
            
            return True, "setup time测量项添加成功"
        except Exception as e:
            return False, f"添加setup time测量项失败: {str(e)}"
    
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
    
    def save_screenshot(self, filename):
        """保存示波器截图到C:\ST\hsd\i2c目录"""
        if not self.connected:
            return False, "未连接到示波器"
        
        try:
            # 固定保存路径为C:\ST\hsd\i2c
            file_path = f"C:/ST/hsd/i2c/{filename}"
            
            # 设置图片格式为PNG
            self.scope.write("SAVE:IMAGE:FILEFORMAT PNG")
            self.log_message(f"设置图片格式为PNG")
            
            # 发送SAVE:IMAGE命令，使用两个单引号包裹路径
            cmd = f"SAVE:IMAGE ''{file_path}''"
            self.scope.write(cmd)
            self.log_message(f"发送命令: {cmd}")
            
            # 等待保存完成
            time.sleep(2)
            
            # 检查命令是否执行成功
            try:
                # 尝试查询最后一个错误
                error = self.scope.query("SYSTem:ERRor?")
                if error and "No error" not in error:
                    self.log_message(f"保存截图时发生错误: {error}")
                    return False, f"保存截图失败: {error}"
            except Exception as e:
                self.log_message(f"查询错误时发生异常: {str(e)}")
            
            self.log_message(f"截图已保存至 {file_path}")
            return True, f"截图已保存至 {file_path}"
        except Exception as e:
            error_msg = f"保存截图失败: {str(e)}"
            self.log_message(error_msg)
            return False, error_msg
    
    def save_data_to_excel(self, folder_name, measurements):
        """将测量数据保存到Excel"""
        # 使用excel_path，如果没有设置则使用save_path
        if self.excel_path:
            excel_dir = self.excel_path
        else:
            excel_dir = self.save_path
            
        if not excel_dir:
            return False, "未设置Excel保存路径"
        
        # 检查是否保存到示波器上
        if excel_dir.lower() == "oscilloscope":
            excel_file = "C:/autotest_I2C.xlsx"
        else:
            excel_file = os.path.join(excel_dir, "autotest_I2C.xlsx")
            
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            # 如果文件存在，读取数据并附加新行
            if os.path.exists(excel_file):
                df = pd.read_excel(excel_file)
            else:
                # 创建新DataFrame
                columns = ["文件夹名称", "时间戳", "CH1_Maximum", "CH1_Minimum", "CH1_Top", "CH1_Base",
                          "CH2_Maximum", "CH2_Minimum", "CH2_Top", "CH2_Base", 
                          "Frequency", "Positive_Pulse_Width", "Negative_Pulse_Width", "Delay_Rise_Edge_70_30"]
                df = pd.DataFrame(columns=columns)
            
            # 添加新数据行
            new_row = {"文件夹名称": folder_name, "时间戳": timestamp}
            
            # 重命名Delay键，如果存在
            if "Delay" in measurements:
                measurements["Delay_Rise_Edge_70_30"] = measurements.pop("Delay")
            
            new_row.update(measurements)
            
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            
            # 保存文件
            df.to_excel(excel_file, index=False)
            
            return True, f"数据已保存至 {excel_file}"
        except Exception as e:
            return False, f"保存数据失败: {str(e)}"
    
    def run_test(self, folder_name):
        """运行完整测试流程"""
        if not self.connected:
            return False, "未连接到示波器"
        
        # 检查图片保存路径
        image_dir=None
        if self.image_path:
            image_dir = self.image_path
        else:
            image_dir = self.save_path
            
        if not image_dir:
            return False, "未设置图片保存路径"
        
        # 如果是保存到本地而非示波器，则创建文件夹
        if image_dir.lower() != "oscilloscope":
            folder_path = os.path.join(image_dir, folder_name)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            # 设置临时图片路径为此文件夹
            temp_image_path = self.image_path
            self.image_path = folder_path
        
        results = {}
        screenshot_status = {"Voltage.png": True, "Freq.png": True, "setup_time.png": True}
        
        try:
            # 添加电压测量并保存截图
            success, message = self.add_voltage_measurements()
            if not success:
                return False, message
            
            success, message, voltage_results = self.get_voltage_measurements()
            if success:
                results.update(voltage_results)
                success, message = self.save_screenshot("Voltage.png")
                if not success:
                    screenshot_status["Voltage.png"] = False
            else:
                screenshot_status["Voltage.png"] = False
            
            # 添加频率和脉宽测量并保存截图
            success, message = self.add_frequency_measurements()
            if not success:
                return False, message
            
            success, message, frequency_results = self.get_frequency_measurements()
            if success:
                results.update(frequency_results)
                success, message = self.save_screenshot("Freq.png")
                if not success:
                    screenshot_status["Freq.png"] = False
            else:
                screenshot_status["Freq.png"] = False
            
            # 添加Setup Time测量并保存截图
            success, message = self.add_delay_measurement()
            if not success:
                return False, message
            
            success, message, delay_results = self.get_delay_measurement()
            if success:
                results.update(delay_results)
                success, message = self.save_screenshot("setup_time.png")
                if not success:
                    screenshot_status["setup_time.png"] = False
            else:
                screenshot_status["setup_time.png"] = False
            
            # 保存数据到Excel
            success, message = self.save_data_to_excel(folder_name, results)
            if not success:
                return False, message
            
            # 还原临时保存路径
            if image_dir.lower() != "oscilloscope":
                self.image_path = temp_image_path
            
            # 返回测试结果信息
            status_info = ""
            for img, status in screenshot_status.items():
                status_text = "成功" if status else "失败"
                status_info += f"{img}: {status_text}, "
            
            return True, f"测试完成. {status_info.rstrip(', ')}"
        except Exception as e:
            # 出错时确保还原临时保存路径
            if image_dir.lower() != "oscilloscope" and 'temp_image_path' in locals():
                self.image_path = temp_image_path
            return False, f"测试过程出错: {str(e)}"


class MSO64ControllerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MSO64示波器控制器")
        self.root.geometry("800x700")
        
        self.controller = MSO64Controller()
        
        self.create_widgets()
    
    def create_widgets(self):
        # 创建标签框架
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
        self.save_to_scope_image = tk.BooleanVar()
        self.save_to_scope_image_check = ttk.Checkbutton(test_frame, text="保存到示波器", 
                                                          variable=self.save_to_scope_image,
                                                          command=self.toggle_image_path)
        self.save_to_scope_image_check.grid(row=0, column=4, padx=5, pady=5)
        
        # Excel保存路径
        ttk.Label(test_frame, text="Excel保存路径:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.excel_path_entry = ttk.Entry(test_frame, width=35)
        self.excel_path_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        
        self.excel_browse_button = ttk.Button(test_frame, text="浏览...", command=self.browse_excel_path)
        self.excel_browse_button.grid(row=1, column=3, padx=5, pady=5)
        
        # 保存到示波器选项
        self.save_to_scope_excel = tk.BooleanVar()
        self.save_to_scope_excel_check = ttk.Checkbutton(test_frame, text="保存到示波器", 
                                                         variable=self.save_to_scope_excel,
                                                         command=self.toggle_excel_path)
        self.save_to_scope_excel_check.grid(row=1, column=4, padx=5, pady=5)
        
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
        
        # 运行测试按钮
        self.run_test_button = ttk.Button(test_frame, text="运行完整测试", command=self.run_test)
        self.run_test_button.grid(row=3, column=3, padx=5, pady=5)
        
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
        
        # 创建电压测量参数表
        self.create_voltage_tab()
        
        # 创建频率测量参数表
        self.create_frequency_tab()
        
        # 创建Setup Time测量参数表
        self.create_delay_tab()
        
        # 创建状态面板
        status_frame = ttk.LabelFrame(self.root, text="保存状态")
        status_frame.pack(fill="x", padx=10, pady=5)
        
        # 创建状态标签
        ttk.Label(status_frame, text="Voltage.png:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.voltage_status = ttk.Label(status_frame, text="未保存")
        self.voltage_status.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(status_frame, text="Freq.png:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.freq_status = ttk.Label(status_frame, text="未保存")
        self.freq_status.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        ttk.Label(status_frame, text="setup_time.png:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.setup_status = ttk.Label(status_frame, text="未保存")
        self.setup_status.grid(row=0, column=5, padx=5, pady=5, sticky="w")
        
        # 日志区域
        log_frame = ttk.LabelFrame(self.root, text="日志")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.log_text = tk.Text(log_frame, height=8, width=70)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 禁用不能使用的按钮
        self.update_button_states(False)
    
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
        self.get_frequency_button.grid(row=4, column=0, columnspan=2, padx=5, pady=10)
    
    def create_delay_tab(self):
        """创建setup time测量参数表"""
        ttk.Label(self.delay_tab, text="Setup Time参数", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(self.delay_tab, text="Setup Time (CH2上升沿70% 到 CH1上升沿30%):").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.delay_value = ttk.Label(self.delay_tab, text="--")
        self.delay_value.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        # 获取setup time测量值按钮
        self.get_delay_button = ttk.Button(self.delay_tab, text="获取Setup Time测量值", command=self.get_delay_measurement)
        self.get_delay_button.grid(row=2, column=0, columnspan=2, padx=5, pady=10)
    
    def update_button_states(self, connected):
        state = "normal" if connected else "disabled"
        self.voltage_button.config(state=state)
        self.frequency_button.config(state=state)
        self.delay_button.config(state=state)
        self.run_test_button.config(state=state)
        self.get_voltage_button.config(state=state)
        self.get_frequency_button.config(state=state)
        self.get_delay_button.config(state=state)
    
    def log_message(self, message):
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
    
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
    
    def browse_image_path(self):
        path = filedialog.askdirectory()
        if path:
            self.image_path_entry.delete(0, tk.END)
            self.image_path_entry.insert(0, path)
            self.controller.image_path = path
            self.log_message(f"已选择图片保存路径: {path}")
    
    def browse_excel_path(self):
        path = filedialog.askdirectory()
        if path:
            self.excel_path_entry.delete(0, tk.END)
            self.excel_path_entry.insert(0, path)
            self.controller.excel_path = path
            self.log_message(f"已选择Excel保存路径: {path}")
    
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
    
    def toggle_excel_path(self):
        """切换Excel保存位置为示波器或本地路径"""
        if self.save_to_scope_excel.get():
            self.excel_path_entry.config(state="disabled")
            self.controller.excel_path = "oscilloscope"
            self.log_message("Excel将保存到示波器")
        else:
            self.excel_path_entry.config(state="normal")
            self.controller.excel_path = self.excel_path_entry.get()
            self.log_message(f"Excel将保存到本地路径: {self.controller.excel_path}")
    
    def add_voltage_measurements(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加电压测量项...")
        success, message = self.controller.add_voltage_measurements()
        self.log_message(message)
        
        if success:
            messagebox.showinfo("成功", "电压测量项添加成功")
            # 切换到电压选项卡
            self.tab_control.select(self.voltage_tab)
        else:
            messagebox.showerror("错误", message)
    
    def add_frequency_measurements(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加频率测量项...")
        success, message = self.controller.add_frequency_measurements()
        self.log_message(message)
        
        if success:
            messagebox.showinfo("成功", "频率测量项添加成功")
            # 切换到频率选项卡
            self.tab_control.select(self.frequency_tab)
        else:
            messagebox.showerror("错误", message)
    
    def add_delay_measurement(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在添加Setup Time测量项...")
        success, message = self.controller.add_delay_measurement()
        self.log_message(message)
        
        if success:
            messagebox.showinfo("成功", "Setup Time测量项添加成功")
            # 切换到延迟选项卡
            self.tab_control.select(self.delay_tab)
        else:
            messagebox.showerror("错误", message)
    
    def get_voltage_measurements(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在获取电压测量值...")
        success, message, measurements = self.controller.get_voltage_measurements()
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
    
    def get_frequency_measurements(self):
        if not self.controller.connected:
            messagebox.showerror("错误", "未连接到示波器")
            return
        
        self.log_message("正在获取频率测量值...")
        success, message, measurements = self.controller.get_frequency_measurements()
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
        
        self.log_message("正在获取Setup Time测量值...")
        success, message, measurements = self.controller.get_delay_measurement()
        self.log_message(message)
        
        if success:
            # 更新延迟测量值显示
            self.delay_value.config(text=f"{measurements['Delay_Rise_Edge_70_30']} s")
            
            messagebox.showinfo("成功", "Setup Time测量值获取成功")
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
        if not self.save_to_scope_excel.get():
            excel_path = self.excel_path_entry.get().strip()
            if not excel_path:
                messagebox.showerror("错误", "请选择Excel保存路径")
                return
            self.controller.excel_path = excel_path
        
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


if __name__ == "__main__":
    root = tk.Tk()
    app = MSO64ControllerGUI(root)
    root.mainloop() 