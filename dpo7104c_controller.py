import os
import time
import pandas as pd
from datetime import datetime
import pyvisa

class DPO7104Controller:
    def __init__(self):
        self.rm = pyvisa.ResourceManager()
        self.scope = None
        self.connected = False
        self.image_path = None  
        self.excel_path = None

    def measure_voltage(self, position, scl, sda):
        return