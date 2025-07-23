import os
import openpyxl
import pandas as pd
import csv

rootPath = os.getcwd()
planPath = os.path.join(rootPath, 'test plan', '13 I2C.xlsx')
wb = openpyxl.load_workbook(planPath)
sheet = wb['I2C']

