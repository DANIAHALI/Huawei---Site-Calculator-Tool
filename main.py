import pandas
import openpyxl
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.styles.borders import Border, Side
from openpyxl import load_workbook
from openpyxl.chart import BarChart3D, Reference, AreaChart, AreaChart3D, Series, LineChart, LineChart3D
import os
import pdb
import math
from openpyxl.styles import NamedStyle, Font, Border, Side
# import excel2img
import time
import datetime
import xlsxwriter
import math
import numpy as np

def Antenna_STATUS(Input_path, Output_path):

    print('Processing Start Please Wait!')
    start = time.time()
    print('Processing!')

    # Input_df = pandas.read_excel(str(os.getcwd()) + '\\Input' + '\\' + 'Input data.xlsx', sheet_name = 'Sheet1')
    Input_df = pandas.read_excel(Input_path, sheet_name = 'Sheet1')
    Sites_name = list(set(Input_df['Site'].tolist()))

    Antenna_Type_1_ = []
    Antenna_Type_2_ = []
    Antenna_Type_3_ = []
    Antenna_Type_4_ = []
    Antenna_Count_1_ = []
    Antenna_Count_2_ = []
    Antenna_Count_3_ = []
    Antenna_Count_4_ = []
    Total_Antenna = []


    for i in range(len(Sites_name)):

        Antenna_Type_1 = []
        Antenna_Type_2 = []
        Antenna_Type_3 = []
        Antenna_Type_4 = []
        Antenna_Count_1 = []
        Antenna_Count_2 = []
        Antenna_Count_3 = []
        Antenna_Count_4 = []

        Individual_sites_df = Input_df[Input_df['Site'] == Sites_name[i]]
        #
        # if Sites_name[i] == 'AJM3384':
        #     pdb.set_trace()

        Antenna_1 = Individual_sites_df['2G Antenna_EOY2020'].tolist()
        Antenna_2 = Individual_sites_df['3G Antenna_EOY2020'].tolist()
        Antenna_3 = Individual_sites_df['LTE800 antenna_EOY2020'].tolist()
        Antenna_4 = Individual_sites_df['LTE1800 antenna_EOY2020'].tolist()

        # Antenna 1 -----------------------------------------------------

        Antenna_1_ = list(set(Antenna_1))

        # if np.isnan(Antenna_1_).any():

        if len(Antenna_1_) == 1:
            if Antenna_1_[0] in Antenna_Type_1 or Antenna_1_[0] in Antenna_Type_2 or Antenna_1_[0] in Antenna_Type_3 or Antenna_1_[0] in Antenna_Type_4:
                Antenna_Type_1.append('')
                Antenna_Count_1.append('')
            else:
                Antenna_Type_1.append(Antenna_1_[0])
                Antenna_Count_1.append(Antenna_1.count(Antenna_1_[0]))


        elif len(Antenna_1_) == 2:

            if Antenna_1_[0] in Antenna_Type_1 or Antenna_1_[0] in Antenna_Type_2 or Antenna_1_[0] in Antenna_Type_3 or Antenna_1_[0] in Antenna_Type_4:
                if Antenna_1_[1] in Antenna_Type_1 or Antenna_1_[1] in Antenna_Type_2 or Antenna_1_[1] in Antenna_Type_3 or Antenna_1_[1] in Antenna_Type_4:
                    Antenna_Type_1.append('')
                    Antenna_Count_1.append('')
                else:
                    Antenna_Type_1.append(Antenna_1_[1])
                    Antenna_Count_1.append(Antenna_1.count(Antenna_1_[1]))

            else:
                Antenna_Type_1.append(Antenna_1_[0])
                Antenna_Count_1.append(Antenna_1.count(Antenna_1_[0]))

        elif len(Antenna_1_) == 3:

            if Antenna_1_[0] in Antenna_Type_1 or Antenna_1_[0] in Antenna_Type_2 or Antenna_1_[0] in Antenna_Type_3 or Antenna_1_[0] in Antenna_Type_4:
                if Antenna_1_[1] in Antenna_Type_1 or Antenna_1_[1] in Antenna_Type_2 or Antenna_1_[1] in Antenna_Type_3 or Antenna_1_[1] in Antenna_Type_4:
                    if Antenna_1_[2] in Antenna_Type_1 or Antenna_1_[2] in Antenna_Type_2 or Antenna_1_[2] in Antenna_Type_3 or Antenna_1_[2] in Antenna_Type_4:
                        Antenna_Type_1.append('')
                        Antenna_Count_1.append('')
                    else:
                        Antenna_Type_1.append(Antenna_1_[2])
                        Antenna_Count_1.append(Antenna_1.count(Antenna_1_[2]))
                else:
                    Antenna_Type_1.append(Antenna_1_[1])
                    Antenna_Count_1.append(Antenna_1.count(Antenna_1_[1]))

            else:
                Antenna_Type_1.append(Antenna_1_[0])
                Antenna_Count_1.append(Antenna_1.count(Antenna_1_[0]))

        elif len(Antenna_1_) == 4:

            if Antenna_1_[0] in Antenna_Type_1 or Antenna_1_[0] in Antenna_Type_2 or Antenna_1_[0] in Antenna_Type_3 or Antenna_1_[0] in Antenna_Type_4:
                if Antenna_1_[1] in Antenna_Type_1 or Antenna_1_[1] in Antenna_Type_2 or Antenna_1_[1] in Antenna_Type_3 or Antenna_1_[1] in Antenna_Type_4:
                    if Antenna_1_[2] in Antenna_Type_1 or Antenna_1_[2] in Antenna_Type_2 or Antenna_1_[2] in Antenna_Type_3 or Antenna_1_[2] in Antenna_Type_4:
                        if Antenna_1_[3] in Antenna_Type_1 or Antenna_1_[3] in Antenna_Type_2 or Antenna_1_[3] in Antenna_Type_3 or Antenna_1_[3] in Antenna_Type_4:
                            Antenna_Type_1.append('')
                            Antenna_Count_1.append('')
                        else:
                            Antenna_Type_1.append(Antenna_1_[3])
                            Antenna_Count_1.append(Antenna_1.count(Antenna_1_[3]))
                    else:
                        Antenna_Type_1.append(Antenna_1_[2])
                        Antenna_Count_1.append(Antenna_1.count(Antenna_1_[2]))
                else:
                    Antenna_Type_1.append(Antenna_1_[1])
                    Antenna_Count_1.append(Antenna_1.count(Antenna_1_[1]))

            else:
                Antenna_Type_1.append(Antenna_1_[0])
                Antenna_Count_1.append(Antenna_1.count(Antenna_1_[0]))

        else:
            Antenna_Type_1.append('')
            Antenna_Count_1.append('')

        # Antenna 2 -----------------------------------------------------

        Antenna_2_ = list(set(Antenna_2))

        if len(Antenna_2_) == 1:
            if Antenna_2_[0] in Antenna_Type_1 or Antenna_2_[0] in Antenna_Type_2 or Antenna_2_[0] in Antenna_Type_3 or Antenna_2_[0] in Antenna_Type_4:
                Antenna_Type_2.append('')
                Antenna_Count_2.append('')
            else:
                Antenna_Type_2.append(Antenna_2_[0])
                Antenna_Count_2.append(Antenna_2.count(Antenna_2_[0]))

        elif len(Antenna_2_) == 2:

            if Antenna_2_[0] in Antenna_Type_1 or Antenna_2_[0] in Antenna_Type_2 or Antenna_2_[0] in Antenna_Type_3 or Antenna_2_[0] in Antenna_Type_4:
                if Antenna_2_[1] in Antenna_Type_1 or Antenna_2_[1] in Antenna_Type_2 or Antenna_2_[1] in Antenna_Type_3 or Antenna_2_[1] in Antenna_Type_4:
                    Antenna_Type_2.append('')
                    Antenna_Count_2.append('')
                else:
                    Antenna_Type_2.append(Antenna_2_[1])
                    Antenna_Count_2.append(Antenna_2.count(Antenna_2_[1]))

            else:
                Antenna_Type_2.append(Antenna_2_[0])
                Antenna_Count_2.append(Antenna_2.count(Antenna_2_[0]))


        elif len(Antenna_2_) == 3:

            if Antenna_2_[0] in Antenna_Type_1 or Antenna_2_[0] in Antenna_Type_2 or Antenna_2_[0] in Antenna_Type_3 or Antenna_2_[0] in Antenna_Type_4:
                if Antenna_2_[1] in Antenna_Type_1 or Antenna_2_[1] in Antenna_Type_2 or Antenna_2_[1] in Antenna_Type_3 or Antenna_2_[1] in Antenna_Type_4:
                    if Antenna_2_[2] in Antenna_Type_1 or Antenna_2_[2] in Antenna_Type_2 or Antenna_2_[2] in Antenna_Type_3 or Antenna_2_[2] in Antenna_Type_4:
                        Antenna_Type_2.append('')
                        Antenna_Count_2.append('')
                    else:
                        Antenna_Type_2.append(Antenna_2_[2])
                        Antenna_Count_2.append(Antenna_2.count(Antenna_2_[2]))
                else:
                    Antenna_Type_2.append(Antenna_2_[1])
                    Antenna_Count_2.append(Antenna_2.count(Antenna_2_[1]))

            else:
                Antenna_Type_2.append(Antenna_2_[0])
                Antenna_Count_2.append(Antenna_2.count(Antenna_2_[0]))

        elif len(Antenna_2_) == 4:

            if Antenna_2_[0] in Antenna_Type_1 or Antenna_2_[0] in Antenna_Type_2 or Antenna_2_[0] in Antenna_Type_3 or Antenna_2_[0] in Antenna_Type_4:
                if Antenna_2_[1] in Antenna_Type_1 or Antenna_2_[1] in Antenna_Type_2 or Antenna_2_[1] in Antenna_Type_3 or Antenna_2_[1] in Antenna_Type_4:
                    if Antenna_2_[2] in Antenna_Type_1 or Antenna_2_[2] in Antenna_Type_2 or Antenna_2_[2] in Antenna_Type_3 or Antenna_2_[2] in Antenna_Type_4:
                        if Antenna_2_[3] in Antenna_Type_1 or Antenna_2_[3] in Antenna_Type_2 or Antenna_2_[3] in Antenna_Type_3 or Antenna_2_[3] in Antenna_Type_4:
                            Antenna_Type_2.append('')
                            Antenna_Count_2.append('')
                        else:
                            Antenna_Type_2.append(Antenna_2_[3])
                            Antenna_Count_2.append(Antenna_2.count(Antenna_2_[3]))
                    else:
                        Antenna_Type_2.append(Antenna_2_[2])
                        Antenna_Count_2.append(Antenna_2.count(Antenna_2_[2]))
                else:
                    Antenna_Type_2.append(Antenna_2_[1])
                    Antenna_Count_2.append(Antenna_2.count(Antenna_2_[1]))

            else:
                Antenna_Type_2.append(Antenna_2_[0])
                Antenna_Count_2.append(Antenna_2.count(Antenna_2_[0]))

        else:
            Antenna_Type_2.append('')
            Antenna_Count_2.append('')

        # if Sites_name[i] == 'AJM3384':
        #     pdb.set_trace()

        # Antenna 3 -----------------------------------------------------

        Antenna_3_ = list(set(Antenna_3))

        if len(Antenna_3_) == 1:
            if Antenna_3_[0] in Antenna_Type_1 or Antenna_3_[0] in Antenna_Type_2 or Antenna_3_[0] in Antenna_Type_3 or Antenna_3_[0] in Antenna_Type_4:
                Antenna_Type_3.append('')
                Antenna_Count_3.append('')
            else:
                Antenna_Type_3.append(Antenna_3_[0])
                Antenna_Count_3.append(Antenna_3.count(Antenna_3_[0]))
        elif len(Antenna_3_) == 2:

            if Antenna_3_[0] in Antenna_Type_1 or Antenna_3_[0] in Antenna_Type_2 or Antenna_3_[0] in Antenna_Type_3 or Antenna_3_[0] in Antenna_Type_4:
                if Antenna_3_[1] in Antenna_Type_1 or Antenna_3_[1] in Antenna_Type_2 or Antenna_3_[1] in Antenna_Type_3 or Antenna_3_[1] in Antenna_Type_4:
                    Antenna_Type_3.append('')
                    Antenna_Count_3.append('')
                else:
                    Antenna_Type_3.append(Antenna_3_[1])
                    Antenna_Count_3.append(Antenna_3.count(Antenna_3_[1]))
            else:
                Antenna_Type_3.append(Antenna_3_[0])
                Antenna_Count_3.append(Antenna_3.count(Antenna_3_[0]))

        elif len(Antenna_3_) == 3:

            if Antenna_3_[0] in Antenna_Type_1 or Antenna_3_[0] in Antenna_Type_2 or Antenna_3_[0] in Antenna_Type_3 or Antenna_3_[0] in Antenna_Type_4:
                if Antenna_3_[1] in Antenna_Type_1 or Antenna_3_[1] in Antenna_Type_2 or Antenna_3_[1] in Antenna_Type_3 or Antenna_3_[1] in Antenna_Type_4:
                    if Antenna_3_[2] in Antenna_Type_1 or Antenna_3_[2] in Antenna_Type_2 or Antenna_3_[2] in Antenna_Type_3 or Antenna_3_[2] in Antenna_Type_4:
                        Antenna_Type_3.append('')
                        Antenna_Count_3.append('')
                    else:
                        Antenna_Type_3.append(Antenna_3_[2])
                        Antenna_Count_3.append(Antenna_3.count(Antenna_3_[2]))
                else:
                    Antenna_Type_3.append(Antenna_3_[1])
                    Antenna_Count_3.append(Antenna_3.count(Antenna_3_[1]))
            else:
                Antenna_Type_3.append(Antenna_3_[0])
                Antenna_Count_3.append(Antenna_3.count(Antenna_3_[0]))

        elif len(Antenna_3_) == 4:

            if Antenna_3_[0] in Antenna_Type_1 or Antenna_3_[0] in Antenna_Type_2 or Antenna_3_[0] in Antenna_Type_3 or Antenna_3_[0] in Antenna_Type_4:
                if Antenna_3_[1] in Antenna_Type_1 or Antenna_3_[1] in Antenna_Type_2 or Antenna_3_[1] in Antenna_Type_3 or Antenna_3_[1] in Antenna_Type_4:
                    if Antenna_3_[2] in Antenna_Type_1 or Antenna_3_[2] in Antenna_Type_2 or Antenna_3_[2] in Antenna_Type_3 or Antenna_3_[2] in Antenna_Type_4:
                        if Antenna_3_[3] in Antenna_Type_1 or Antenna_3_[3] in Antenna_Type_2 or Antenna_3_[3] in Antenna_Type_3 or Antenna_3_[3] in Antenna_Type_4:
                            Antenna_Type_3.append('')
                            Antenna_Count_3.append('')
                        else:
                            Antenna_Type_3.append(Antenna_3_[3])
                            Antenna_Count_3.append(Antenna_3.count(Antenna_3_[3]))
                    else:
                        Antenna_Type_3.append(Antenna_3_[2])
                        Antenna_Count_3.append(Antenna_3.count(Antenna_3_[2]))
                else:
                    Antenna_Type_3.append(Antenna_3_[1])
                    Antenna_Count_3.append(Antenna_3.count(Antenna_3_[1]))
            else:
                Antenna_Type_3.append(Antenna_3_[0])
                Antenna_Count_3.append(Antenna_3.count(Antenna_3_[0]))

        else:
            Antenna_Type_3.append('')
            Antenna_Count_3.append('')

        # Antenna 4 -----------------------------------------------------

        Antenna_4_ = list(set(Antenna_4))

        if len(Antenna_4_) == 1:
            if Antenna_4_[0] in Antenna_Type_1 or Antenna_4_[0] in Antenna_Type_2 or Antenna_4_[0] in Antenna_Type_3 or Antenna_4_[0] in Antenna_Type_4:
                Antenna_Type_4.append('')
                Antenna_Count_4.append('')
            else:
                Antenna_Type_4.append(Antenna_4_[0])
                Antenna_Count_4.append(Antenna_4.count(Antenna_4_[0]))
        elif len(Antenna_4_) == 2:

            if Antenna_4_[0] in Antenna_Type_1 or Antenna_4_[0] in Antenna_Type_2 or Antenna_4_[0] in Antenna_Type_3 or Antenna_4_[0] in Antenna_Type_4:
                if Antenna_4_[1] in Antenna_Type_1 or Antenna_4_[1] in Antenna_Type_2 or Antenna_4_[1] in Antenna_Type_3 or Antenna_4_[1] in Antenna_Type_4:
                    Antenna_Type_4.append('')
                    Antenna_Count_4.append('')
                else:
                    Antenna_Type_4.append(Antenna_4_[1])
                    Antenna_Count_4.append(Antenna_4.count(Antenna_4_[1]))
            else:
                Antenna_Type_4.append(Antenna_4_[0])
                Antenna_Count_4.append(Antenna_4.count(Antenna_4_[0]))

        elif len(Antenna_4_) == 3:

            if Antenna_4_[0] in Antenna_Type_1 or Antenna_4_[0] in Antenna_Type_2 or Antenna_4_[0] in Antenna_Type_3 or Antenna_4_[0] in Antenna_Type_4:
                if Antenna_4_[1] in Antenna_Type_1 or Antenna_4_[1] in Antenna_Type_2 or Antenna_4_[1] in Antenna_Type_3 or Antenna_4_[1] in Antenna_Type_4:
                    if Antenna_4_[2] in Antenna_Type_1 or Antenna_4_[2] in Antenna_Type_2 or Antenna_4_[2] in Antenna_Type_3 or Antenna_4_[2] in Antenna_Type_4:
                        Antenna_Type_4.append('')
                        Antenna_Count_4.append('')
                    else:
                        Antenna_Type_4.append(Antenna_4_[2])
                        Antenna_Count_4.append(Antenna_4.count(Antenna_4_[2]))
                else:
                    Antenna_Type_4.append(Antenna_4_[1])
                    Antenna_Count_4.append(Antenna_4.count(Antenna_4_[1]))
            else:
                Antenna_Type_4.append(Antenna_4_[0])
                Antenna_Count_4.append(Antenna_4.count(Antenna_4_[0]))

        elif len(Antenna_4_) == 4:

            if Antenna_4_[0] in Antenna_Type_1 or Antenna_4_[0] in Antenna_Type_2 or Antenna_4_[0] in Antenna_Type_3 or Antenna_4_[0] in Antenna_Type_4:
                if Antenna_4_[1] in Antenna_Type_1 or Antenna_4_[1] in Antenna_Type_2 or Antenna_4_[1] in Antenna_Type_3 or Antenna_4_[1] in Antenna_Type_4:
                    if Antenna_4_[2] in Antenna_Type_1 or Antenna_4_[2] in Antenna_Type_2 or Antenna_4_[2] in Antenna_Type_3 or Antenna_4_[2] in Antenna_Type_4:
                        if Antenna_4_[3] in Antenna_Type_1 or Antenna_4_[3] in Antenna_Type_2 or Antenna_4_[3] in Antenna_Type_3 or Antenna_4_[3] in Antenna_Type_4:
                            Antenna_Type_4.append('')
                            Antenna_Count_4.append('')
                        else:
                            Antenna_Type_4.append(Antenna_4_[3])
                            Antenna_Count_4.append(Antenna_4.count(Antenna_4_[3]))
                    else:
                        Antenna_Type_4.append(Antenna_4_[2])
                        Antenna_Count_4.append(Antenna_4.count(Antenna_4_[2]))
                else:
                    Antenna_Type_4.append(Antenna_4_[1])
                    Antenna_Count_4.append(Antenna_4.count(Antenna_4_[1]))
            else:
                Antenna_Type_4.append(Antenna_4_[0])
                Antenna_Count_4.append(Antenna_4.count(Antenna_4_[0]))

        else:
            Antenna_Type_4.append('')
            Antenna_Count_4.append('')

        # pdb.set_trace()
        Antenna_Type_1_.append(Antenna_Type_1[0])
        Antenna_Type_2_.append(Antenna_Type_2[0])
        Antenna_Type_3_.append(Antenna_Type_3[0])
        Antenna_Type_4_.append(Antenna_Type_4[0])

        Sum = 0

        try:
            if math.isnan(Antenna_Type_1[0]):
                Antenna_Count_1_.append('')
            else:
                Antenna_Count_1_.append(Antenna_Count_1[0])
                if Antenna_Count_1[0] != '' :
                    Sum = Sum + int(Antenna_Count_1[0])
                else:
                    Sum = Sum + 0
        except:
            Antenna_Count_1_.append(Antenna_Count_1[0])
            if Antenna_Count_1[0] != '':
                Sum = Sum + int(Antenna_Count_1[0])
            else:
                Sum = Sum + 0

        try:
            if math.isnan(Antenna_Type_2[0]):
                Antenna_Count_2_.append('')
            else:
                Antenna_Count_2_.append(Antenna_Count_2[0])
                if Antenna_Count_2[0] != '':
                    Sum = Sum + int(Antenna_Count_2[0])
                else:
                    Sum = Sum + 0
        except:
            Antenna_Count_2_.append(Antenna_Count_2[0])
            if Antenna_Count_2[0] != '':
                Sum = Sum + int(Antenna_Count_2[0])
            else:
                Sum = Sum + 0

        try:
            if math.isnan(Antenna_Type_3[0]):
                Antenna_Count_3_.append('')
            else:
                Antenna_Count_3_.append(Antenna_Count_3[0])
                if Antenna_Count_3[0] != '':
                    Sum = Sum + int(Antenna_Count_3[0])
                else:
                    Sum = Sum + 0
        except:
            Antenna_Count_3_.append(Antenna_Count_3[0])
            if Antenna_Count_3[0] != '':
                Sum = Sum + int(Antenna_Count_3[0])
            else:
                Sum = Sum + 0

        try:
            if math.isnan(Antenna_Type_4[0]):
                Antenna_Count_4_.append('')
            else:
                Antenna_Count_4_.append(Antenna_Count_4[0])
                if Antenna_Count_4[0] != '':
                    Sum = Sum + int(Antenna_Count_4[0])
                else:
                    Sum = Sum + 0
        except:
            Antenna_Count_4_.append(Antenna_Count_4[0])
            if Antenna_Count_4[0] != '':
                Sum = Sum + int(Antenna_Count_4[0])
            else:
                Sum = Sum + 0

        # Antenna_Count_2_.append(Antenna_Count_2[0])
        # Antenna_Count_3_.append(Antenna_Count_3[0])
        # Antenna_Count_4_.append(Antenna_Count_4[0])
        if Sites_name[i] == 'AJM3384':
            pdb.set_trace()
        Total_Antenna.append(Sum)

    pdb.set_trace()
    def sequence_(Antenna_Type_1_, Antenna_Type_2_, Antenna_Type_3_, Antenna_Type_4_, Antenna_Count_1_, Antenna_Count_2_, Antenna_Count_3_, Antenna_Count_4_):
        Antenna_Type_1_ = ['' if i == '' or i != i else i for i in Antenna_Type_1_]
        Antenna_Type_2_ = ['' if i == '' or i != i else i for i in Antenna_Type_2_]
        Antenna_Type_3_ = ['' if i == '' or i != i else i for i in Antenna_Type_3_]
        Antenna_Type_4_ = ['' if i == '' or i != i else i for i in Antenna_Type_4_]


        Antena1 = []
        Antena1Count = []
        Antena2 = []
        Antena2Count = []
        Antena3 = []
        Antena3Count = []
        Antena4 = []
        Antena4Count = []

        # for i in range(len(Antenna_Type_1_)):
        #     Antena1.append('1')
        #     Antena2.append('1')
        #     Antena3.append('1')
        #     Antena4.append('1')


        for i in range(len(Antenna_Type_1_)):
            if Antenna_Type_1_[i] == '':
                if Antenna_Type_2_[i] != '':
                    Antena1.append(Antenna_Type_2_[i])
                    Antena1Count.append(Antenna_Count_2_[i])
                    Antena2.append('')
                    Antena2Count.append('')
                    Antena3.append('Na')
                    Antena3Count.append('Na')
                    Antena4.append('Na')
                    Antena4Count.append('Na')

                elif Antenna_Type_3_[i] != '':
                    Antena1.append(Antenna_Type_3_[i])
                    Antena1Count.append(Antenna_Count_3_[i])
                    Antena3.append('')
                    Antena3Count.append('')
                    Antena2.append('Na')
                    Antena2Count.append('Na')
                    Antena4.append('Na')
                    Antena4Count.append('Na')

                elif Antenna_Type_4_[i] != '':
                    Antena1.append(Antenna_Type_4_[i])
                    Antena1Count.append(Antenna_Count_4_[i])
                    Antena4.append('')
                    Antena4Count.append('')
                    Antena2.append('Na')
                    Antena2Count.append('Na')
                    Antena3.append('Na')
                    Antena3Count.append('Na')
                else:
                    Antena1.append(Antenna_Type_1_[i])
                    Antena1Count.append(Antenna_Count_1_[i])
                    Antena2.append('Na')
                    Antena2Count.append('Na')
                    Antena3.append('Na')
                    Antena3Count.append('Na')
                    Antena4.append('Na')
                    Antena4Count.append('Na')
            else:
                Antena1.append(Antenna_Type_1_[i])
                Antena1Count.append(Antenna_Count_1_[i])
                Antena2.append('Na')
                Antena2Count.append('Na')
                Antena3.append('Na')
                Antena3Count.append('Na')
                Antena4.append('Na')
                Antena4Count.append('Na')



        for i in range(len(Antenna_Type_1_)):
            if Antenna_Type_2_[i] == '' and Antena2[i] != '':

                if Antenna_Type_3_[i] != '' and Antena3[i] != '':
                # if Antenna_Type_3_[i] != '':
                    Antena2[i] = Antenna_Type_3_[i]
                    Antena2Count[i] = Antenna_Count_3_[i]
                    Antena3[i] =  ''
                    Antena3Count[i] =  ''

                elif Antenna_Type_4_[i] != '' and  Antena4[i] != '':
                    Antena2[i] = Antenna_Type_4_[i]
                    Antena2Count[i] = Antenna_Count_4_[i]
                    Antena4[i] =  ''
                    Antena4Count[i] =  ''
                else:
                    if Antena2[i] != '':
                        Antena2[i] = Antenna_Type_2_[i]
                        Antena2Count[i] = Antenna_Count_2_[i]

                    else:
                        Antena2[i] = ''
                        Antena2Count[i] = ''
            else:
                if Antena2[i] != '':
                    Antena2[i] = Antenna_Type_2_[i]
                    Antena2Count[i] = Antenna_Count_2_[i]

                else:
                    Antena2[i] = ''
                    Antena2Count[i] = ''




        for i in range(len(Antenna_Type_1_)):
            if Antenna_Type_3_[i] == ''  and Antena3[i] != '':

                if Antenna_Type_4_[i] != '' and Antena4[i] != '':
                # if Antenna_Type_4_[i] != '':
                    Antena3[i] = Antenna_Type_4_[i]
                    Antena3Count[i] = Antenna_Count_4_[i]
                    Antena4[i] =  ''
                    Antena4Count[i] =  ''
                else:

                    if Antena3[i] != '':
                        Antena3[i] = Antenna_Type_3_[i]
                        Antena3Count[i] = Antenna_Count_3_[i]

                    else:
                        Antena3[i] = ''
                        Antena3Count[i] = ''
                    #
                    # Antena3[i] = Antenna_Type_3_[i]
                    # Antena3Count[i] = Antenna_Count_3_[i]

            else:

                if Antena3[i] != '':
                    Antena3[i] = Antenna_Type_3_[i]
                    Antena3Count[i] = Antenna_Count_3_[i]

                else:
                    Antena3[i] = ''
                    Antena3Count[i] = ''

        for i in range(len(Antenna_Type_1_)):
            if Antena4[i] != '':
                Antena4[i] = Antenna_Type_4_[i]
                Antena4Count[i] = Antenna_Count_4_[i]
            else:
                Antena4[i] = ''
                Antena4Count[i] = ''

        Antena1 = ['' if i == 'Na' else i for i in Antena1]
        Antena1Count = ['' if i == 'Na' else i for i in Antena1Count]
        Antena2 = ['' if i == 'Na' else i for i in Antena2]
        Antena2Count = ['' if i == 'Na' else i for i in Antena2Count]
        Antena3 = ['' if i == 'Na' else i for i in Antena3]
        Antena3Count = ['' if i == 'Na' else i for i in Antena3Count]
        Antena4 = ['' if i == 'Na' else i for i in Antena4]
        Antena4Count = ['' if i == 'Na' else i for i in Antena4Count]

        return Antena1, Antena2, Antena3, Antena4, Antena1Count, Antena2Count, Antena3Count, Antena4Count


    # for i in range(len(Antenna_Type_1_)):
    #     if Antenna_Type_3_[i] == '' or Antenna_Type_3_[i] != Antenna_Type_3_[i]:
    #         Antena4.append('')
    #         Antena4Count.append('')
    #     else:
    #         Antena4.append(Antenna_Type_4_[i])
    #         Antena4Count.append(Antenna_Count_4_[i])

    Antena1, Antena2, Antena3, Antena4, Antena1Count, Antena2Count, Antena3Count, Antena4Count = sequence_(Antenna_Type_1_, Antenna_Type_2_, Antenna_Type_3_, Antenna_Type_4_, Antenna_Count_1_, Antenna_Count_2_, Antenna_Count_3_, Antenna_Count_4_)
    Antena1, Antena2, Antena3, Antena4, Antena1Count, Antena2Count, Antena3Count, Antena4Count = sequence_(Antena1, Antena2, Antena3, Antena4, Antena1Count, Antena2Count, Antena3Count, Antena4Count)
    Antena1, Antena2, Antena3, Antena4, Antena1Count, Antena2Count, Antena3Count, Antena4Count = sequence_(Antena1, Antena2, Antena3, Antena4, Antena1Count, Antena2Count, Antena3Count, Antena4Count)


    output_df = pandas.DataFrame()
    output_df['Site'] = Sites_name
    output_df['Antenna_Type1'] = Antena1
    output_df['Antenna_Type2'] = Antena2
    output_df['Antenna_Type3'] = Antena3
    output_df['Antenna_Type4'] = Antena4
    output_df['Antenna Count Type_1'] = Antena1Count
    output_df['Antenna Count Type_2'] = Antena2Count
    output_df['Antenna Count Type_3'] = Antena3Count
    output_df['Antenna Count Type_4'] = Antena4Count
    output_df['Total Current Antenna'] = Total_Antenna


    writer = pandas.ExcelWriter(Output_path + '\\' + 'Output data.xlsx', engine='xlsxwriter')
    output_df.to_excel(writer, sheet_name='Output', index=False)
    writer.save()
    writer.close()

    end = time.time()
    Execute_Time = "{:.3f}".format((end - start) / 60)
    print('The Execution Time of this Tool is %s minutes.' % Execute_Time)
    time.sleep(1)
    print('Execution Completed Succcessfully!')
    time.sleep(1)
    print('')
    print('')
    print('---------------Huawei RF Middle East----------------')
    print('---------For Support: Danish Ali(dwx854280)---------')
    print('---------------Contact: 00971508552942--------------')
    time.sleep(3)
