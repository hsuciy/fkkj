# -*- coding: utf-8 -*-
# @Time : 2021/9/22 14:53
# @Author : Xiaojun Liu
# @Project : wnsd_infer
import yaml
import csv
import datetime
import os
import numpy as np
import pandas as pd
import logging
from logging import handlers
import openpyxl
import json
import sys

sys.setrecursionlimit(50000)


class Logger:
    level_relation = {
        'debug': logging.DEBUG,
        'info': logging.INFO,
        'warning': logging.WARNING,
        'error': logging.ERROR,
        'critcal': logging.CRITICAL
    }

    def __init__(self, filename, level='info', when='D', backupCount=3, fmt='%(asctime)s -%(pathname)s '
                                                                            '[line:%(lineno)d]-%(levelname)s: %(message)s'):
        # 往屏幕上输入
        self.logger = logging.getLogger(filename)
        format_str = logging.Formatter(fmt)
        self.logger.setLevel(self.level_relation.get(level))
        sh = logging.StreamHandler()
        sh.setFormatter(format_str)
        th = handlers.TimedRotatingFileHandler(filename=filename, when=when, backupCount=backupCount, encoding='utf8')
        th.setFormatter(format_str)
        self.logger.addHandler(th)
        self.logger.addHandler(sh)


# 生成lms文件
def create_lms_file(data):
    lms_info_list = list()
    create_column = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
    for row in data.itertuples():
        device_name, output = row[3], row[7]
        columns_info_list = output.split(',')
        temp_list = list()
        for item in columns_info_list:
            if 'R[' in item:
                res = item.split('R[')[1].strip(']')
                temp_list.append(res)
        for num in temp_list:
            new_line = [num, device_name, 'unit', 0, -1.0, -1.0, -1.0, -1.0, -1.0, 1.0, 90]
            lms_info_list.append(new_line)
    data_res = pd.DataFrame(lms_info_list, columns=create_column)
    writer_file = r'C:\Mnchao\py\TEG_Catalog.xlsx'
    writer = pd.ExcelWriter(writer_file, engine='openpyxl', mode='a')
    data_res.to_excel(writer, sheet_name='lsm', index=False)
    writer.save()
    lms_path = r'C:\Mnchao\py\lms_model.lms'
    with open(lms_path, 'a+', encoding='utf8') as fr:
        for line in lms_info_list:
            # info = ' '.join(line)
            # 写到文件里面，必须都为str
            info = [str(i) for i in line]
            info_res = ' '.join(info)
            fr.write(info_res + '\n')

# 读取文件信息,并写到tst sheet
def read_excel(excel_path, log_info):
    # 读取文件所有的sheet(需要注意人打开excel修改后会被加密，需要解密才能读取)
    result = pd.read_excel(excel_path, engine='openpyxl', sheet_name=None)
    # 获取不同的列名
    week_date_list = list()
    month_date_list = list()
    num = 0
    # 遍历每个sheet
    for i in result.keys():
        num += 1
        if num == 1:
            # data = pd.DataFrame(pd.read_excel(excel_path, engine='openpyxl', sheet_name=i, keep_default_na=False))
            data = pd.DataFrame(pd.read_excel(excel_path, engine='openpyxl', sheet_name=i))
            columns_name = data.columns.tolist()
            # 生成excel的列明
            create_column = ['Mark', 'Module_Name', 'Device_Name', 'Algorithm', 'Input', 'Pad_Number', 'Output', 'Comment']
            # 这两个是需要的值
            psn = int(input("input整数序:"))
            Vdd = input("input Vdd:")
            information_list = []
            symbol = 0
            Ptype = 0
            tst_temp_list = list()
            lms_tmp_list = list()
            # 遍历每个sheet的所有行
            for row in data.itertuples():
                try:
                    series_value = pd.Series(row[1:], index=columns_name)
                    series_value.fillna('None', inplace=True)
                    Name_In = ''
                    Pad_In = ''
                    TK = series_value['K01']
                    Mos_Tr = series_value['MOS_Tr']
                    Type = series_value['TYPE']
                    DNW = series_value['DNW']
                    Tlayer = series_value['Test Layer']
                    KDev = series_value['Key Device']
                    Dsc = series_value['Description']
                    D = series_value['D']
                    G = series_value['G']
                    S = series_value['S']
                    B = series_value['B']
                    Drw_L = series_value['L']
                    Drw_W = series_value['W']
                    Rel_L = series_value['L(real)']
                    Rel_W = series_value['W(real)']
                    if Rel_L !='None' and Rel_W != 'None' :
                        Ral_L = '{}'.format(round(Rel_L, 3))
                        Ral_W = '{}'.format(round(Rel_W, 3))
                        Rll_L = str(Ral_L).replace(".", "")
                        Rll_W = str(Ral_W).replace(".", "")
                    if Mos_Tr !='None'and Rll_W !='None'and Rll_L!='None' and TK!='None':
                        Name_In = Mos_Tr + '_' + Type + '_' + DNW + '_W' + Rll_W + '_L' + Rll_L + '_' + TK
                    if D !='None' and G !='None'and S !='None'and B !='None':
                        Pad_In = 'D={},G={},S={},B={}'.format(D, G, S, B)
                    if Mos_Tr.startswith('N'):
                        symbol = "1"
                        Ptype = "2"
                    else:
                        symbol = "-1"
                        Ptype = "1"
                    res_dict = dict()
                    if Mos_Tr == "NTN":
                        res_dict = {"VTL": ['', TK, 'VTL_' + Name_In, 'Vt_fc_rapid',
                                             'D_p="' + str(D) + '",G_p="' + str(G) + '",S_p="' + str(S) + '",B1p="' + str(
                                                 B) + '",Vd=0.05*(' + str(symbol) + '),Vg_stop=' + Vdd + '*(' + str(
                                                 symbol) + '),Ic=40n*(' + str(symbol) + '*' + str(Ptype) + '),Igc=0.01*(' + str(
                                                 symbol) + '),W=' + Ral_W + ',L=' + Ral_L + ',Op="1|0",Dvt=-0.2*(' + str(
                                                 symbol) + ')', '',
                                             "R[" + str(psn) + "]," + "RA" + str(psn) + "," + "RB" + str(
                                                 psn) + "," + "RC" + str(
                                                 psn) + "," + "RD" + str(psn), ''],
                                     "VTS": ['', TK, 'VTS_' + Name_In, 'Vt_fc_rapid',
                                             'D_p="' + str(D) + ',G_p="' + str(G) + '",S_p="' + str(S) + '",B1p="' + str(
                                                 B) + '",Vd=' + Vdd + '*(' + str(symbol) + '),Vg_stop=' + Vdd + '*(' + str(
                                                 symbol) + '),Ic=80n*(' + str(symbol) + '),Igc=0.01*(' + str(
                                                 symbol) + '),W=' + Ral_W + ',L=' + Ral_L + ',Op="1|0",Dvt=-0.2*(' + str(
                                                 symbol) + ')', Pad_In,
                                             "R[" + str(psn + 1) + "]," + "RA" + str(psn + 1) + "," + "RB" + str(
                                                 psn + 1) + "," + "RC" + str(
                                                 psn + 1) + "," + "RD" + str(psn + 1), ''],
                                     "DIBL": ['', TK, 'DIBL_' + Name_In, 'ITO1',
                                              'In=(`R[' + str(psn) + ']`-`R[' + str(psn + 1) + ']`)/(' + Vdd + '-0.05)', '',
                                              "R[" + str(psn + 2) + "]", ''],
                                     "SWING": ['', TK, 'Swing_' + Name_In, 'Slopsweep_new',
                                               'Hi_p="' + str(G) + '",Start=-0.5*(' + str(symbol) + '),Stop=1.5*(' + str(
                                                   symbol) + '),Step=0.05,Lo_p="' + str(D) + '",Lo_v=0,B1p="' + str(
                                                   S) + '",B1v=0.05*(' + str(
                                                   symbol) + '),B2p="' + str(B) + '",B2v=-0.7*(' + str(
                                                   symbol) + '),Id1=-1E-9*' + Ral_W + '*(' + str(
                                                   symbol) + '),Id2=-1E-8*' + Ral_W + '*(' + str(
                                                   symbol) + '),Mode=4', '', "R[" + str(psn + 3) + "]", ''],
                                     "VTGM": ['', TK, 'VTGM_' + Name_In, 'Vt_gms_quick',
                                              'Hi_p="' + str(D) + '",Lo_p="' + str(G) + '",B1p="' + str(S) + '",B2p="' + str(
                                                  B) + '",Hi_v=0.05*(' + str(
                                                  symbol) + '),Lo_v=0,B1v=0,B2v=0,Vgstart=0,Vgstop=2*(' + str(
                                                  symbol) + '),Co=0.05,Vstep=0.05,R=1E-4,B2c=0.05*(' + str(symbol) + ')', '',
                                              "R[" + str(psn + 4) + "]," + "RA" + str(psn + 4), ''],
                                     "VTGMVB": ['', TK, 'VTGM_07G_' + Name_In, 'Vt_gms_quick',
                                                'Hi_p="' + str(D) + '",Lo_p="' + str(G) + '",B1p="' + str(S) + '",B2p="' + str(
                                                    B) + '",Hi_v=0.05*(' + str(symbol) + '),Lo_v=0,B1v=0,B2v=-0.7*(' + str(
                                                    symbol) + '),Vgstart=0,Vgstop=2*(' + str(
                                                    symbol) + '),Co=0.05,Vstep=0.05,R=1E-4,B2c=0.05*(' + str(symbol) + ')', '',
                                                "R[" + str(psn + 5) + "]," + "RA" + str(psn + 5), ''],
                                     "BodyEff": ['', TK, 'Gamma_' + Name_In, 'ITO1',
                                                 'In=(`R[' + str(psn + 5) + ']`-`R[' + str(psn + 4) + ']`)*1000', Pad_In,
                                                 "R[" + str(psn + 6) + "]",
                                                 ''],
                                     "IDL": ['', TK, 'IDL_' + Name_In, 'Current',
                                             'Hi_p="' + str(D) + '",Lo_p="' + str(G) + '",B1p="' + str(S) + '",B2p="' + str(
                                                 B) + '",W=' + Ral_W + ',Hi_v=0.05*(' + str(symbol) + '),Lo_v=1.2*(' + str(
                                                 symbol) + '),In=2,Unit=1E-6*(' + str(symbol) + '),B1c=0.05*(' + str(
                                                 symbol) + '),B2c=0.05 *(' + str(symbol) + '),R=0,Co=0.05,W=' + Ral_W, '',
                                             "R[" + str(psn + 7) + "]", ''],
                                     "IDS": ['', TK, 'IDS_' + Name_In, 'Current',
                                             'Hi_p="' + str(D) + '",Lo_p="' + str(G) + '",B1p="' + str(S) + '",B2p="' + str(
                                                 B) + '",W=' + Ral_W + ',Hi_v=' + Vdd + '*(' + str(
                                                 symbol) + '),Lo_v=' + Vdd + '*(' + str(
                                                 symbol) + '),In=2,Unit=1E-6*(' + str(symbol) + '),B1c=0.05*(' + str(
                                                 symbol) + '),B2c=0.05*(' + str(symbol) + '),R=0,Co=0.05', '',
                                             "R[" + str(psn + 8) + "]", ''],
                                     "IOF": ['', TK, 'IOF_' + Name_In, 'Current_Multi',
                                             'Hi_p="' + str(D) + '",Lo_p="' + str(G) + '",B1p="' + str(S) + '",B2p="' + str(
                                                 B) + '",W=' + Ral_W + ',Hi_v=' + Vdd + '*(' + str(
                                                 symbol) + '),Lo_v=0,In=2,Unit=1E-12,B1c=0.05,B2c=0.05,R=0,Co=0.05', '',
                                             "R[" + str(psn + 9) + "]," + "R[" + str(psn + 10) + "]," + "R[" + str(
                                                 psn + 11) + "]," + "R[" + str(psn + 12) + "]", ''],
                                     "IGIDL": ['', TK, 'IGIDL_' + Name_In, 'Current',
                                               'Hi_p="' + str(D) + '",Lo_p="' + str(G) + '",Hi_v=' + Vdd + '*(' + str(
                                                   symbol) + '),Lo_v=-0.25*(' + str(
                                                   symbol) + '),In=2,W=' + Ral_W + ',Unit=1.E-12,B1p="' + str(
                                                   B) + '",B1c=0.05,B2c=0.05,R=0,Co=0.05,B1v=-0.7*(' + str(symbol) + ')', '',
                                               "R[" + str(psn + 13) + "]", ''],
                                     "ISUB": ['', TK, 'ISUB_' + Name_In, 'Current',
                                              'Hi_p="' + str(B) + '",Lo_p="' + str(G) + '",B1p="' + str(S) + '",B2p="' + str(
                                                  D) + '",W=' + Ral_W + ',Hi_v=0,Lo_v=0.1*(' + str(
                                                  symbol) + '),Vgstop=1.5*(' + str(
                                                  symbol) + '),B2v=1.575*(' + str(
                                                  symbol) + '),In=2,Unit=1E-12,B1c=0.05,B2c=0.05,R=0,Co=0.05,St=3', '',
                                              "R[" + str(psn + 14) + "]", ''],
                                     "IB": ['', TK, 'IB_' + Name_In, 'Current',
                                            'Hi_p="' + str(B) + '",Lo_p="' + str(G) + '",B1p="' + str(S) + '",B2p="' + str(
                                                D) + '",W=' + Ral_W + ',Hi_v=0,Lo_v=' + Vdd + '*(' + str(
                                                symbol) + '),In=2,Unit=1E-6,B2v=' + Vdd + '*(' + str(
                                                symbol) + '),B1c=0.05,B2c=0.05,R=0,Co=0.05', '', "R[" + str(psn + 15) + "]",
                                            ''],
                                     "IG": ['', TK, 'IG_' + Name_In, 'Current',
                                            'Hi_p="' + str(G) + '",Lo_p="' + str(D) + '",B1p="' + str(S) + '",B2p="' + str(
                                                B) + '",W=' + Ral_W + ',Hi_v=' + Vdd + ',Lo_v=' + Vdd + ',In=2,Unit=1E-6,B1c=0.05,B2c=0.05,R=0,Co=0.05',
                                            '', "R[" + str(psn + 16) + "]", ''],
                                     "VBD": ['', TK, 'VBD_' + Name_In, 'Vbd_sweep',
                                             'Hi_p="' + str(D) + '",Lo_p="' + str(G) + '|' + str(S) + '|' + str(
                                                 B) + '",Stop=10*(' + str(
                                                 symbol) + '),Step=0.1,Ic=100n*(' + str(
                                                 symbol) + '),R=100n,Co=0.001,Skip=5,Num=4', '',
                                             "R[" + str(psn + 17) + "]", '']
                                     }
                        psn += 18
                    if Mos_Tr == "NTNLV":
                        res_dict = {
                            "Cap_tox": ['', TK, 'Cap_' + Name_In, 'Captox_scmu_call',
                                        'Hi_p="3",Lo_p="2",Hi_v=0.6,Vc=0.05,Frq=1k,R=1e-9,In=3,Area=1,Copen=0,Unit=1E18', '',
                                        "R[" + str(psn) + "]," + "R[" + str(psn + 1) + "]", ''],
                            "LKG": ['', TK, 'LKG_' + Name_In, 'FVSI', 'FV=0.6', Pad_In,
                                    "R[" + str(psn) + "]," + "R[" + str(psn + 1) + "]", ''],
                            "BKV": ['', TK, 'BKV_' + Name_In, 'VBD', 'VG=VS=VB=0,VDmin=0,Vdmax=10,IDT=0.1uA*' + Ral_W, Pad_In,
                                    "R[" + str(psn) + "]," + "R[" + str(psn + 1) + "]", '']
                        }
                        psn += 2
                    for value in res_dict.values():
                        tst_temp_list.append(value)
                except Exception as e:
                    log_info.logger.info(e)
                    print(e)
            data_res = pd.DataFrame(tst_temp_list, columns=create_column)
            create_lms_file(data_res)
            new_excel_path = r'C:\Mnchao\py\TEG_Catalog2.xlsx'
            writer_file = r'C:\Mnchao\py\TEG_Catalog.xlsx'
            writer = pd.ExcelWriter(writer_file,engine='openpyxl', mode='a')
            data_res.to_excel(writer, sheet_name='ntst', index=False)
            writer.save()
            # 将信息写道tsf文件中
            tsf_path = r'C:\Mnchao\py\tst_model.tsf'
            with open(tsf_path, 'a+', encoding='utf8') as fr:
                first_line = '=	Test	dummy	1 dummy	05/11/2022	10:07:04	specs			description\n'
                fr.write(first_line)
                second_line = '>	Module Name	Device Name	Algorithm	Input	Pad Number	Output	Limit		Comment \n'
                fr.write(second_line)
                third_line = '# \n'
                fr.write(second_line)
                for row in data_res.itertuples():
                    #module_name, Device_name, Algorithm, Input, Pad_number, output, limit, comment =
                    res_name = row[2:]
                    new_line = ' '.join(res_name)
                    fr.write(new_line + '\n')
    return True


def folder_handle():
    excel_path = r'C:\Mnchao\py\TEG_Catalog.xlsx'
    base_excel_path = os.path.dirname(excel_path)
    # 创建日志记录异常文件夹
    log_folder = base_excel_path + '/' + 'log'
    if not os.path.exists(log_folder):
        os.makedirs(log_folder)
    now_date = datetime.datetime.now().strftime("%Y_%m_%d")
    log_file_path = log_folder + '/' + now_date + '.log'
    log_info = Logger(log_file_path)
    # for file in os.listdir(excel_path):
    #     file_path = excel_path + '/' + file
    if os.path.isfile(excel_path) and '.xlsx' in excel_path:
        read_excel(excel_path, log_info)
        complete_run_time = datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
        log_info.logger.info('{} complete the task: {}'.format(excel_path, complete_run_time))


if __name__ == '__main__':
    folder_handle()
