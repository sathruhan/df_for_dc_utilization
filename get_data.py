import pandas as pd
import numpy as np
import yaml
import os
import openpyxl

path = r'D:\Python\pythonProject\host\files'


def get_data():
    file_path = r'report_name.yml'
    with open(file_path, 'r') as file:
        data = yaml.safe_load(file)
    return data.get('reports', [])


def create_var(report_list):
    file_path_dict = {}
    for report in report_list:
        var_name_file_path = report.replace('-', '_')
        file_path = os.path.join(path, report) + '.xlsx'
        file_path_dict[var_name_file_path] = file_path
        print(var_name_file_path)
    return file_path_dict


def create_df(var_name_file_path):
    data_frames_dict = {}
    for var_name, dataset in var_name_file_path.items():
        df = pd.read_excel(dataset)
        data_frames_dict[var_name] = df
    return data_frames_dict


if __name__ == "__main__":
    report_list = get_data()
    file_path_dict = create_var(report_list)
    # print(file_path_dict['Client_Wise_Stock_Ageing_Report_EGDC'])
    # create_df(report_list, file_path_dict)
    data_frames_dict = create_df(file_path_dict)

    df_cbm_

