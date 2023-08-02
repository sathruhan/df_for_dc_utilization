import pandas as pd
import numpy as np
import yaml
import os
import openpyxl

path = r'D:\Python\pythonProject\host\files'
file_path_dict = {}
dataframe_dict = {}


def get_data():
    file_path = r'report_name.yml'
    with open(file_path, 'r') as file:
        data = yaml.safe_load(file)
    return data.get('reports', [])


def create_var(report_list):
    global file_path_dict
    for report in report_list:
        var_name_file_path = report.replace('-', '_')
        file_path = os.path.join(path, report) + '.xlsx'
        file_path_dict[var_name_file_path] = file_path
        # print(var_name_file_path)


def create_df():
    global file_path_dict, dataframe_dict

    for file_name, file_path in file_path_dict.items():
        # print(file_name)

        # Check if 'inventory' word appears in the file name
        if 'Inventory_Report' in file_name:

            # Read the data into DataFrame, skipping the header_row if necessary
            dataframe_dict[file_name] = pd.read_excel(file_path, skiprows=1)
        else:

            # Read the data into DataFrame, skipping the header_row if necessary
            dataframe_dict[file_name] = pd.read_excel(file_path)


def create_occupancy_df():
    global dataframe_dict

    # egdc cbm (separating HGKPL)

    df_egdc = dataframe_dict['Client_Level_Inventory_Summary_Report_EGDC']
    egdc_cbm = df_egdc['Cbm'].sum()
    hgkpl_cbm = df_egdc.loc[df_egdc['Client Code'] == 'HGKPL', 'Cbm'].values[0]
    egdc_cbm = egdc_cbm - hgkpl_cbm

    # eskd cbm (remove the ESLP)

    df_eskd = dataframe_dict['Client_Level_Inventory_Summary_Report_ESKD']
    eskd_cbm = df_eskd['Cbm'].sum()
    eslp_cbm = df_eskd.loc[df_eskd['Client Code'] == 'ESLP', 'Cbm'].values[0]
    eskd_cbm = eskd_cbm - eslp_cbm

    # lppl (separating EMARPH) and calculating the cold room from inventory reports

    df_lppl = dataframe_dict['Client_Level_Inventory_Summary_Report_LPPL']
    lppl_cbm = df_lppl['Cbm'].sum()
    emarph_cbm = df_lppl.loc[df_lppl['Client Code'] == 'EMARPH', 'Cbm'].values[0]

    df_ars = dataframe_dict['Inventory_Report_ARS']
    df_ars = df_ars.loc[df_ars['Client So'] == '2-8']
    ars_cbm = df_ars['Cbm'].sum()
    if ars_cbm < 2:  # safety margin is 2 cbm
        ars_cbm = 2

    df_drdns = dataframe_dict['Inventory_Report_Durdans']
    # print(df_drdns.columns.tolist())
    # df_drdns = df_drdns['Client So'].astype(str)
    # print(df_drdns['Client So'], df_drdns.dtypes)
    df_drdns = df_drdns.loc[df_drdns['Client So'] == '2-8']
    drdns_cbm = df_drdns['Cbm'].sum()

    df_hms = dataframe_dict['Inventory_Report_HMS']
    # print(df_hms.columns.tolist())
    df_hms = df_hms.loc[df_hms['Client So'] == '2-8']
    hms_cbm = df_hms['Cbm'].sum()
    if hms_cbm < 10:  # safety margin is 10 cbm
        hms_cbm = 10

    cold_room_cbm = emarph_cbm + ars_cbm + drdns_cbm + hms_cbm

    df_soft = dataframe_dict['Inventory_Report_SOFT']
    soft_cbm = df_soft['Cbm'].sum()

    lppl_cbm = lppl_cbm - cold_room_cbm + soft_cbm

    # nuge

    df_nuge = dataframe_dict['Client_Level_Inventory_Summary_Report_NUGE']
    nuge_cbm = df_nuge['Cbm'].sum()

    nestle_plt = int(input('Enter the Nestle Main DC total bins : '))
    bipl_in_slip_sheet = int(input('Enter the Nestle BIPL pallets in slip sheet : '))
    bipl_in_rack = int(input('Enter the Nestle BIPL in rack : '))
    bipl_on_woodn = int(input('Enter the Nestle BIPL on woodn : '))

    reserved = 931  # reserved for quality purpose
    main_dc_plt = nestle_plt + reserved
    bipl_plt = bipl_on_woodn + bipl_in_rack + bipl_in_slip_sheet

    data = {'Client_Name': ['Expo Global Freeport', 'Kandana DC', 'Orugodawatta DC', 'Orugodawatta DC Cold Room',
                            'Peliyagoda DC', 'Hirdaramani Woven', 'Nestle - Main DC', 'BIPL'],
            'Capacity': [22000, 1100, 10300, 100, 11750, 4300, 10131, 1500],
            'Fluctuation Rate': [0.03, 0.03, 0.07, 0.03, 0.03, 0.03, 0, 0],
            'Revenue Rate': [3254, 3002, 2920, 2920, 2492, 0, 0, 0],
            'Net Profit Rate': [736, 707, 533, 533, 620, 0, 0, 0],
            'Occupied CBM': [egdc_cbm, eskd_cbm, lppl_cbm, cold_room_cbm, nuge_cbm, hgkpl_cbm, main_dc_plt, bipl_plt]
            }
    df_occupancy = pd.DataFrame(data)

    # Utilization %
    df_occupancy['Utilization %'] = (df_occupancy['Occupied CBM'] / df_occupancy['Capacity']) * 100
    df_occupancy['Utilization %'] = df_occupancy['Utilization %'].apply(lambda x: min(x, 100))

    # Unoccupied CBM
    df_occupancy['Unoccupied CBM'] = df_occupancy['Capacity'] - df_occupancy['Occupied CBM']
    df_occupancy['Unoccupied CBM'] = df_occupancy['Unoccupied CBM'].apply(lambda x: max(x, 0))

    # For demand fluctuation + Honeycomb
    df_occupancy['For demand fluctuation + Honeycomb'] = df_occupancy['Capacity']*df_occupancy['Fluctuation Rate']

    # Actual Sellable CBM
    df_occupancy['Actual Sellable CBM'] = df_occupancy['Unoccupied CBM'] - df_occupancy[
        'For demand fluctuation + Honeycomb']
    df_occupancy['Actual Sellable CBM'] = df_occupancy['Actual Sellable CBM'].apply(lambda x: max(x, 0))

    # Occupied CBM with Honeycomb
    df_occupancy['Occupied CBM with Honeycomb'] = df_occupancy['Capacity'] - df_occupancy['Actual Sellable CBM']

    # revenue
    df_occupancy['Revenue'] = df_occupancy['Actual Sellable CBM'] * df_occupancy['Revenue Rate']/30.5

    # net profit
    df_occupancy['Net Profit'] = df_occupancy['Actual Sellable CBM'] * df_occupancy['Net Profit Rate']/30.5

    print(df_occupancy)


if __name__ == "__main__":
    report_list = get_data()
    create_var(report_list)
    create_df()
    create_occupancy_df()
