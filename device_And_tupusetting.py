import pandas as pd

from PlatformTable import output_template_all


def device_info(file1, file2, file4):
    df1_1 = pd.read_excel(file1, dtype=str, sheet_name="设备档案")
    df1_2 = pd.read_excel(file1, dtype=str, sheet_name="输出模板")
    df2 = pd.read_excel(file2, dtype=str)

    for index, series in df1_2.iterrows():
        if str(series['数据项（特征）编码'])[-3:] == '000':
            df1_2.at[index, '测点（通道）类型'] = '温度'
            df1_2.at[index, '数据项（特征）名称'] = '轴承温度'
            df1_2.at[index, '数据项（特征）类型'] = '温度'
            df1_2.at[index, '数据类型'] = '模拟量'
            df1_2.at[index, '单位'] = '℃'

    rows_to_be_deleted = []
    for index, series in df1_2.iterrows():
        if (str(series['测点（点位）编码'])[-2:-1] in ['X', 'Y', 'Z'] and
                str(series['数据项（特征）编码'])[-3:] not in ['001', '003', '004', '005', '008'] and
                series['测点（通道）类型'] == '加速度'):
            rows_to_be_deleted.append(index)

    df1_2 = df1_2.drop(rows_to_be_deleted)  # 删除无线传感器的多余特征项

    # 建立device_info
    df4 = df1_2
    equipcod_list = []
    pointcod_list = []
    datatype_list = []
    routecod_list = []
    for index, series in df4.iterrows():
        if series['测点（通道）类型'] in ['加速度', '温度']:
            equipcod_list.append(series['设备编码'])
            pointcod_list.append(series['测点（点位）编码'])
            datatype_list.append(series['数据项（特征）编码'])
            routecod_list.append(series['通道编码'])

    df_deviceinfo = pd.DataFrame(
        columns=['区域', '设备编码', '测点编号', '数据项编码', 'MAC地址', '通道编号', '通道类型', '通道值']
    )
    df_deviceinfo['设备编码'] = equipcod_list
    df_deviceinfo['测点编号'] = pointcod_list
    df_deviceinfo['数据项编码'] = datatype_list
    df_deviceinfo['通道编号'] = routecod_list  # 创建dataframe对象，填入表头和四列数据

    typecod_keylist = df2['数据项代号'].to_list()
    routekey_valuelist = df2["数据项（特征）类型"].to_list()
    dict1 = dict(zip(typecod_keylist, routekey_valuelist))  # 创建数据项编码后三位编号和通道值的对应字典

    equipcod_keylist = df1_1['*设备编码'].to_list()
    area_valuelist = df1_1["* 所属区域"].to_list()
    dict2 = dict(zip(equipcod_keylist, area_valuelist))

    for index, series in df_deviceinfo.iterrows():
        equip_cod = series['设备编码']
        if equip_cod in dict2:
            df_deviceinfo.at[index, '区域'] = dict2[equip_cod]
        else:
            df_deviceinfo.at[index, '区域'] = 'Unknown'  # 可以设置一个默认值，以便调试

        if series['测点编号'][-2:-1] not in ['X', 'Y', 'Z']:
            mac = str(series['通道编号'])[:-3]
            df_deviceinfo.at[index, 'MAC地址'] = mac
        else:
            mac2 = str(series['通道编号'])[:-5]
            df_deviceinfo.at[index, 'MAC地址'] = mac2  # 填写无线传感器的’MAC地址‘，实则为网关的sn号

        if series['测点编号'][-1:] == 'A':
            df_deviceinfo.at[index, '通道类型'] = 'ACCELEROMETER'
        else:
            df_deviceinfo.at[index, '通道类型'] = 'TEMPERATURE'

        data_type_cod3 = series['数据项编码'][-3:]
        if series['测点编号'][-2] in ['X', 'Y', 'Z']:
            if data_type_cod3 == '001':
                df_deviceinfo.at[index, '通道值'] = 'integratRMS'
            elif data_type_cod3 == '003':
                df_deviceinfo.at[index, '通道值'] = 'rmsValues'
            elif data_type_cod3 == '004':
                df_deviceinfo.at[index, '通道值'] = 'diagnosisPk'
            elif data_type_cod3 == '005':
                df_deviceinfo.at[index, '通道值'] = 'envelopEnergy'
            elif data_type_cod3 == '008':
                df_deviceinfo.at[index, '通道值'] = 'integratPk'
            elif data_type_cod3 == '000':
                df_deviceinfo.at[index, '通道值'] = 'TemperatureBot'
        else:
            df_deviceinfo.at[index, '通道值'] = dict1.get(data_type_cod3, 'Unknown')  # 使用dict1查找通道值，默认Unknown

    df_deviceinfo.to_excel(file4, sheet_name='Sheet1', index=False)  # 保存为device-info_new.xlsx
    return df_deviceinfo


def tupuSetting(file1, file4):
    df1_2 = pd.read_excel(file1, dtype=str, sheet_name="输出模板")

    df_tupu = pd.DataFrame(columns=['设备名称', '设备编码', '测点（点位）名称', '测点（点位）编码', '测点（通道）类型',
                                    '波形数据名称', '波形数据编码', '波形数据类型', '数据类型', '单位',
                                    '抽样频率（Hz）', '采样时长(s)', '高通滤波（Hz）', '分析截止频率（Hz）',
                                    '采样点数（需求）'])
    pointcod_list_only = []
    pointname_list_only = []
    equipcod_list_only = []
    equipname_list_only = []
    pointcod_list_only_wireless = []
    pointname_list_only_wireless = []
    equipcod_list_only_wireless = []
    equipname_list_only_wireless = []
    for index, series in df1_2.iterrows():
        if series['测点（通道）类型'] == '加速度' and series['测点（点位）编码'] not in pointcod_list_only \
                and series['测点（点位）编码'] not in pointcod_list_only_wireless:
            if series['测点（点位）编码'][-2:-1] not in ['X', 'Y', 'Z']:
                pointcod_list_only.append(series['测点（点位）编码'])
                pointname_list_only.append(series['测点（点位）名称'])
                equipcod_list_only.append(series['设备编码'])
                equipname_list_only.append(series['设备名称'])
            else:
                pointcod_list_only_wireless.append(series['测点（点位）编码'])
                pointname_list_only_wireless.append(series['测点（点位）名称'])
                equipcod_list_only_wireless.append(series['设备编码'])
                equipname_list_only_wireless.append(series['设备名称'])
        else:
            continue

    pointcod_list_onlyrepeated = [x1 for x1 in pointcod_list_only for _ in range(9)]
    pointname_list_onlyrepeated = [x2 for x2 in pointname_list_only for _ in range(9)]
    equipcod_list_onlyrepeated = [x3 for x3 in equipcod_list_only for _ in range(9)]
    equipname_list_onlyrepeated = [x4 for x4 in equipname_list_only for _ in range(9)]
    df_tupu['测点（点位）编码'] = pointcod_list_onlyrepeated + pointcod_list_only_wireless
    df_tupu['测点（点位）名称'] = pointname_list_onlyrepeated + pointname_list_only_wireless
    df_tupu['设备编码'] = equipcod_list_onlyrepeated + equipcod_list_only_wireless
    df_tupu['设备名称'] = equipname_list_onlyrepeated + equipname_list_only_wireless

    repeat_times = len(pointcod_list_onlyrepeated) // 9
    repeat_times_wireless = len(pointcod_list_only_wireless)

    data1 = ['加速度', '加速度', '加速度', '加速度', '加速度', '加速度', '速度', '速度', '速度']
    data1_repeated = data1 * repeat_times + ['加速度'] * repeat_times_wireless
    df_tupu['测点（通道）类型'] = data1_repeated

    data2 = ['20K高频加速度波形', '40K高频加速度波形', '80K高频加速度波形', '16K低频加速度波形', '32K低频加速度波形',
             '64K低频加速度波形', '8K速度波形', '16K速度波形', '32K速度波形']
    data2_repeated = data2 * repeat_times + ['12.8K高频加速度波形'] * repeat_times_wireless
    df_tupu['波形数据名称'] = data2_repeated

    data3 = ['GP08', 'GP16', 'GP32', 'DP32', 'DP64', 'DP128', 'SD32', 'SD64', 'SD128']
    data3_repeated = data3 * repeat_times + ['YS'] * repeat_times_wireless
    df_tupu['波形数据编码'] = data3_repeated

    data4 = ['高频加速度波形(0.1~10KHz)', '高频加速度波形(0.1~10KHz)', '高频加速度波形(0.1~10KHz)',
             '低频加速度波形(0.1~2KHz)', '低频加速度波形(0.1~2KHz)', '低频加速度波形(0.1~2KHz)',
             '速度波形(0.1~1000Hz)', '速度波形(0.1~1000Hz)', '速度波形(0.1~1000Hz)']
    data4_repeated = data4 * repeat_times + ['无线加速度波形'] * repeat_times_wireless
    df_tupu['波形数据类型'] = data4_repeated

    data5 = ['加速度波形', '加速度波形', '加速度波形', '加速度波形', '加速度波形', '加速度波形', '速度波形', '速度波形',
             '速度波形']
    data5_repeated = data5 * repeat_times + ['加速度波形'] * repeat_times_wireless
    df_tupu['数据类型'] = data5_repeated

    data6 = ['m/s2', 'm/s2', 'm/s2', 'm/s2', 'm/s2', 'm/s2', 'mm/s', 'mm/s', 'mm/s']
    data6_repeated = data6 * repeat_times + ['m/s2'] * repeat_times_wireless
    df_tupu['单位'] = data6_repeated

    data7 = ['32768', '32768', '32768', '8192', '8192', '8192', '8192', '8192', '8192']
    data7_repeated = data7 * repeat_times + ['12800'] * repeat_times_wireless
    df_tupu['抽样频率（Hz）'] = data7_repeated

    data8 = ['0.8', '1.6', '3.2', '3.2', '6.4', '12.8', '3.2', '6.4', '12.8']
    data8_repeated = data8 * repeat_times + ['1'] * repeat_times_wireless
    df_tupu['采样时长(s)'] = data8_repeated

    data9 = ['0.1', '0.1', '0.1', '0.1', '0.1', '0.1', '0.1', '0.1', '0.1']
    data9_repeated = data9 * repeat_times + ['3'] * repeat_times_wireless
    df_tupu['高通滤波（Hz）'] = data9_repeated

    data10 = ['10000', '10000', '10000', '2000', '2000', '2000', '1000', '1000', '1000']
    data10_repeated = data10 * repeat_times + ['5000'] * repeat_times_wireless
    df_tupu['分析截止频率（Hz）'] = data10_repeated

    data11 = ['20480', '40960', '81920', '16384', '32768', '65536', '8192', '16384', '32768']
    data11_repeated = data11 * repeat_times + [''] * repeat_times_wireless
    df_tupu['采样点数（需求）'] = data11_repeated

    df_tupu.to_excel(file4, sheet_name='Sheet1', index=False)


if __name__ == "__main__":
    # output_template_all("test/excel/data_all（测试）.xlsx", "后台文件/my_def_对应注释.xlsx", "/",
    #                     "test/excel/平台导入表.xlsx", )
    device_info("test/excel/平台导入表.xlsx", "后台文件/my_def_对应注释.xlsx", "test/device.xlsx")
    tupuSetting("test/excel/平台导入表.xlsx", "test/tupusetting.xlsx")
