import os
import shutil
import subprocess
import re
import pandas as pd
from xml.etree import ElementTree as ET

def Motion2Event_Process():
    nier_cli_mgrr_path = r'.\Yours\nier_cli_mgrr\nier_cli_mgrr.exe'

    pattern = re.compile(r'pl\d{4}')# 正则表达式匹配以"pl"开头并跟随四位数字的模式

    for root, dirs, files in os.walk(r'.\Yours'):# 搜索匹配的文件夹
        for dir_name in dirs:
            if pattern.match(dir_name):
                bxm_folder_path = os.path.join(root, dir_name)
                # 找到匹配的文件夹后，打印路径
                print(f'Found .bxm folder: {bxm_folder_path}')
                break

    # 创建"_se_bxm_Group"文件夹
    se_bxm_group_path = os.path.join(bxm_folder_path, "_se_bxm_Group")
    os.makedirs(se_bxm_group_path, exist_ok=True)

    # 将所有文件名末尾为"_se"的.bxm文件拷贝到"_se_bxm_Group"文件夹
    for file_name in os.listdir(bxm_folder_path):
        if file_name.endswith("_se.bxm"):
            shutil.copy2(os.path.join(bxm_folder_path, file_name), se_bxm_group_path)

    # 使用subprocess方法调用nier_cli_mgrr.exe来打开"_se_bxm_Group"文件夹里的所有文件
    # for file_name in os.listdir(se_bxm_group_path):
    #     subprocess.run([nier_cli_mgrr_path, os.path.join(se_bxm_group_path, file_name)])

    file_list = [os.path.join(se_bxm_group_path, file_name) for file_name in os.listdir(se_bxm_group_path)]
    subprocess.run([nier_cli_mgrr_path, *file_list])

    # 创建"_se_bxm_xml_Group"文件夹
    se_bxm_xml_group_path = os.path.join(bxm_folder_path, "_se_bxm_xml_Group")
    os.makedirs(se_bxm_xml_group_path, exist_ok=True)

    # 将生成的所有xml文件移动到"_se_bxm_xml_Group"文件夹
    for file_name in os.listdir(se_bxm_group_path):
        if file_name.endswith(".xml"):
            # 检查目标路径是否存在同名文件
            if os.path.exists(os.path.join(se_bxm_xml_group_path, file_name)):
                # 如果存在，删除它
                os.remove(os.path.join(se_bxm_xml_group_path, file_name))
            shutil.move(os.path.join(se_bxm_group_path, file_name), se_bxm_xml_group_path)

    # 打印"_se_bxm_xml_Group"文件夹的绝对路径
    # print(f"_se_bxm_xml_Group Absolute Path: {se_bxm_xml_group_path}")

    # 对文件夹“_se_bxm_xml_Group”执行“从xml里提取EventName”这个任务
    # 这里可以调用之前编写的“从xml里提取EventName”的函数
    # 由于这个函数需要在您的环境中运行，所以我在这里只是用注释提示您
    # 请将这个函数的代码复制到这里，然后取消下面这行的注释
    # 从xml里提取EventName(se_bxm_xml_group_path)
    path = se_bxm_xml_group_path

    # 获取目录下所有的XML文件
    xml_files = [f for f in os.listdir(path) if f.endswith('.xml')]

    # 初始化一个列表来存储Excel文件的数据
    data = []

    # 初始化一个变量来存储最大的事件数量
    max_events = 0

    # 处理每一个XML文件
    for xml_file in xml_files:
        # 提取文件名中第一个和第二个下划线之间的字符串
        file_name_string = re.search('_(.*?)_', xml_file).group(1)
        # print(f"MotionID: {file_name_string}")

        # 解析XML文件
        tree = ET.parse(os.path.join(path, xml_file))
        root = tree.getroot()

        # 初始化一个计数器用于事件名称
        event_counter = 0

        # 初始化一个字典来存储事件名称
        event_dict = {}

        # 查找所有属性为"EventName"且以"PL"开头的元素
        for elem in root.iter():
            for key, value in elem.attrib.items():
                if key == "EventName" and value.startswith("PL"):
                    # 保存事件名称到一个变量
                    event_name = f"PL_vo_Event_{event_counter}"
                    event_value = value
                    # print(f"{event_name}: {event_value}")

                    # 将事件名称添加到字典中
                    event_dict[event_name] = event_value

                    # 事件计数器加一
                    event_counter += 1

        # 更新最大的事件数量
        max_events = max(max_events, event_counter)

        # 将文件名字符串和事件名称添加到数据列表中
        data.append([file_name_string] + list(event_dict.values()))

    # 创建列名列表
    columns = ['MotionID'] + [f"PL_vo_Event_{i}" for i in range(max_events)]

    # 将数据列表转换为DataFrame
    df = pd.DataFrame(data, columns=columns)

    # 将DataFrame保存为Excel文件
    df.to_excel(os.path.join(r'.\Yours', "GBFR#2Motion_EventName_ComparisonTable.xlsx"), sheet_name='Motion2EventName', index=False)

    # 打印成功消息
    print(f'Excel updated data saved in {os.path.join(r'.\Yours', "GBFR#2Motion_EventName_ComparisonTable.xlsx")}')

if __name__ == "__main__":
    Motion2Event_Process()

