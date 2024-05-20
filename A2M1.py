import os
import subprocess
import re
import pandas as pd

def extract_json_from_msg(msg_path, msgpacker_exe):
    try:
        subprocess.run([msgpacker_exe, msg_path])
        json_path = os.path.splitext(msg_path)[0] + ".json"
        if os.path.exists(json_path):
            return json_path
        else:
            return None
    except FileNotFoundError:
        return None

def Action2Motion_Process():
    # 设置msgpacker_exe变量为nier_cli_mgrr.exe的路径
    msgpacker_exe = r'.\Yours\MsgPacker\MsgPacker.exe'

    # 搜索以pl开头，_action.msg结尾的文件
    for root, dirs, files in os.walk(r'.\Yours'):
        for file in files:
            if file.startswith('pl') and file.endswith('_action.msg'):
                msg_path = os.path.join(root, file)
                # 找到文件后，打印路径
                print(f'Found .msg file: {msg_path}')
                break

    json_path = extract_json_from_msg(msg_path, msgpacker_exe)
    if json_path:
        # 在这里添加代码来提取您想要的内容并将其放入 Excel 中
        # 例如，您可以使用 openpyxl 库来创建和编辑 Excel 文件
        # 确认文件大小
        file_path = json_path
        file_size = os.path.getsize(file_path)
        # print(f'Json file size: {file_size} bytes')

        # 读取文件
        with open(file_path, 'r', encoding='iso-8859-1', buffering=3 * file_size) as f:
            content = f.read()
            # print(f'File content: {content}')

        # 使用正则表达式查找"ActionInfo"模块
        pattern = r'"ActionInfo":\s*{[^}]*}'
        matches = re.findall(pattern, content, re.DOTALL)

        # 提取"id_", "abilityTag_"和"saveMotId01_"到"saveMotId10_"的值
        data = []
        for i, match in enumerate(matches, 1):
            id_ = re.search(r'"id_":\s*([^,]*)', match)
            abilityTag_ = re.search(r'"abilityTag_":\s*([^,]*)', match)
            saveMotIDs = [re.search(fr'"saveMotId0{j}_":\s*([^,]*)', match) for j in range(1, 11)]
            row = [id_.group(1) if id_ else None, abilityTag_.group(1) if abilityTag_ else None] + [
                m.group(1) if m else None for m in saveMotIDs]
            # 去掉双引号并将"-"转化为空字符串
            row = [x.strip('"') if x and x.strip('"') != '-' else '' for x in row]
            data.append(row)
            # print(f"{row}")
            if i >= 200:
                break

        # 创建DataFrame并保存到Excel文件
        df = pd.DataFrame(data, columns=['ActionID', 'AbilityTag'] + [f'MotionID_{j}' for j in range(1, 11)])
        df_path = os.path.join(r'.\Yours', 'GBFR#1Action_Motion_ComparisonTable.xlsx')
        df.to_excel(df_path, sheet_name='Action2Motion', index=False)
        print(f'Excel updated data saved in {df_path}')
    else:
        print("Convert Failed！Please Check Path.")

if __name__ == "__main__":
    Action2Motion_Process()
