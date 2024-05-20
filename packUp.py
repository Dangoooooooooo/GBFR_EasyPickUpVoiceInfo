import os
import subprocess

def package_scripts(scripts, new_names, icon=None):
    for script, new_name in zip(scripts, new_names):
        command = ["pyinstaller"]

        # 添加参数
        command.append("-F")  # 创建单一的可执行文件
        command.append("-w")  # 不打开命令行窗口

        if icon:  # 如果提供了图标，添加图标参数
            command.append("-i")
            command.append(icon)

        command.append(script)  # 添加要打包的脚本

        # 执行命令
        subprocess.run(command)

        # 重命名生成的exe文件
        os.rename(f"dist/{os.path.splitext(script)[0]}.exe", f"dist/{new_name}.exe")

# 要打包的脚本列表
# scripts = ["A2M.py", "M2E.py", "E2wID1.py", "MixUp1.py", "AltWavName1.py", "GenCvsMaptxt1.py", "STARTER.py"]
# scripts = ["A2M.py", "M2E.py", "E2wID1.py", "MixUp1.py", "AltWavName1.py", "bnkRepack1.py", "STARTER.py"]
scripts = ["STARTER11.py"]

# 新的exe文件名列表
# new_names = ["GBFR#1PickUp_A2M", "GBFR#2PickUp_M2E", "GBFR#3PickUp_E2wID", "GBFR#4AllMixUp", "GBFR#5Alter_wav_Name", "GBFR#6Generate_ConversionMap_txt", "GBFR#0EasySTART"]
# new_names = ["GBFR#1PickUp_A2M", "GBFR#2PickUp_M2E", "GBFR#3PickUp_E2wID", "GBFR#4AllMixUp", "GBFR#5Alter_wav_Name", "GBFR#6bnkRepack", "GBFR#0EasySTART"]
new_names = ["GBFR#EasyPickUpVoiceInfo"]

# 打包脚本
package_scripts(scripts, new_names, icon="favicon.ico")
