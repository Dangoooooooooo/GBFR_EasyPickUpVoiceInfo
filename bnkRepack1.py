import os
import glob
import shutil
import pandas as pd
import re
from modules.bnk import repackBNKFile

def bnkRepack_Process():
    # 读取Excel文件
    excel_files = glob.glob('.\\Yours\\GBFR#0Pl*VoiceInfo_MixUp*.xlsx')
    df = pd.read_excel(excel_files[0], sheet_name='YourRecord')

    # 遍历子文件夹
    for subdir in glob.glob('.\\Yours\\wem_FromWwise\\*'):
        if os.path.isdir(subdir) and subdir.endswith('.bnk'):
            subdir_name = os.path.basename(subdir)
            print(f"Dealing with {subdir_name} folder...")
            if subdir_name in df['bnk_Belonging'].astype(str).values:
                wem_files = glob.glob(subdir + '\\*.wem')
                if len(wem_files) > 0:
                    new_subdir = subdir + '_TEMP'
                    # 检查文件夹是否存在，如果不存在，再创建
                    if not os.path.exists(new_subdir):
                        os.makedirs(new_subdir)
                    for file in wem_files:
                        filename = os.path.basename(file)
                        # 处理文件名，删除下划线和后面的8位数字或字母
                        filename_processed = re.sub(r'_\w{8}\.wem$', '', filename)
                        for index, row in df.iterrows():
                            if str(row['PL_vo_Event_WithSerial']) == filename_processed and str(row['bnk_Belonging']) == subdir_name:
                                # print(f"{str(row['PL_vo_Event_WithSerial'])} - {str(row['bnk_Belonging'])} - {row['NumSerial']} - {str(row['wemID'])}")
                                new_filename = str(row['wemID']) + '.wem'
                                new_file_path = os.path.join(new_subdir, new_filename)
                                # 如果文件已存在，先删除
                                if os.path.exists(new_file_path):
                                    os.remove(new_file_path)
                                # 检查文件是否存在，如果存在，再复制
                                if os.path.exists(file):
                                    shutil.copy(file, new_file_path)
            print(f"Successfully converted all [.wem] files named EventName in the {subdir_name} folder to .wem files with the original ID")

    # 执行命令行
    for subdir in glob.glob('.\\Yours\\wem_FromWwise\\*_TEMP'):
        print(f"{subdir}")
        subdir_name = os.path.basename(subdir)
        # 检查subdir_name的后缀
        if subdir_name.endswith('.bnk_TEMP'):
            SOURCEFILE = f"{os.path.join('.\\Yours\\', subdir_name.replace(".bnk_TEMP",".bnk"))}"
            INPUTDIR = f"{subdir}"
            EXPORTDIR = os.path.join(INPUTDIR,"export")
            repackBNKFile(SOURCEFILE, INPUTDIR, EXPORTDIR)

    # 在所有的处理完成后
    os.makedirs('.\\Yours\\bnkEXPORT', exist_ok=True)  # 创建.\Yours\bnkEXPORT文件夹，如果已经存在，不会抛出错误
    export_dirs = glob.glob('.\\Yours\\wem_FromWwise\\*_TEMP\\export')
    for export_dir in export_dirs:
        # 检查是否存在export文件夹
        if os.path.exists(export_dir):
            # 将export文件夹中的所有文件移动到.\Yours\bnkEXPORT文件夹下
            for file in glob.glob(export_dir + '\\*'):
                shutil.move(file, '.\\Yours\\bnkEXPORT\\' + os.path.basename(file))
                # 删除new_subdir文件夹
                # shutil.rmtree(export_dir[:-7])  # 去掉'\export'得到new_subdir的路径
    print(f"Successfully moved all [.bnk] files in the '.\\Yours\\bnkEXPORT\\' folder")

if __name__ == "__main__":
    bnkRepack_Process()