import os
import shutil
import pandas as pd
import glob

# 找到路径.\Yours\下以“GBFR#0Pl”开头，以“VoiceInfo_MixUp”结尾的Excel表
def AltWavName_Process():
    excel_files = glob.glob(r".\Yours\GBFR#0Pl*VoiceInfo_MixUp.xlsx")
    # 读取找到的第一个Excel文件
    df = pd.read_excel(excel_files[0])

    # 创建目标目录
    if not os.path.exists(r".\Yours\wav_AltName"):
        os.makedirs(r".\Yours\wav_AltName")

    # 遍历Excel文件中的每一行
    for index, row in df.iterrows():
        # 获取Alternative Voice Name和PL_vo_Event_WithSerial的值
        alt_voice_name = row['Alternative Voice Name']
        pl_vo_event_with_serial = row['PL_vo_Event_WithSerial']
        bnk_belonging = row['bnk_Belonging']

        # 在wav_org及其所有子目录中查找原始文件
        org_file_path = None
        for root, dirs, files in os.walk(r".\Yours\wav_org"):
            if f"{alt_voice_name}.wav" in files:
                org_file_path = os.path.join(root, f"{alt_voice_name}.wav")
                break

        # 检查原始文件是否存在
        if org_file_path and os.path.exists(org_file_path):
            # 根据bnk_Belonging的值，将文件放到.\Yours\wav_AltName下以bnk_Belonging为名的子文件夹里
            target_dir = os.path.join(r".\Yours\wav_AltName", str(bnk_belonging))
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)
            # 构建目标文件的路径
            if pd.notnull(row['Alternative Voice Name']):
                if pd.notnull(row['PL_vo_Event_WithSerial']):
                    pl_vo_event_with_serial = str(pl_vo_event_with_serial).replace('\n', ',')  # 删除换行符
                    target_file_path = os.path.join(target_dir, f"{pl_vo_event_with_serial}.wav")
                else:  # 如果PL_vo_Event_WithSerial没值，就返回原值
                    target_file_path = os.path.join(target_dir, f"{alt_voice_name}.wav")
                # 拷贝并重命名文件
                shutil.copy(org_file_path, target_file_path)

            # 根据bnk_Belonging的值，创建.\Yours\wem_FromWwise下以bnk_Belonging为名的子文件夹
            target_dir_wem = os.path.join(r".\Yours\wem_FromWwise", str(bnk_belonging))
            if not os.path.exists(target_dir_wem):
                os.makedirs(target_dir_wem)

    print("wav_AltName is generated in: .\\Yours\\wav_AltName\\")
    print("wem_FromWwise is generated in: .\\Yours\\wem_FromWwise\\")

if __name__ == "__main__":
    AltWavName_Process()
