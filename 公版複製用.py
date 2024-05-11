import os
import shutil

# 1. 找到名稱中包含"112_2班疲勞量表_公版"的xlsx檔案
source_files = [file for file in os.listdir() if file.endswith('.xlsx') and '112_2班疲勞量表_公版' in file]
if not source_files:
    print("找不到符合條件的xlsx檔案")
    exit()

source_file = source_files[0]

# 2. 複製檔案並命名為47至90，不包含55
target_folder = "疲勞量表"
if not os.path.exists(target_folder):
    os.makedirs(target_folder)

for number in range(47, 91):
    if number != 55:
        target_file = os.path.join(target_folder, str(number) + ".xlsx")
        shutil.copyfile(source_file, target_file)
        print(f"已複製檔案至 {target_file}")
