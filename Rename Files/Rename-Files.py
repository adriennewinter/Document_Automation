# -----------------------------------------------------------------------------------
# Batch rename files in a given folder

# Author: Adrienne Winter, 2022.

# README:
# This script will save renamed files into a folder called Renamed_Files and ignores any 
# files or directories with no .extension

# Depending on how you want the files to be renamed, you may need to edit the new_name variable.
# Remember to add an escape \ in your file path name -> "C:File\\Path\\Name"
# -----------------------------------------------------------------------------------

import os
import shutil

# Setting variables
folder_path = "C:\\OpenCV_Dev\\build\\install\\x64\\vc15\\lib"

# Create folder to save renamed files into
os.makedirs(f"{folder_path}\\Renamed_Files")
new_folder_path = f"{folder_path}\\Renamed_Files"

# Iterate through all files in folder_path
for file in os.listdir(folder_path):
    if not os.path.isfile(f"{folder_path}\\{file}"):
        print(f"Skipping {file} because it is not a file.")
    else:
        # Copy all the files over to a new folder
        shutil.copy(f"{folder_path}\\{file}",f"{new_folder_path}\\{file}")
        # Set the new_name
        file_type = file.split(".")[-1]
        old_name = file.split(".")[0]
        new_name = old_name[0:len(old_name)-1]
        # Rename the copied files in the new folder
        os.rename(f"{new_folder_path}\\{old_name}.{file_type}",f"{new_folder_path}\\{new_name}.{file_type}")
        print(f"Saving {old_name}.{file_type} as {new_name}.{file_type}")