import os
import json




this_py_file_path = os.path.abspath(__file__)
this_py_file_dir_path = os.path.dirname(this_py_file_path)

config_folder_dir_path = os.path.join(this_py_file_dir_path, "config_folder")
folder_paths_list = os.listdir(config_folder_dir_path)
config_dir_path = os.path.join(config_folder_dir_path, folder_paths_list[0])
config_path = os.path.join(config_dir_path, "config.txt")



def read_file_as_long_string(config_path):
    with open(config_path, encoding="utf-8") as config_dct:
        lines = config_dct.readlines()
        lines_fixed = []
        for line in lines:
            if line[-1] == '\n':
                lines_fixed += [line[0:-1]]
            else:
                lines_fixed += [line]
    file_as_long_string = "".join(lines_fixed)
    return file_as_long_string 

config_json = read_file_as_long_string(config_path=config_path)
config = json.loads(config_json)

print(config)
