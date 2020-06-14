import os
this_py_file_path = os.path.abspath(__file__)
this_py_file_dir_path = os.path.dirname(this_py_file_path)
config_dir_path = os.path.join(this_py_file_dir_path,"weekly_report")
print(config_dir_path)