config_dir_path = r"C:\Users\charlielin\Envs\pyol\pyol\AutoSendEmail_Pro\weekly_report"
with open(config_dir_path + r"\body_config.txt", encoding="utf-8") as file3:
    body = file3.readlines()
body_config = "".join(body)
print(body_config)