with open(r"C:\Users\charlielin\Envs\pyol\pyol\Weekly_Report\To.txt", encoding="utf-8") as file1:
    lines = file1.readlines()
    a = 0
    if lines[a] != "":
        print(lines[a])
        a = a+1    