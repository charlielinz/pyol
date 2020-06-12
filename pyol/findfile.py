import os
path = r"C:\Users\charlielin\Desktop\檔案\工作週報"

a = os.listdir(path)
a.reverse()

location = path + "\\" + a[0]
print(location)
