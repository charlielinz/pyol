with open(r"C:\Users\charlielin\Envs\pyol\pyol\Body_weeklyreport.txt", encoding="utf-8") as content:
    body = content.readlines()
bodyconcat = ""
for text in body:
    bodyconcat += text

body_content = "".join(body)


print(body_content)