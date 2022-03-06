from controlpcbyemail import EmailClass
from controlpcbyemail import ControlPCbyEmail

import json
import mtools


with open('./resources/config.json', 'r') as f:
    options = json.load(f)
print(options)
ce = ControlPCbyEmail()
email = EmailClass(options)
res = email.send('主题', '内容啊', attach_path="c:\\1.jpg")
# res = email.send('MT')
# res = email.get('list')
print(res)
