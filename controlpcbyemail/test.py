from control_pc_by_email import EmailClass
from control_pc_by_email import ControlPCbyEmail

import json
import zjy_tools


# with open('./resources/config.json', 'r') as f:
#     options = json.load(f)
# print(options)
ce = ControlPCbyEmail()
# email = EmailClass(options)
# res = email.send('主题', '内容啊', attach_path="c:\\1.jpg")

ce.run()

