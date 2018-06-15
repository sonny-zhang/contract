from Code.twoword_contrast import TwoWords
from Code import config
import os

logpath = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'result')
logfiles = os.listdir(logpath)
if 'result.log' in logfiles:
    os.remove(os.path.join(os.path.abspath(os.path.dirname(__file__)), 'result', 'result.log'))
else:
    pass

status = TwoWords()
chance = config.chance

if chance == 'A':
    print('你选择的录入类型为：A：合同录入+合同文件比对')
    status.write_and_contrast()
elif chance == 'B':
    print('你选择的录入类型为：B：只对合同进行比对')
    status.contrast_only()
elif chance == 'C':
    print('你选择的录入类型为：C：只对合同进行录入')
    status.write_only()
else:
    print('请输入正确的英文编号')
print(status)