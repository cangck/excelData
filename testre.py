import re

s = '1-4106106372'
res = re.search('(.*)-(.*)', s)
print(res.group(2))
