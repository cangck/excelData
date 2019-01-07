import re
import requests
import hashlib
import time

# respose=requests.get('http://www.xiaohuar.com/v/')
# # print(respose.status_code)# 响应的状态码
# # print(respose.content)  #返回字节信息
# # print(respose.text)  #返回文本内容
# urls=re.findall(r'class="items".*?href="(.*?)"',respose.text,re.S)  #re.S 把文本信息转换成1行匹配
# url=urls[5]
# result=requests.get(url)
# mp4_url=re.findall(r'id="media".*?src="(.*?)"',result.text,re.S)[0]
#
# video=requests.get(mp4_url)
#
# with open('D:\\a.mp4','wb') as f:
#     f.write(video.content)
#


def get_index(url):
    respose = requests.get(url)
    if respose.status_code==200:
        return respose.text

def parse_index(res):
    urls = re.findall(r'class="items".*?href="(.*?)"', res,re.S)  # re.S 把文本信息转换成1行匹配
    return urls


def get_detail(urls):
    for url in urls:
        if not url.startswith('http'):
            url='http://www.xiaohuar.com%s' %url
        result = requests.get(url)
        if result.status_code==200 :
            mp4_url_list = re.findall(r'id="media".*?src="(.*?)"', result.text, re.S)
            if mp4_url_list:
                mp4_url=mp4_url_list[0]
                print(mp4_url)
                save(mp4_url)


def save(url):
    video = requests.get(url)
    if video.status_code==200:
        m=hashlib.md5()
        m.updata(url.encode('utf-8'))
        m.updata(str(time.time()).encode('utf-8'))
        filename=r'%s.mp4'% m.hexdigest()
        filepath=r'D:\\%s'%filename
        with open(filepath, 'wb') as f:
            f.write(video.content)

def main():
    for i in range(5):
        res1 = get_index('http://www.xiaohuar.com/list-3-%s.html'% i )
        res2 = parse_index(res1)
        get_detail(res2)

if __name__ == '__main__':
    main()