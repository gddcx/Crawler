import os
import requests
from bs4 import BeautifulSoup

url_dict = {
    '2015': 'http://www.nsfc.gov.cn/nsfc/cen/xmzn/2015xmzn/18/index.html',
    '2016':'http://www.nsfc.gov.cn/nsfc/cen/xmzn/2016xmzn/17/index.html',
    '2017':'http://www.nsfc.gov.cn/nsfc/cen/xmzn/2017xmzn/15/index.html',
    '2018':'http://www.nsfc.gov.cn/nsfc/cen/xmzn/2018xmzn/15/index.html',
    '2019':'http://www.nsfc.gov.cn/nsfc/cen/xmzn/2019xmzn/15/index.html',
}

# 伪装成chrome
headers = {'User-Agent': 'User-Agent:Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
           'Connection': 'close'}
for year, url in url_dict.items():
    save_path = os.path.join(r'D:\Project', year, 'raw')
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    html = requests.get(url)
    soup = BeautifulSoup(html.content, 'lxml')
    data = soup.select('a.jjyw')
    print(data)
    # 不同学部
    for d in data:
        ul = d.get('href')
        ul = url.replace('index.html', ul)
        # ul = http://www.nsfc.gov.cn/nsfc/cen/xmzn/2018xmzn/15/01.html
        html = requests.get(ul, headers=headers)
        soup = BeautifulSoup(html.content, 'lxml')
        img_list = soup.select('p > img')
        # 每个学部多个图片
        for img in img_list:
            src = img.get('src')
            img_url = src.replace('..', os.path.dirname(os.path.dirname(ul)))
            # img_url = http://www.nsfc.gov.cn/nsfc/cen/xmzn/2018xmzn/images/15/01_01.jpg
            try:
                img = requests.get(img_url)
                with open(os.path.join(save_path, os.path.basename(img_url)), 'wb') as wf:
                    wf.write(img.content)
            except:
                print(img_url)

#
root = "http://www.nsfc.gov.cn"
url_2020 = "http://www.nsfc.gov.cn/publish/portal0/xmzn/2020/16/"
save_path = os.path.join(r'D:\Project', '2020', 'raw')
if not os.path.exists(save_path):
    os.makedirs(save_path)
html = requests.get(url_2020)
soup = BeautifulSoup(html.content, 'lxml')
data = soup.select('tr>td>a')
for d in data:
    ul = root + d.get('href')
    html = requests.get(ul)
    soup = BeautifulSoup(html.content, 'lxml')
    img_tag_list = soup.select('p>img')
    print(len(img_tag_list))
    for img_tag in img_tag_list:
        img_path = img_tag.get('src')
        img_url = root + img_path
        img = requests.get(img_url)
        with open(os.path.join(save_path, os.path.basename(img_path)), 'wb') as wf:
            wf.write(img.content)
