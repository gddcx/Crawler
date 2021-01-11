from selenium import webdriver
# from seleniumrequests import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException
import time
from bs4 import BeautifulSoup
import pandas as pd
import glob
import re
import os


detection_times = 1

def click(d, element, index=-1):
    try:
        e = WebDriverWait(driver, 5, 0.5).until(
            EC.presence_of_all_elements_located((By.XPATH, element))
        )
        time.sleep(5)
        e[index].click()
    except:
        # 人机检测
        try:
            global detection_times
            sm_ico = d.find_element_by_xpath("//div[@id='SM_BTN_1']/div[@id='rectMask']")
            time.sleep(5)
            sm_ico.click()
            time.sleep(5)
            action = ActionChains(d)
            elem = "//div[@id='nc_" + str(detection_times) + "_n1t']/span[@id='nc_" + str(detection_times) + "_n1z']"
            detection_times += 1
            dragger = d.find_element_by_xpath(elem)
            action.click_and_hold(dragger).perform()
            action.drag_and_drop_by_offset(dragger, 258, 0).perform()
        except:
            pass
        time.sleep(5)
        click(d, element, index)
        time.sleep(5)

class Search(object):
    def __init__(self, d, university):
        self.university = university
        self.driver = d
        self.soup = None
        self.titles = None
        self.other_item_info = None
        self.law_info = None

    def find_all(self):
        # 检索
        self.driver.find_element_by_id('textarea').send_keys("(ALL=(" + self.university + " )) AND (AD=[20070101 to 20171231])")
        click(self.driver, "//div[@class='button citation_btn']/input[@type='button']")
        # 必须加延时，等待加载出来，不然报错
        time.sleep(5)
        # 显示 100 条
        show_items_per_page = driver.find_elements_by_xpath("//div[@class='page']/div[@class='page_left']/a")
        show_items_per_page[-1].click()
        time.sleep(5)
        print('sort by date')
        # 按申请日排序
        click(self.driver, "//div[@class='checked-field']/p[@id='sortText']")
        time.sleep(5)
        click(self.driver, "//div[@class='two_menu-con two_menu-con1']/ul/li[@id='AD_ASC']/div")
        time.sleep(5)
        # 点击全字段
        click(self.driver, "//li[@id='alertCustomFields']/div[@class='checked-field fields']")
        time.sleep(5)
        click(self.driver, "//input[@id='customall']")
        time.sleep(5)
        click(self.driver, "//input[@class='retrieval']", index=45)

    @property
    def total_patent(self):
        return self.driver.find_element_by_xpath("//div[@class='patent_count']/span[@id='totalCountspan']").text

    def pull_page_source(self):
        # 获取html信息
        time.sleep(5)
        content = self.driver.page_source.encode('utf-8')
        # 要下载数据，要加载、运行js的地方都需要延时一下
        time.sleep(5)
        self.soup = BeautifulSoup(content, 'lxml')

    def get_titles(self):
        # titles
        titles = self.soup.select('tr > td > div > div' '.title-name')
        self.titles = [t.string for t in titles]

    def get_other(self):
        label_ul = self.soup.select('tr > td > div > ul')
        patent_information_list_original = []
        for i in range(100):
            li = label_ul[i].select('li')
            items_dict = {}
            for l in li:
                try:
                    k = l.strong.text.strip()
                    v = l.div.attrs['value'].strip()
                    items_dict[k] = v
                except:
                    try:
                        k = l.strong.text.strip()
                        v = l.div.text.strip()
                        items_dict[k] = v
                    except:
                        continue
            patent_information_list_original.append(items_dict)
        item_require = ['申请号：', '申请日：', '申请人：', '权利要求数量：', '同族国家：', '同族数量：', '被引证次数：', 'IPC分类号：']
        patent_info_list_arrange = []
        for patent_info in patent_information_list_original:
            one_patent_require = []
            for k in item_require:
                if k in patent_info.keys():
                    v = patent_info[k]
                    one_patent_require.append(v)
                else:
                    one_patent_require.append(' ')
            patent_info_list_arrange.append(one_patent_require)

        self.other_item_info = patent_info_list_arrange

    def _get_law_one_patent(self):
        time.sleep(4)
        content_local = self.driver.page_source.encode('utf-8')
        time.sleep(4)
        soup_local = BeautifulSoup(content_local, 'lxml')
        table_rows = soup_local.select('div > div > div > div > table > tr')[1:]
        # table_list = []
        table_string = ''
        for row in table_rows:
            # 所有列
            tds = row.select('td')
            for t in tds:
                # table_list.append(''.join(t.text.split()).replace("'", ''))
                table_string = table_string + '# ' + t.text #''.join(t.text.split())
        # return table_list
        return table_string

    def get_law_info(self):
        law_info_list = []
        # 进入法律状态页面
        click(self.driver, "//div[@class='title-name']/a[@class='highlight_ALL']", 0)
        time.sleep(4)
        while len(self.driver.window_handles) < 2:
            click(self.driver, "//div[@class='title-name']/a[@class='highlight_ALL']", 0)
            time.sleep(4)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        click(self.driver, "//li[@id='lawTab']")
        time.sleep(4)
        # 第一篇patent
        law_info_list.append(self._get_law_one_patent())
        # 下一篇patent
        for i in range(19):
            click(self.driver, "//a[@id='nextBtn']")
            time.sleep(4)
            law_info_list.append(self._get_law_one_patent())
        self.law_info = law_info_list

    @property
    def patent_info(self):
        self.pull_page_source()
        self.get_titles()
        self.get_other()
        time.sleep(4)
        # self.get_law_info()
        # print(self.law_info)
        twenty_patent = []
        # for t, info, law in zip(self.titles, self.other_item_info, self.law_info):
        for t, info in zip(self.titles, self.other_item_info):
            print(t, info)
            one_patent = []
            one_patent.append(t)
            one_patent.extend(info)
            # one_patent.append(law)
            twenty_patent.append(one_patent)
        return twenty_patent


def write_to_excel(twenty_patent, university, page):
    df = pd.DataFrame(twenty_patent)
    df.to_excel(university + str(page) + '-' + str(page + 99) + '.xlsx')


def restart(path='.'):
    files = glob.glob('./*xlsx')
    files.sort(key=lambda x: os.path.getmtime(x))
    if files:
        newest = files[-1]
        pattern_letter = re.compile(r'A|B')
        pattern_uni = re.compile(r'[\u4e00-\u9fa5]+')
        pattern_num = re.compile(r'[0-9]+')
        class_ = pattern_letter.findall(newest)[0]
        university = pattern_uni.findall(newest)[0]
        num = pattern_num.findall(newest)[-1]
        return class_, university, num
    else:
        return None, None, None


if __name__ == '__main__':
    # '北京大学'
    A_class = ['中国人民大学', '清华大学', '北京航空航天大学', '北京理工大学', '中国农业大学', '北京师范大学', '中央民族大学',
               '南开大学', '天津大学', '大连理工大学', '吉林大学', '哈尔滨工业大学', '复旦大学', '同济大学', '上海交通大学', '华东师范大学', '南京大学',
               '东南大学', '浙江大学', '中国科学技术大学', '厦门大学', '山东大学', '中国海洋大学', '武汉大学', '华中科技大学', '中南大学', '中山大学',
               '华南理工大学', '四川大学', '重庆大学', '电子科技大学', '西安交通大学', '西北工业大学',
               '兰州大学', '国防科技大学']
    # A_class = [ '南京大学', '东南大学', '浙江大学', '中国科学技术大学', '厦门大学', '山东大学', '中国海洋大学', '武汉大学', '华中科技大学', '中南大学', '中山大学',
    #            '华南理工大学', '四川大学', '重庆大学', '电子科技大学', '西安交通大学', '西北工业大学',
    #            '兰州大学', '国防科技大学']

    B_class = ['东北大学', '郑州大学', '湖南大学', '云南大学', '西北农林科技大学', '新疆大学']

    JNU = '暨南大学'

    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    # 不加载图片
    # options.add_argument('blink-settings=imagesEnabled=false')
    # 避免被检测
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features")
    options.add_argument("--disable-blink-features=AutomationControlled")
    # 设置浏览器为中文
    options.add_experimental_option('prefs', {'intl.accept_languages': 'zh,zh_CN'})

    driver = webdriver.Chrome(chrome_options=options)
    driver.maximize_window()
    # 避免被检测
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
        Object.defineProperty(navigator, 'webdriver', {
          get: () => undefined
        })
      """
    })
    driver.execute_cdp_cmd("Network.enable", {})
    driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {"headers": {"User-Agent": "browser1"}})


    driver.get('https://www.incopat.com')
    # 点击IP登录
    click(driver, "//a[@id='ipLoginBtn']")
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(4)
    # 点击高级检索
    click(driver, "//a[@class='gj-search']")
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(4)

    # 重新开始
    class_to_list = {'A': A_class, 'B': B_class, 'JNU': JNU}
    class_, uni, num = restart()
    if class_:
        list_now = class_to_list[class_]
        id_ = list_now.index(uni)
        list_now = list_now[id_:]
        if class_ == 'A':
            A_class = A_class[id_:]
        elif class_ == 'B':
            A_class = []
            B_class = B_class[id_:]
        else:
            A_class = []
            B_class = []

    for k, v in {'A': A_class, 'B': B_class, 'JNU': [JNU]}.items():
        for university in v:
            search_university = Search(driver, university)
            driver.refresh()
            time.sleep(5)
            search_university.find_all()

            total_patent = search_university.total_patent

            start = 1
            if num:
                for i in range(int(num) // 100):
                    time.sleep(10)
                    click(driver, "//div[@id='pageSolrDiv']/a")
                start += int(num)
            num = 0

            current_page_num = driver.find_element_by_xpath("//div[@id='pageSolrDiv']/span[@class='current']").text

            if total_patent == start:
                continue
            print('start:', start, ' total patent:', total_patent)
            for page in range(start, int(total_patent), 100):
                try:
                    patent_info = search_university.patent_info
                    write_to_excel(patent_info, k + '-' + university, page)
                    click(driver, "//div[@id='pageSolrDiv']/a")
                    time.sleep(10)
                    if driver.find_element_by_xpath(
                            "//div[@id='pageSolrDiv']/span[@class='current']").text == current_page_num:
                        click(driver, "//div[@id='pageSolrDiv']/a")
                        time.sleep(10)
                    current_page_num = driver.find_element_by_xpath(
                        "//div[@id='pageSolrDiv']/span[@class='current']").text
                except:
                    break
