import os

import pandas as pd
from bs4 import BeautifulSoup
import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

start = time.clock()


def get_rankings(path, URL, title):
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option('excludeSwitches', ['enable-automation'])  # 隐藏 测试软件tab
    # options.add_argument('--headless')    # 静默执行
    # options.add_argument("user-agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36'")
    browser = webdriver.Chrome(options=options)
    browser.get(URL)
    time.sleep(2)
    no_pagedown = 1
    shcools = browser.find_element_by_css_selector(".filter-bar__CountContainer-sc-1glfoa-5.kFwGjm").text.replace(
        ' schools', '').replace(',', '')
    while no_pagedown:
        try:
            browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")  # 移动到页面最下方
            time.sleep(4)

            soup = BeautifulSoup(browser.page_source, 'lxml')
            dataNumber = len(soup.find_all("h2",
                                           class_="Heading__HeadingStyled-sc-1w5xk2o-0-h2 heunUr Heading-sc-1w5xk2o-1 cRrhAX md-mb2"))
            print(f'\r当前已加载{dataNumber}条数据,共需加载{shcools}条', end='')

            button_element = '.button__ButtonStyled-sc-1vhaw8r-1.kDQStt.pager__ButtonStyled-sc-1i8e93j-1.dypUdv.type-secondary.size-large'
            exists = check_element_exists(browser, 'css', button_element)

            if exists:
                button = browser.find_element_by_css_selector(button_element)
                # 当元素遮挡导致无法点击时，进行移动点击，有可能误点广告
                webdriver.ActionChains(browser).move_to_element(button).click(button).perform()

            no_pagedown = 0 if dataNumber >= int(shcools) else no_pagedown

        except Exception as e:
            print('Error:', e)

    soup = BeautifulSoup(browser.page_source, 'lxml')
    divList = soup.find_all('div', class_='DetailCardGlobalUniversities__TextContainer-sc-1v60hm5-3 fInsHn')
    browser.close()
    dataReturn = []

    for div in divList:
        name = div.find('h2').find('a').text
        link = div.find('h2').find("a")['href']
        loc = div.find("p", class_="Paragraph-sc-1iyax29-0 pyUjv").text
        score = div.find_all("dd", class_="QuickStatHug__Description-hb1bl8-1 eXguFl")[0].text
        regist = div.find_all("dd", class_="QuickStatHug__Description-hb1bl8-1 eXguFl")[1].text
        rank = div.find("div", class_="RankList__Rank-sc-2xewen-2 fxzjOx ranked has-badge").text.replace('#', '')
        rank = rank if not rank is None else 'N/A'  # rank存在空情况
        dataReturn.append({'排名': rank, '院校': name, '国家': loc, '评分': score, '注册': regist, '网址': link, })

    writer = pd.ExcelWriter(path, engine='openpyxl')
    if os.path.exists(path):
        writer.book = load_workbook(path)
    df = pd.DataFrame(dataReturn)
    df.to_excel(writer, sheet_name=title, encoding='utf-8', index=False, columns=dataReturn[0].keys())
    writer.save()


def check_element_exists(driver, condition, element):
    # 检查元素是否存在
    try:
        if condition == 'class':
            driver.find_element_by_class_name(element)
        elif condition == 'id':
            driver.find_element_by_id(element)
        elif condition == 'xpath':
            driver.find_element_by_xpath(element)
        elif condition == 'css':
            driver.find_element_by_css_selector(element)
        return True
    except Exception as e:
        print(f'\n寻找元素出错:', e)
        return False


if __name__ == '__main__':
    path = r'../DataCache/22USNews_demo.xlsx'
    page_urls = {
        # 'world': 'https://www.usnews.com/education/best-global-universities/search',
        # 'africa': 'https://www.usnews.com/education/best-global-universities/africa',
        # 'asia': 'https://www.usnews.com/education/best-global-universities/asia',
        'australia-new-zealand': 'https://www.usnews.com/education/best-global-universities/australia-new-zealand',
        # 'europe': 'https://www.usnews.com/education/best-global-universities/europe',
        # 'latin-america': 'https://www.usnews.com/education/best-global-universities/latin-america',
    }

    for urlkey in page_urls:
        start_time = int(round(time.time()))
        get_rankings(path, page_urls[urlkey], urlkey)
        print(f'\nElapsed:{round(time.clock() - start, 2)} Seconds for: {urlkey}')

    print(f"Total time: {round(time.clock() - start, 2)} seconds.")
