from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.action_chains import ActionChains
import sys
import random
import time
import xlwt as xl
import atexit
'''
Bugs:
1、如果10年没有栏目会导致死循环；
2、获取出版周期不准确，如果有多余信息，会导致定位不准确，获取错误的信息
Fix this Bugs : 2024.3.12，控制刷新次数为5，将出版周期改为期刊的基本信息进行输出
Bugs:
1、爬取时间长被服务器拒绝
Fix this Bugs : 2024.3.13，添加爬取的起始页面，相当于爬完每页执行搜索
'''
use_infor = '''python cnki.py keyword1 keyword2 is_core_journals (关键词不限量)
for example:
    python cnki.py 计算机 :检索关于计算机的期刊，包括普刊
    python cnki.py 计算机 1/任意不是0的字符 :检索关于计算机的北大核心期刊
    python cnki.py 计算机 管理 1 :检索关于计算机的北大核心期刊、检索关于管理的北大核心期刊
输出：cnki.xlsx文件，其中每个关键词对应一个sheet，关键词为sheet名称
    每个期刊的信息有：期刊链接、期刊名称，期刊基本信息，主办单位，荣誉，复合影响因子，综合影响因子，近十年栏目，最近一期基金论文占比'''
def exitFunction(wb):
    wb.save("cnki.xlsx")
#期刊名称，出版周期，主办单位，荣誉，复合影响因子，综合影响因子，近十年栏目，基金占比
def decodeJournalInfo(url):
    r_list = []
    driver.get(url)
    driver.implicitly_wait(0.5)
    try:
        WebDriverWait(driver, 100).until(EC.presence_of_element_located(
            (By.XPATH, '''//*[@id="J_sumBtn-stretch"]'''))).click()
    except:
        return "Have no more information"
    try:
        journal_name = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
            (By.XPATH, '''//*[@id="qk"]/div[2]/dl[1]/dd[1]/h3[1]'''))).text
        r_list.append(journal_name)
    except:
        #return "Dont find journal_name"
        r_list.append(" ")
    try:
        journal_baseinfo = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
            (By.XPATH, '''//*[@id="JournalBaseInfo"]''')))
        baseinfo_label = journal_baseinfo.find_elements(By.TAG_NAME, 'label')
        baseinfo_span = journal_baseinfo.find_elements(By.TAG_NAME, 'span')
        base_m = ""
        for t_i in range(len(baseinfo_span)):
            base_m = base_m + baseinfo_label[t_i].text + ":" + baseinfo_span[t_i].text + "-"
        r_list.append(base_m)
    except:
        #return "Dont find journal_period"
        r_list.append(" ")
    try:
        journal_organizers = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
            (By.XPATH, '''//*[@id="JournalBaseInfo"]/li[2]/p[1]/span[1]'''))).text
        r_list.append(journal_organizers)
    except:
        #return "Dont find journal_organizers"
        r_list.append(" ")
    try:
        journal_honor = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
            (By.XPATH, '''//*[@id="evaluateInfo"]/li[3]''')))
        journal_honor_li = journal_honor.find_elements(By.CLASS_NAME, 'database')
        l_h = ""
        for e in journal_honor_li:
            l_h = l_h + e.text + "-"
        r_list.append(l_h)
    except:
        #return "Dont find journal_honor"
        r_list.append(" ")
    #r_list.append(journal_honor)
    try:
        journal_composite_if = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
            (By.XPATH, '''//*[@id="evaluateInfo"]/li[2]/p[1]/span[1]'''))).text
        r_list.append(journal_composite_if)
    except:
        #return "Dont find journal_composite_if"
        r_list.append(" ")
    try:
        journal_comprehensive_if = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
            (By.XPATH, '''//*[@id="evaluateInfo"]/li[2]/p[2]/span[1]'''))).text
        r_list.append(journal_comprehensive_if)
    except:
        #return "Dont find journal_comprehensive_if"
        r_list.append(" ")
    #获取近10年栏目
    try:
        WebDriverWait(driver, 100).until(EC.presence_of_element_located(
            (By.XPATH, '''//*[@id="selectprograma"]/a[1]'''))).click()
    except:
        return "Dont find years form"
    time.sleep(random.randint(5,8))
    # try:
    #     WebDriverWait(driver, 100).until(EC.presence_of_element_located(
    #         (By.XPATH, '''//*[@id="recentThree"]/a[1]'''))).click()
    # except:
    #     return "Dont find 3 years"
    # time.sleep(random.randint(3,6))
    try:
        collayer = driver.find_element(By.XPATH,'''//*[@id="collayer"]''')
        collayer_li = collayer.find_elements(By.TAG_NAME, 'a')
        l_m = ""
        for e in collayer_li:
            l_m = l_m + e.text + "-"
        r_list.append(l_m)
    except:
        return "Dont find collayer"
    #获取基金占比统计,只获取最近一年的
    try:
        WebDriverWait(driver, 100).until(EC.presence_of_element_located(
            (By.XPATH, '''//*[@id="selectstatistics"]/a[1]'''))).click()
    except:
        return "Dont find fund"
    time.sleep(random.randint(10,15))
    #通过xpath找不到svg标签
    # latest_publication_number = WebDriverWait(driver, 1).until(EC.presence_of_element_located(
    #     (By.XPATH, '''//*[@id="yearcontainer"]/div[1]/svg[1]/g[5]/g[1]/text[1]/tspan[1]'''))).text
    # latest_publication_number_by_fund = WebDriverWait(driver, 1).until(EC.presence_of_element_located(
    #     #     (By.XPATH, '''//*[@id="Foundationcontainer"]/div[1]/svg[1]/g[7]/g[1]/text[1]/tspan[1]'''))).text
    try:
        latest_publication = driver.find_element(By.XPATH,'''//*[@id="yearcontainer"]/div[1]''')
        latest_publication_g = latest_publication.find_elements(By.TAG_NAME,'tspan')
        latest_publication_number = latest_publication_g[0].text
        if latest_publication_number == '':
            return "latest_publication_number is '' "
    except:
        return "Dont find publication information"
    try:
        latest_publication_by_fund = driver.find_element(By.XPATH,'''//*[@id="Foundationcontainer"]/div[1]''')
        latest_publication_by_fund_g = latest_publication_by_fund.find_elements(By.TAG_NAME,'tspan')
        latest_publication_number_by_fund = latest_publication_by_fund_g[0].text
    except:
        return "Dont find fund publication information"
    r_list.append(int(latest_publication_number_by_fund)/int(latest_publication_number))
    return r_list
def webserver():
    # # get直接返回，不再等待界面加载完成
    # desired_capabilities = DesiredCapabilities.EDGE
    # desired_capabilities["pageLoadStrategy"] = "none"
    # 设置微软驱动器的环境
    options = webdriver.EdgeOptions()
    # # 设置浏览器不加载图片，提高速度
    # options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    # 创建一个微软驱动器
    driver = webdriver.Edge(options=options)
    return driver
def writexlsx(sh,in_url,in_list,row):
    sh.write(row,0,in_url)
    for i in range(len(in_list)):
        sh.write(row,i+1,in_list[i])
def getQk(keyword,onlybd,sh,start_page):
    driver.get("https://navi.cnki.net/knavi/")
    driver.implicitly_wait(5)
    # 修改属性，使下拉框显示
    opt = driver.find_element(By.ID, 'searchbar_list')  # 定位元素
    # 执行 js 脚本进行属性的修改；arguments[0]代表第一个属性
    driver.execute_script("arguments[0].setAttribute('style', 'display: block;')",opt)
    ActionChains(driver).move_to_element(opt).perform()
    opt_1 = driver.find_element(By.CSS_SELECTOR, 'li[navitype="journals"]')
    ActionChains(driver).move_to_element(opt_1).perform()
    ActionChains(driver).click(opt_1).perform()
        # 传入关键字
    WebDriverWait(driver, 100).until(
        EC.presence_of_element_located((By.ID, 'txt_1_value1'))).send_keys(keyword)
        # 点击搜索
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'btnSearch'))).click()
    time.sleep(random.randint(1,4))
    if onlybd:
        WebDriverWait(driver, 100).until(EC.presence_of_element_located(
            (By.XPATH, '''//*[@id="rightnavi"]/div[1]/div[1]/div[1]/div[1]/a[1]'''))).click()
    time.sleep(random.randint(2, 4))
    # 先提取所有链接，再一个个去查找！！！！！！！！！！
    # 获取总文献数和页数
    res_unm = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
        (By.XPATH, '''//*[@id="rightnavi"]/div[1]/div[1]/span[1]/em'''))).text
    print('resultNumbers:', res_unm)
    res_unm = int(res_unm)
    global raw
    ret = 0
    if res_unm != 0:
    # 读取每一条信息，包括期刊名称、出版周期、复合影响因子、综合影响因子、是否北核、主要栏目（用-拼接）、基金论文占比（上年度平均）
    # 获取页数
        pages_number = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
            (By.XPATH, '''//*[@id="rightnavi"]/div[1]/div[1]/span[2]/em[2]'''))).text
        print('pagesNumbers:',pages_number)
        pages_number = int(pages_number)
        journals_list = []
        ret = pages_number
        if pages_number == 1:
            for i in range(1, res_unm+1):
                tmp_journal = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
                    (By.XPATH, f'''//*[@id="rightnavi"]/div[1]/div[2]/ul[1]/li[{i}]/a[1]'''))).get_attribute('href')
                journals_list.append(tmp_journal)
                # writexlsx(sh,tmp_journal, decodeJournalInfo(tmp_journal),raw)
                # raw = raw + 1
        else:
            for i in range(1, pages_number + 1):
                time.sleep(random.randint(2,5))
                now_page_number = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
                    (By.XPATH, '''//*[@id="rightnavi"]/div[1]/div[1]/span[2]/em[1]'''))).text
                now_page_number = int(now_page_number)
                if i<start_page:
                    WebDriverWait(driver, 100).until(EC.presence_of_element_located(
                        (By.XPATH, '''//*[@id="rightnavi"]/div[1]/div[1]/span[2]/a[2]'''))).click()
                    continue
                if now_page_number == start_page:
                    break
            if start_page == pages_number:
                for i in range(1, res_unm - 21 * (pages_number - 1) + 1):
                    tmp_journal = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
                        (By.XPATH, f'''//*[@id="rightnavi"]/div[1]/div[2]/ul[1]/li[{i}]/a[1]'''))).get_attribute('href')
                    journals_list.append(tmp_journal)
            else:
                for j in range(1, 22):
                    tmp_journal = WebDriverWait(driver, 100).until(EC.presence_of_element_located(
                        (By.XPATH, f'''//*[@id="rightnavi"]/div[1]/div[2]/ul[1]/li[{j}]/a[1]'''))).get_attribute('href')
                    journals_list.append(tmp_journal)
        for j in range(len(journals_list)):
            print("Starting get the page {} ,the number".format(str(start_page)), str(j + 1), "/{} .......".format(len(journals_list)))
            try_numbers = 5
            #确保获取信息
            while 1:
                r = decodeJournalInfo(journals_list[j])
                if isinstance(r,list):
                    writexlsx(sh, journals_list[j],r , raw)
                    print(journals_list[j], decodeJournalInfo(journals_list[j]))
                    raw = raw + 1
                    break
                else:
                    if try_numbers == 0:
                        break
                    try_numbers = try_numbers-1
                    time.sleep(5-try_numbers)
    return ret
raw = 0
if len(sys.argv) == 1:
    print("缺少参数，请按照下面用法正确输入参数！")
    print(use_infor)
elif len(sys.argv) == 2:
    driver = webserver()
    wb = xl.Workbook("UTF-8")
    sh = wb.add_sheet(sys.argv[1])
    start_number = 1
    atexit.register(exitFunction,wb)
    while 1:
        n = getQk(sys.argv[1],False,sh,start_number)
        if n<=start_number:
            break
        start_number = start_number+1
    wb.save("cnki.xlsx")
else :
    driver = webserver()
    wb = xl.Workbook("UTF-8")
    if sys.argv[-1]!="0":
        for i in range(1,len(sys.argv)-1):
            raw = 0
            sh = wb.add_sheet(sys.argv[i])
            start_number = 1
            while 1:
                n = getQk(sys.argv[1], True, sh, start_number)
                if n <= start_number:
                    break
                start_number = start_number + 1
    else:
        for i in range(1,len(sys.argv)-1):
            sh = wb.add_sheet(sys.argv[i])
            start_number = 1
            while 1:
                n = getQk(sys.argv[1], False, sh, start_number)
                if n <= start_number:
                    break
                start_number = start_number + 1
    wb.save("cnki.xlsx")
