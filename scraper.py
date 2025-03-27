import time
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from bs4 import BeautifulSoup

account = '133********' # 你的电话
password = '******'     # 你的密码

# 处理我们需要获取的信息
def basic_nes(driver, herf1):
    driver.get(herf1)   # 进入新的网址
    time.sleep(20)      # 暂停20s
    
    # 查找公司的代码
    phone = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[1]/div/div[2]/div/div/div[1]/div/div[1]/div[1]/span/span[2]/span/span[1]').text
    return phone

################################################# 下面是谷歌浏览器的基本登录验证 #################################################
# 模拟使用Chrome浏览器登陆
options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")  
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_argument("--start-maximized")
driver =  webdriver.Chrome(options=options)
driver.implicitly_wait(10)

# 打开搜索页
driver.get('https://www.qcc.com/?utm_source=baidu1&utm_medium=cpc&utm_term=pzsy')
time.sleep(5)  # 暂停5s
time_start = time.time()

# 模拟登陆：Selenium Locating Elements by Xpath
time.sleep(np.abs(np.random.normal(2, 0.5)))

# 切换到密码登录
driver.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/div/div[2]/div/div/div[4]/img').click()
time.sleep(np.abs(np.random.normal(2, 0.5)))
driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/div[2]/div/div/div[2]/div[1]/div[2]/a").click()
time.sleep(np.abs(np.random.normal(2, 0.5)))

# 填写基本信息
driver.find_element(By.XPATH, "//input[@placeholder='请输入手机号码/用户名']").send_keys(account)
driver.find_element(By.XPATH, "//input[@placeholder='请输入密码']").send_keys(password)
time.sleep(np.abs(np.random.normal(2, 0.3)))

# 登录按钮
driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/div[2]/div/div/div[2]/div[3]/button").click()

# 人工验证码
input("请完成验证码验证，输入 ENTER 继续")
time_end = time.time()
print('您的本次登录共用时{}秒。'.format(int(time_end - time_start)))
time.sleep(10)

################################################# 下面是对天津化工筛选公司名称操作 #################################################
# 待输入搜索公司
corps = []

driver.get('https://www.qcc.com/?utm_source=baidu1&utm_medium=cpc&utm_term=pzsy')
time.sleep(np.abs(np.random.normal(2, 2)))

# 清空搜索框
driver.find_element(By.ID, 'searchKey').clear()
# 在搜索框中输入查询企业名单
driver.find_element(By.ID, 'searchKey').send_keys('天津化工')
# 开始搜索
driver.find_element(By.XPATH, "/html/body/div/div[2]/section[1]/div/div/div/div[1]/div/div/span/button").click()        
time.sleep(np.random.randint(5, 11) + np.random.normal(1, 0.5))

# 程序翻页控制
start_page = 1
num_page = 2

# 公司名称存储体
for page_now in np.arange(start_page, start_page + num_page):
    if page_now != 1:
        # 寻找页数元素
        element = driver.find_element(By.XPATH, '(//input[@class="form-control"])[last()]')
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        element.clear()
    
        element.send_keys(str(page_now))
        time.sleep(np.abs(np.random.randint(1, 4) + np.random.normal(1, 0.5)))
        
        # 点击确定按钮跳转
        driver.find_element(By.XPATH, "//a[@class='input-jump-btn' and text()='确定']").click()
        time.sleep(np.abs(np.random.randint(1, 4) + np.random.normal(1, 0.5)))
    
    # 获取网页源码并解析
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'html.parser')
    
    # 根据实际HTML结构调整选择器
    spans = soup.find_all('span')
    for span in spans:
        # 对标签进行清洗
        bs = BeautifulSoup(str(span), 'html.parser')
        clean_text = bs.get_text(strip=True)
    
        # 检测里面是否存在关键字
        # print('开始检索:', clean_text)
        if clean_text.endswith("公司存续") or clean_text.endswith("厂存续"):
            corps.append(clean_text[:-2])
    
# pf = pd.DataFrame({"Company": corps})
# file_name = '企业名称爬取.xlsx'
# with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
#     # 将 DataFrame 写入 Excel
#     pf.to_excel(writer, sheet_name='Name', index=False)
    
#     # 获取工作簿和工作表对象
#     workbook  = writer.book
#     worksheet = workbook['Name']
#     # 设置列宽
#     worksheet.column_dimensions['A'].width = 50
    
# print(f"Excel 文件 (文件名: {file_name}) 已保存")

# 扩展内容
corps.extend([
    '武汉中粮肉食品有限公司',
    '北京中星微人工智能芯片技术有限公司'
])

################################################# 下面是对确定的公司名称进行的爬取操作 #################################################
date_f = list()
for i in range(len(corps)):
    #待输入搜索公司
    corp = corps[i]
    
    # 进入到搜索页面
    driver.get('https://www.qcc.com/?utm_source=baidu1&utm_medium=cpc&utm_term=pzsy')
    time.sleep(np.abs(np.random.normal(2, 2)))
    
    # 清空搜索框
    driver.find_element(By.ID, 'searchKey').clear()
    # 在搜索框中输入查询企业名单
    driver.find_element(By.ID, 'searchKey').send_keys(corp)
    
    # 开始搜索
    driver.find_element(By.XPATH, "/html/body/div/div[2]/section[1]/div/div/div/div[1]/div/div/span/button").click()        
    time.sleep(np.random.randint(5, 11) + np.random.normal(1, 0.5))
    
    cname = driver.find_element(By.XPATH, '//a[@class="title copy-value"]').text
    href1 = None
    while 1:
        try:
            href1 = driver.find_element(By.XPATH, '//a[@class="title copy-value"]').get_attribute("href")
            print('====================', corp, '==================')
            if href1 is not None:
                break
        except:
            time.sleep(np.random.random(1, 2))   
    cnt = 0
    while 1:
        cnt += 1
        if cnt<5:
            try:
                print("开始尝试解析 " + href1)
                date_f.append([corp, basic_nes(driver, href1)])
                break
            except:
                pass
        else:
            print('请单独对', corp, '的信息进行获取')
            break

driver.close()

pf =  pd.DataFrame(date_f)
pf.columns = [
    'Company',
    'Code'                
]

file_name = '企业代码.xlsx'
with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    # 将 DataFrame 写入 Excel
    pf.to_excel(writer, sheet_name='Code', index=False)
    
    # 获取工作簿和工作表对象
    workbook  = writer.book
    worksheet = workbook['Code']
    # 设置列宽
    worksheet.column_dimensions['A'].width = 50
    worksheet.column_dimensions['B'].width = 30
    
print(f"Excel 文件 (文件名: {file_name}) 已保存")

print('==================已完成企业信息下载==================')