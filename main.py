#!/usr/bin/env python
# coding: utf-8

# # PKU Course Crawler
# 
# 北京大学课程信息爬取工具（Python脚本+Jupyter Notebook版本）
# 
# ## 特性
# 
# - 🕸️ 基于Selenium的网页自动化爬取
# - 🔐 北大门户账号登录, 无外泄风险
# - ⚙️ 可配置学年/院系/学期/学生类型过滤
# - 📈 Notebook版本含数据处理和分析功能
# - 🤖 自动管理ChromeDriver
# 
# ## 0. 准备工作
# #### 包导入

# In[5]:


from pypinyin import pinyin, Style
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os
import re
import json
import pandas as pd
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException,  WebDriverException


# #### 设置加载

# In[6]:


with open("setting.json", "r", encoding="utf-8") as f:
    loaded_data = json.load(f)  


# #### 常用函数

# In[23]:


# 修改后的安全获取表格数据函数
def get_table_data_safely(driver):
    """安全获取表格数据的函数，包含多重保险机制"""
    try:
        # 使用改善后的等待条件
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.row-border-table"))
        )
        # 检查是否有数据行（包含表头时至少2行）
        WebDriverWait(driver, 3).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, "table.row-border-table tr")) >= 2
        )
    except TimeoutException:
        return []  # 没有有效数据时提前返回

    try:
        # 一次性获取整个表格的HTML（原子操作）
        table_html = driver.find_element(By.CSS_SELECTOR, "table.row-border-table").get_attribute("outerHTML")
    except StaleElementReferenceException:
        return []  # 如果表格已消失返回空

    # 使用BeautifulSoup解析静态HTML
    soup = BeautifulSoup(table_html, "html.parser")
    rows = soup.select("tr")[1:]  # 跳过表头

    table_data = []
    for row in rows:
        cols = row.find_all("td")
        # 直接获取文本内容，无需与浏览器元素交互
        table_data.append([col.get_text(strip=True) for col in cols])
    
    return table_data


def safe_click(driver, element):
    """更稳定的点击方式"""
    try:
        element.click()
    except WebDriverException:
        driver.execute_script("arguments[0].click();", element)


def click_portal_entry(driver, wait, entry_id):
    """等待门户足迹入口真正可点击后再进入，避免遮挡导致点击失败"""
    entry = wait.until(EC.element_to_be_clickable((By.ID, entry_id)))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", entry)
    safe_click(driver, entry)

def extract_department_name(department):
    match = re.match(r'^\d{5}-(.*)', department)
    if match:
        return match.group(1)
    return department


# ## 1. 课表查询爬虫
# 本部分将爬下来的内容导入到 '课程数据汇总.xlsx'
# #### 登录, 打开课程查询页面

# In[33]:


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://iaaa.pku.edu.cn/iaaa/oauth.jsp?appID=portal2017&appName=%E5%8C%97%E4%BA%AC%E5%A4%A7%E5%AD%A6%E6%A0%A1%E5%86%85%E4%BF%A1%E6%81%AF%E9%97%A8%E6%88%B7%E6%96%B0%E7%89%88&redirectUrl=https%3A%2F%2Fportal.pku.edu.cn%2Fportal2017%2FssoLogin.do")
wait = WebDriverWait(driver, 10)

username_field = wait.until(EC.presence_of_element_located((By.ID, 'user_name'))) 
password_field = wait.until(EC.presence_of_element_located((By.ID, 'password')))  
username_field.send_keys(loaded_data["username"])  
password_field.send_keys(loaded_data["password"])  
password_field.send_keys(Keys.RETURN)

click_portal_entry(driver, wait, 'courseQuery')
time.sleep(2)  
driver.switch_to.window(driver.window_handles[-1])  

year_input = wait.until(EC.visibility_of_element_located(
    (By.CSS_SELECTOR, 'input[ng-model="year"]')
))


# #### 遍历收集数据

# In[35]:


all_data = []

for term in loaded_data["year"]:
    year_input.clear()
    year_input.send_keys(term)
    time.sleep(1)
    
    dept_select = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "select.cell-sel[ng-model='dept']"))
    )
    all_depts = dept_select.find_elements(By.TAG_NAME, "option")[1:]

    # 如果 loaded_data["dept"] 为空，则遍历所有院系, 否则只遍历指定院系
    for dept in all_depts:
        if loaded_data["dept"] and dept.text not in loaded_data["dept"]:
            continue
        dept_name = dept.text
        dept.click()
        time.sleep(1) 
        table_select = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "select.cell-sel[ng-model='tableType']"))
        )
        all_tables = table_select.find_elements(By.TAG_NAME, "option")[1:]
        # 如果 loaded_data["stu"] 为空，则遍历所有表格, 否则只遍历指定表格
        for table in all_tables:
            if loaded_data["stu"] and table.text not in loaded_data["stu"]:
                continue
            table_name = table.text
            table.click()
            time.sleep(1)
            term_select = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "select.cell-sel[ng-model='term']"))
            )
            all_terms = term_select.find_elements(By.TAG_NAME, "option")[1:]
            # 如果 loaded_data["term"] 为空，则遍历所有内部学期, 否则只遍历指定内部学期
            for inner_term in all_terms:
                if loaded_data["term"] and inner_term.text not in loaded_data["term"]:
                    continue
                inner_term_name = inner_term.text
                inner_term.click()
                time.sleep(1)
                
                # 点击查询按钮
                search_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "input.cell-btn[ng-click='retrCourseTable()']"))
                )
                search_button.click()
                time.sleep(1)
                
                # 等待表格加载完成
                try:
                    table_data = get_table_data_safely(driver)
                except StaleElementReferenceException:
                    # 发生异常时重试一次
                    table_data = get_table_data_safely(driver)
                    
                if not table_data:
                    continue
                # 处理已缓存的数据
                for col_texts in table_data:
                    course_data = {
                        "学年学期": term,
                        "院系": dept_name,
                        "表格类型": table_name,
                        "内部学期": inner_term_name,
                        "课程名": col_texts[0] if len(col_texts) > 0 else "",
                        "课程类别": col_texts[1] if len(col_texts) > 1 else "",
                        "参考学分": col_texts[2] if len(col_texts) > 2 else "",
                        "班号": col_texts[3] if len(col_texts) > 3 else "",
                        "授课教师": col_texts[4] if len(col_texts) > 4 else "",
                        "起止周": col_texts[5] if len(col_texts) > 5 else "",
                        "上课时间": col_texts[6] if len(col_texts) > 6 else "",
                        "备注": col_texts[7] if len(col_texts) > 7 else ""
                    }
                    all_data.append(course_data)
df = pd.DataFrame(all_data)



# #### 保存为表格文件

# In[24]:


df["院系"] = df["院系"].apply(extract_department_name)
df.to_excel("课表信息汇总.xlsx", index=False)
# df.to_csv("课表信息汇总.csv", index=False, encoding='utf-8-sig')
print("数据收集完成，已保存到'课表信息汇总.xlsx'")


# ## 2. 课程介绍爬虫
# 用于课号标注
# ### 登录, 打开课程介绍页面

# In[11]:


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://iaaa.pku.edu.cn/iaaa/oauth.jsp?appID=portal2017&appName=%E5%8C%97%E4%BA%AC%E5%A4%A7%E5%AD%A6%E6%A0%A1%E5%86%85%E4%BF%A1%E6%81%AF%E9%97%A8%E6%88%B7%E6%96%B0%E7%89%88&redirectUrl=https%3A%2F%2Fportal.pku.edu.cn%2Fportal2017%2FssoLogin.do")
wait = WebDriverWait(driver, 10)

username_field = wait.until(EC.presence_of_element_located((By.ID, 'user_name'))) 
password_field = wait.until(EC.presence_of_element_located((By.ID, 'password')))  
username_field.send_keys(loaded_data["username"])  
password_field.send_keys(loaded_data["password"])  
password_field.send_keys(Keys.RETURN)

click_portal_entry(driver, wait, 'courseIntro')
time.sleep(2)  
driver.switch_to.window(driver.window_handles[-1])  

dept_select0 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "select.cell-sel[ng-model='ciDept']")))


# #### 遍历收集数据

# In[12]:


all_data0 = []
all_depts0 = dept_select0.find_elements(By.TAG_NAME, "option")[1:]
    # 如果 loaded_data["dept"] 为空，则遍历所有院系, 否则只遍历指定院系
for dept in all_depts0:
    if loaded_data["dept"] and dept.text not in loaded_data["dept"]:
        continue
    dept_name = dept.text
    dept.click()
    time.sleep(1)
    search_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "input.cell-btn[ng-click='retrCourseList()']"))
    )
    search_button.click()
    time.sleep(1)
    
    # 等待表格加载完成
    try:
        table_data = get_table_data_safely(driver)
    except StaleElementReferenceException:
        # 发生异常时重试一次
        table_data = get_table_data_safely(driver)
        
    if not table_data:
        continue
    # 处理已缓存的数据
    for col_texts in table_data:
        course_data = {
            "院系": dept_name,
            "课程号": col_texts[0] if len(col_texts) > 0 else "",
            "课程名": col_texts[1] if len(col_texts) > 1 else "",
            "课程英文名": col_texts[2] if len(col_texts) > 2 else "",
            "参考学分": col_texts[3] if len(col_texts) > 3 else "",
            "周学时": col_texts[4] if len(col_texts) > 4 else "",
            "总学时": col_texts[5] if len(col_texts) > 5 else "",
            "修读对象": col_texts[6] if len(col_texts) > 6 else "",
        }
        all_data0.append(course_data)
df0 = pd.DataFrame(all_data0)


 


# #### 保存为表格文件

# In[25]:


df0["院系"] = df0["院系"].apply(extract_department_name)
df0.to_excel("课程信息汇总.xlsx", index=False)
# df0.to_csv("课程信息汇总.csv", index=False, encoding='utf-8-sig')
print("数据收集完成，已保存到'课程信息汇总.xlsx'")


# ## 3. 数据处理

# In[26]:


# 读数据
df = pd.read_excel("课表信息汇总.xlsx")
df0 = pd.read_excel("课程信息汇总.xlsx")


# In[27]:


# 首先用课程名匹配, 为df添加df0的内容, 如果重复的键值不一样, 比如参考学分_x和参考学分_y, 则报错
df0 = df0.rename(columns={"课程名": "课程名_课程"})
df1 = pd.merge(df, df0, how="left", left_on="课程名", right_on="课程名_课程", suffixes=("", "_y"))
# 删除多余的列
df1 = df1.drop(columns=["课程名_课程", "参考学分_y", "院系_y"])
# 调整列顺序, 原本是学年学期,院系,表格类型,内部学期,课程名,课程类别,参考学分,班号,授课教师,起止周,上课时间,备注,课程号,课程英文名,周学时,总学时,修读对象
# 调整为学年学期,院系,表格类型,内部学期,课程号,课程名,课程英文名,班号,修读对象,课程类别,参考学分,周学时,总学时,授课教师,起止周,上课时间,备注
df1 = df1[["学年学期", "院系", "表格类型", "内部学期", "课程号", "课程名", "课程英文名", "班号", "修读对象", "课程类别", "参考学分", "周学时", "总学时", "授课教师", "起止周", "上课时间", "备注"]]
# 按课程号合并
df1.to_excel("课表信息汇总+.xlsx", index=False)


# #### 帮 pkuhub.cn 生成一个json

# In[28]:

