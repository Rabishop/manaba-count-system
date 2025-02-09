import requests
import json
import time
from urllib.parse import urlparse
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

def load_cookies_from_file(cookie_file='cookies.json'):
    try:
        with open(cookie_file, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        print("请先导出 Chrome 浏览器中的 cookies 并保存为 cookies.json 文件")
        return None

def check_oral_submission(session, base_url, href, headers):
    """检查是否有口头提交"""
    try:
        full_url = f"{base_url}/ct/{href}" if not href.startswith('http') else href
        response = session.get(full_url, headers=headers)
        response.raise_for_status()
        
        # print("full_url", full_url)

        # 保存详细页面HTML以供调试
        with open('detail_page.html', 'w', encoding='utf-8') as f:
            f.write(response.text)
            
        soup = BeautifulSoup(response.text, 'html.parser')
        oral_check = soup.find('span', class_='checked')
        return 1 if oral_check and 'はい' in oral_check.text else 0
    except Exception as e:
        print(f"检查口头提交时发生错误: {str(e)}")
        return 0

def process_submission_table(html_content, all_students, session, base_url, headers):
    """处理提交表格，更新学生提交信息"""
    # 设定学生名单
    student_list = [
        "木下　悠吾", "五江渕　陸", "小須田　侑暉","小林　紘也", "阪田　智也", "佐久間　悠斗", "櫻井　隆之介", "真谷　健悟",
        "関井　恭介"
    ]
    

    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 查找所有的表格行（排除标题行）
    rows = soup.find_all('tr')

    for row in rows[1:]:
        # 获取姓名
        name_cell = row.find('td', class_='left name')
        if not name_cell:
            continue
            
        # 清理姓名文本：优先获取链接中的文本，如果没有链接则获取单元格文本
        name_link = name_cell.find('a')
        name = name_cell.get_text(strip=True)
        
        # 如果名字不在学生名单中，跳过
        if name not in student_list:
            continue
            
        # 检查口头提交
        if name_link:
            href = name_link.get('href')
            oral_submission = check_oral_submission(session, base_url, href, headers)
        else:
            oral_submission = 0
            
        # 获取提交日期
        date_cell = row.find_all('td')[2]
        submission_date = date_cell.get_text(strip=True)

        # 初始化学生记录（如果不存在）
        if name not in all_students:
            all_students[name] = {
                'count': 0,
                'oral_count': 0,
                'dates': []
            }
            
        # 更新口头提交计数
        all_students[name]['oral_count'] += oral_submission
            
        # 记录提交信息（如果不是未提出且日期不重复）
        if '未提出' not in submission_date and submission_date not in all_students[name]['dates']:
            all_students[name]['count'] += 1
            all_students[name]['dates'].append(submission_date)

def save_to_excel(students):
    """将学生提交信息保存到Excel"""
    df_data = [
        {
            '姓名': name,
            '出勤次数': len(set(date.split()[0] for date in data['dates'])),
            '口头提交': data['oral_count'],
            '提交次数': data['count'],
            '提交日期': ', '.join(data['dates'])
        }
        for name, data in students.items()
    ]
    
    df = pd.DataFrame(df_data)
    df.to_excel('提交情况统计.xlsx', index=False)
    print(f"已将统计结果保存到 '提交情况统计.xlsx'")

def main():
    # 加载 cookies
    cookies = load_cookies_from_file()
    if not cookies:
        return
    
    # 创建会话并设置
    session = requests.Session()
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
    }
    
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])
    
    try:
        # 访问主页面
        base_url = "https://manaba.tsukuba.ac.jp"
        url = f"{base_url}/ct/course_3298888_surveyadm_examlist"
        response = session.get(url, headers=headers)
        response.raise_for_status()
        
        # 查找所有相关链接
        soup = BeautifulSoup(response.text, 'html.parser')
        links = soup.find_all('a', href=lambda x: x and ('surveyadm_gradetop'in x))
        unique_links = list({link['href']: link for link in links}.values())
        
        print(f"找到 {len(unique_links)} 个唯一链接")
        
        # 存储所有学生的提交信息
        all_students = {}
        
        # 处理每个链接
        for link in unique_links[1:]:
            href = link['href']
            full_url = f"{base_url}/ct/{href}" if not href.startswith('http') else href
            print(f"正在处理链接: {full_url}")
            
            # 获取并处理页面内容
            page_response = session.get(full_url, headers=headers)
            page_response.raise_for_status()
            process_submission_table(page_response.text, all_students, session, base_url, headers)
            
            # time.sleep(1)  # 避免请求过快
        
        # 保存结果
        save_to_excel(all_students)

    except Exception as e:
        print(f"发生错误: {str(e)}")

if __name__ == "__main__":
    main()