import bacdive
import re
import openpyxl
import json
import requests
from bs4 import BeautifulSoup
import os

# 自定义属名--------------------------------------------------------------------------------------------------
genus = input("请输入属的英文名（例如丹毒丝菌属，则输入 Erysipelothrix ）：").lower()



# 爬虫-----------------------------------------------------------------------------------------------------------------

def worm():
    # 抓取属网址中的种链接，获取BacDive Number--------------------------------------------------------------------------------
    ids = []
    pubs = []
    dois = []

    # 获取网页的HTML内容
    url = f"https://lpsn.dsmz.de/genus/{genus}"
    response = requests.get(url)
    html = response.text

    # 解析HTML内容
    soup = BeautifulSoup(html, 'html.parser')

    # 找到表格
    table = soup.find('table', {'class': 'detail-table sortable'})

    # 遍历表格的每一行
    for row in table.find_all('tr'):
        # 获取该行的所有<td>标签
        cells = row.find_all('td')

        # 检查第三个<td>标签的文本是否包含"correct name"
        if len(cells) > 2 and cells[2].text.strip() == 'correct name':
            # 找出该行中的<a>标签
            link = row.find('a')

            # 获取链接对应的网页HTML内容
            href = link.get('href')
            if href.startswith('/'):
                href = 'https://lpsn.dsmz.de' + href
            response = requests.get(href)
            html = response.text

            # 对每个链接，获取链接对应的网页HTML内容
            # 解析HTML内容
            soup = BeautifulSoup(html, 'html.parser')

            # 查找class为"bacdive-link"的<p>标签
            bacdive_link = soup.find('p', {'class': 'bacdive-link'})
            ijsem_list = soup.find('b', string='IJSEM list: ')
            doi_link = soup.find('a', {'class': 'doi-link'})

            if bacdive_link:
                # 查找其中的数字
                bacdive_urls = bacdive_link.find_all('a')
                bacdive_url = bacdive_urls[1]['href']
                number = re.search(r'/(\d+)/?$', bacdive_url)

                if number:
                    ids.append(number.group(1))
                    print(number.group(1))

                    if ijsem_list:
                        i_tag = ijsem_list.find_next_sibling()
                        if i_tag:
                            pubs.append(i_tag.string)
                            print(i_tag.string)
                            print(doi_link['href'])
                            dois.append(doi_link['href'])
                            print("-------------------------------------------------------------------------------")
                        else:
                            print("No <i> tag found after IJSEM list in", href)
                    else:
                        print("No IJSEM list found in", href)
                        pubs.append("NA")
                        dois.append('NA')
                        print("-------------------------------------------------------------------------------")

                else:
                    print("No number found in", href)
            else:
                print("No bacdive-link found in", href)
                print("-------------------------------------------------------------------------------")


    id_string = ";".join(ids)
    pub_string = ";".join(pubs)
    doi_string = ";".join(dois)
    # print(id_string)
    # print(pub_string)

    file = open(f"{genus}-id.txt", "w")
    # 将字符串写入文件
    file.write(id_string)
    # 关闭文件
    file.close()

    file = open(f"{genus}-pub.txt", "w")
    # 将字符串写入文件
    file.write(pub_string)
    # 关闭文件
    file.close()

    file = open(f"{genus}-doi.txt", "w")
    # 将字符串写入文件
    file.write(doi_string)
    # 关闭文件
    file.close()


file_path = f"./{genus}-id.txt"

# 检查文件是否存在
if not os.path.exists(file_path):
    # 如果文件不存在，执行函数
    print("##########################开始执行爬虫程序##########################")
    worm()
else:
    # 如果文件存在，不执行函数
    print(f"{genus}-id.txt文件已存在，不再执行爬虫程序，开始获取种信息")



# 访问BacDive API--------------------------------------------------------------------------------------------

client = bacdive.BacdiveClient('zylhuang@outlook.com', 'Zyl@200166')

## search with a BacDive ID
with open(f'{genus}-id.txt', 'r') as file:
    # 使用'read()'方法读取文件内容
    id_string = file.read()


with open(f'{genus}-pub.txt', 'r') as file:
    pub_string = file.read()
# 使用分号分割字符串
pubs = pub_string.split(';')


with open(f'{genus}-doi.txt', 'r') as file:
    doi_string = file.read()
# 使用分号分割字符串
dois = doi_string.split(';')




# print(id_string)

client.search(id=id_string)

pattern_isolation = r'isolated from ([\w\s]+)\.'
pattern_name = r'full scientific name\': \'(.*?)\''
pattern_culture = r'\'name\': \'(.*?)\''
pattern_temper = r'\'temperature\': \'(.*?)\''
pattern_ph = r'\'culture pH\': \[(.*?)\]'
pattern_oxygen = r'\'oxygen tolerance\': \'(.*?)\''
pattern_bacDive = r'\'BacDive-ID\': (\d+)'
pattern_DSM = r'\'DSM-Number\': (\d+)'


data = [
        ['种名', "发表期刊", '分离位置', '培养基', '温度', 'pH', '氧适应性', 'DOI', 'BacDive', 'DSM Number']  
    ]

index = 0

## retrieval of data for the strain previously searched
print("##########################开始生成Excel表格##########################")
for strain in client.retrieve():

    row = []

    # print(strain)

    # 将字典转换为字符串
    strain_str = str(strain)

    # 种名匹配
    match_name = re.search(pattern_name, strain_str)
    # 检查是否找到匹配项并提取结果
    if match_name:
        extracted_text = match_name.group(1)
        extracted_text = extracted_text.replace('<I>', '')
        extracted_text = extracted_text.replace('</I>', '')
        print("Name: " + extracted_text)  
        row.extend([extracted_text])
    else:
        print("未找到name")
        row.extend(['NA'])


    row.extend([pubs[index]])
    

    # 分离位置匹配
    match_isolation = re.search(pattern_isolation, strain_str)
    # 检查是否找到匹配项并提取结果
    if match_isolation:
        extracted_text = match_isolation.group(1)
        print("Isolated From: " + extracted_text)
        row.extend([extracted_text])
    else:
        print("未找到isolation")
        row.extend(['NA'])

    match_culture = re.findall(pattern_culture, strain_str)
    # 检查是否找到匹配项并提取结果
    if match_culture:
        extracted_text = match_culture 
        string = ', '.join(extracted_text)
        print("Culture Medium: " + string)
        row.extend([string])
    else:
        print("未找到culture medium")
        row.extend(['NA'])

    match_temper = re.search(pattern_temper, strain_str)
    # 检查是否找到匹配项并提取结果
    if match_temper:
        extracted_text = match_temper.group(1)
        row.extend([extracted_text])
        print("Temperature: " + extracted_text)
    else:
        print("未找到Temperature")
        row.extend(['NA'])

    match_ph = re.search(pattern_ph, strain_str)
    # 检查是否找到匹配项并提取结果
    if match_ph:
        extracted_text = match_ph.group(1)
        # 将字符串分割成多个部分
        parts = extracted_text.split('}, ')
        # 将每一部分转为字典，并添加到列表中
        tempo = []
        for part in parts:
            # 如果不是最后一部分，需要添加回被split移除的'}'
            if part != parts[-1]:
                part += '}'
            # 将字符串转为字典
            tempo.append(json.loads(part.replace("'", "\"")))

        result = []
        for item in tempo:
            result.append(f"{item['type']}: {item['pH']}")

        output = '; '.join(result)
        print("pH: " + output)
        row.extend([output])
    else:
        print("未找到ph值")
        row.extend(['NA'])

    match_oxygen = re.findall(pattern_oxygen, strain_str)
    # 检查是否找到匹配项并提取结果
    if match_oxygen:
        extracted_text = match_oxygen 
        string = ', '.join(extracted_text)
        print("Oxygen Tolerance: " + string)
        row.extend([string])
    else:
        print("未找到oxygen tolerance")
        row.extend(['NA'])


    row.extend([dois[index]])



    match_bacDive = re.search(pattern_bacDive, strain_str)
    # 检查是否找到匹配项并提取结果
    if match_bacDive:
        extracted_text = match_bacDive.group(1)
        print("BacDive ID: " + extracted_text)  
        row.extend([extracted_text])
    else:
        print("未找到bacdive ID")
        row.extend(['NA'])

    match_DSM = re.search(pattern_DSM, strain_str)
    # 检查是否找到匹配项并提取结果
    if match_DSM:
        extracted_text = match_DSM.group(1)
        print("DSM Number: " + extracted_text)  
        row.extend([extracted_text])
    else:
        print("未找到DSM")
        row.extend(['NA'])


    index += 1

    data.append(row)

    print("-------------------------------------------------------------------------")


# print(data)
# 生成Excel表格---------------------------------------------------------------------------------------
# 创建一个新的 Excel 工作簿
workbook = openpyxl.Workbook()

# 获取当前活动的工作表（默认情况下，新工作簿只有一个工作表）
worksheet = workbook.active

# 将数据写入工作表
for row in data:
    worksheet.append(row)

workbook.save(f'{genus}.xlsx')


print("@@@@@@@@@@@@@@@@@@程序结束@@@@@@@@@@@@@@@@@@")
print(f"请在程序的同一目录下，查找{genus}.xlsx文件来查看执行结果")
print(f"{genus}-id.txt、{genus}-doi.txt和{genus}-pub.txt为缓存文件，使用完后可删除")
os.system("pause")