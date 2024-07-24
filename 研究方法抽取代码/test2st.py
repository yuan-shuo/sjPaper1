def allFromExcel(file_path):  
    """  
    从Excel文件中读取'Abstract'列的所有数据。  
      
    参数:  
    file_path (str): Excel文件的路径。  
      
    返回:  
    list of str: 包含每行'Abstract'数据的列表。  
    """  
    # 使用pandas读取Excel文件  
    df = pd.read_excel(file_path)  
      
    # 假设'Abstract'列存在于DataFrame中  
    # 如果列名不同，请相应地更改它  
    abstracts = df['Abstract'].tolist()  
      
    # 返回包含所有'Abstract'数据的列表  
    return abstracts

def find_matches_in_text(text, words_list):  
    """  
    在文本中查找列表内的元素，不区分大小写比对。  
      
    参数:  
    text (str): 要搜索的文本。  
    words_list (list of str): 要在文本中查找的单词列表。  
      
    返回:  
    list of str: 匹配成功的字符串列表（在原始列表中）。  
    """  
    # 将文本转换为小写，以便不区分大小写地搜索  
    text_lower = text.lower()  
      
    # 初始化一个空列表来存储匹配到的单词  
    matches = []  
      
    # 遍历单词列表  
    for word in words_list:  
        # 将单词也转换为小写，以便进行比较  
        word_lower = word.lower()  
          
        # 如果文本包含该单词（不区分大小写），则添加到匹配列表中  
        if word_lower in text_lower:  
            matches.append(word)  
      
    # 返回匹配到的单词列表  
    return matches

import os  
import pandas as pd  
  
# def read_articles_from_all_excel_files(folder_path):  
#     """  
#     从指定文件夹路径下的所有Excel文件中读取'Article Title'、'DOI'和'Abstract'列的数据。  
#     非字符串类型的数据将被转换为空字符串。  
  
#     参数:  
#     folder_path (str): 包含Excel文件的文件夹路径。  
  
#     返回:  
#     list of dicts: 每个字典包含一行数据的'Article Title'、'DOI'和'Abstract'。  
#     """  
#     all_articles = []  
  
#     for filename in os.listdir(folder_path):  
#         if filename.endswith('.xlsx') or filename.endswith('.xls'):  
#             file_path = os.path.join(folder_path, filename)  
  
#             try:  
#                 df = pd.read_excel(file_path)  
  
#                 # 检查所需的列是否存在  
#                 required_columns = ['Article Title', 'DOI', 'Abstract']  
#                 missing_columns = [col for col in required_columns if col not in df.columns]  
#                 if missing_columns:  
#                     print(f"Warning: '{filename}' missing columns: {', '.join(missing_columns)}")  
#                     continue  
  
#                 # 转换非字符串类型为空字符串  
#                 for col in required_columns:  
#                     df[col] = df[col].fillna('').astype(str)  
  
#                 # 将DataFrame转换为字典列表  
#                 articles = df[required_columns].to_dict(orient='records')  
#                 all_articles.extend(articles)  
  
#             except Exception as e:  
#                 print(f"Error reading '{filename}': {e}")  
  
#     return all_articles

# try1
# def read_articles_from_all_excel_files(folder_path):
#     """ 
#     从指定文件夹路径下的所有Excel文件中读取'Article Title'、'DOI'、'Abstract'和'Times Cited, All Databases'列的数据。  
#     非字符串类型的数据将被转换为空字符串。
  
#     参数:  
#     folder_path (str): 包含Excel文件的文件夹路径。
  
#     返回:  
#     list of dicts: 每个字典包含一行数据的'Article Title'、'DOI'、'Abstract'和'Times Cited, All Databases'。  
#     """  
#     all_articles = []  
#     required_columns = ['Article Title', 'DOI', 'Abstract', 'Times Cited, All Databases']
  
#     for filename in os.listdir(folder_path):  
#         if filename.endswith('.xlsx') or filename.endswith('.xls'):  
#             file_path = os.path.join(folder_path, filename)  
#             try:  
#                 df = pd.read_excel(file_path)  
#                 # 检查所需的列是否存在  
#                 missing_columns = [col for col in required_columns if col not in df.columns]  
#                 if missing_columns:  
#                     print(f"Warning: '{filename}' missing columns: {', '.join(missing_columns)}")  
#                     continue  
                
#                 # 转换非字符串类型为空字符串  
#                 for col in required_columns:  
#                     df[col] = df[col].fillna('').astype(str)  
                
#                 # 将DataFrame转换为字典列表  
#                 articles = df[required_columns].to_dict(orient='records')  
#                 all_articles.extend(articles)  
                
#             except Exception as e:  
#                 print(f"Error reading '{filename}': {e}")  
#     return all_articles

# try2
def read_articles_from_all_excel_files(folder_path):
    """ 
    从指定文件夹路径下的所有Excel文件中读取'Article Title'、'DOI'、'Abstract'和'Times Cited, All Databases'列的数据。  
    非字符串类型的数据将被转换为空字符串，'Times Cited, All Databases'列的值将转换为数值类型。
  
    参数:  
    folder_path (str): 包含Excel文件的文件夹路径。
  
    返回:  
    list of dicts: 每个字典包含一行数据的'Article Title'、'DOI'、'Abstract'和'Times Cited, All Databases'。  
    """  
    all_articles = []  
    required_columns = ['Article Title', 'DOI', 'Abstract', 'Times Cited, All Databases']
  
    for filename in os.listdir(folder_path):  
        if filename.endswith('.xlsx') or filename.endswith('.xls'):  
            file_path = os.path.join(folder_path, filename)  
            try:  
                df = pd.read_excel(file_path)  
                # 检查所需的列是否存在  
                missing_columns = [col for col in required_columns if col not in df.columns]  
                if missing_columns:  
                    print(f"Warning: '{filename}' missing columns: {', '.join(missing_columns)}")  
                    continue  
                
                # 转换非字符串类型为空字符串  
                for col in required_columns:  
                    df[col] = df[col].fillna('').astype(str)
                
                # 将'Times Cited, All Databases'列的值转换为数值类型
                df['Times Cited, All Databases'] = pd.to_numeric(df['Times Cited, All Databases'], errors='coerce')
                
                # 将DataFrame转换为字典列表  
                articles = df[required_columns].to_dict(orient='records')  
                all_articles.extend(articles)  
                
            except Exception as e:  
                print(f"Error reading '{filename}': {e}")  
    return all_articles

def list_to_dict_with_empty_lists(input_list):  
    """  
    将输入列表转换为一个字典，其中键是列表中的元素，值是一个空列表。  
  
    参数:  
    input_list (list): 输入的列表。  
  
    返回:  
    dict: 键为输入列表元素，值为空列表的字典。  
    """  
    # 使用字典推导式来创建字典  
    return {item: [] for item in input_list}

import pandas as pd

# def create_excel_tables(data_dict):  
#     # 初始化两个DataFrame来存储DOI和Article Title  
#     df_doi = pd.DataFrame()  
#     df_title = pd.DataFrame()  
  
#     # 遍历字典  
#     for key, value_list in data_dict.items():  
#         # 如果value_list不为空  
#         if value_list:  
#             # 提取DOI和Article Title  
#             dois = [item.get('DOI', '') for item in value_list]  
#             titles = [item.get('Article Title', '') for item in value_list]  
  
#             # 创建一个临时DataFrame，并将列名设置为当前键  
#             temp_doi = pd.DataFrame(dois, columns=[key])  
#             temp_title = pd.DataFrame(titles, columns=[key])  
  
#             # 将临时DataFrame追加到主DataFrame  
#             df_doi = pd.concat([df_doi, temp_doi], axis=1)  
#             df_title = pd.concat([df_title, temp_title], axis=1)  
  
#     # 写入Excel文件  
#     with pd.ExcelWriter('output_dois.xlsx', engine='openpyxl') as writer:  
#         df_doi.to_excel(writer, sheet_name='DOIs')  
  
#     with pd.ExcelWriter('output_titles.xlsx', engine='openpyxl') as writer:  
#         df_title.to_excel(writer, sheet_name='Article Titles') 

# try2!
def create_excel_tables(data_dict):  
    # 初始化两个DataFrame来存储DOI和Article Title  
    df_doi = pd.DataFrame()  
    df_title = pd.DataFrame()  
    df_times_cited = pd.DataFrame()
  
    # 遍历字典  
    for key, value_list in data_dict.items():  
        # 如果value_list不为空  
        if value_list:  
            # 提取DOI和Article Title  
            dois = [item.get('DOI', '') for item in value_list]  
            titles = [item.get('Article Title', '') for item in value_list]  
            times_cited = [item.get('Times Cited, All Databases', 0) for item in value_list]  # 假设不存在的引用次数为0
            print(times_cited)
  
            # 创建一个临时DataFrame，并将列名设置为当前键  
            temp_doi = pd.DataFrame(dois, columns=[key])  
            temp_title = pd.DataFrame(titles, columns=[key])  
            temp_times_cited = pd.DataFrame(times_cited, columns=[key])
  
            # 将临时DataFrame追加到主DataFrame  
            df_doi = pd.concat([df_doi, temp_doi], axis=1)  
            df_title = pd.concat([df_title, temp_title], axis=1)  
            df_times_cited = pd.concat([df_times_cited, temp_times_cited], axis=1)
  
    # 写入Excel文件  
    with pd.ExcelWriter('output_dois.xlsx', engine='openpyxl') as writer:  
        df_doi.to_excel(writer, sheet_name='DOIs')  
  
    with pd.ExcelWriter('output_titles.xlsx', engine='openpyxl') as writer:  
        df_title.to_excel(writer, sheet_name='Article Titles') 

    with pd.ExcelWriter('output_times_cited.xlsx', engine='openpyxl') as writer:
        df_times_cited.to_excel(writer, sheet_name='Times Cited, All Databases')

import heapq

def get_top_methods(method_counts, topN):
    # 使用heapq.nlargest来获取值最大的前15个键
    top_methods = heapq.nlargest(topN, method_counts, key=method_counts.get)
    return top_methods 

from openpyxl import Workbook

def create_excel_from_list(elements):
    # 创建一个新的工作簿
    wb = Workbook()
    # 选择默认的工作表
    ws = wb.active
    
    # 遍历列表元素，并将它们写入Excel的每一行
    for index, element in enumerate(elements):
        ws.cell(row=index + 1, column=1).value = element
    
    # 保存工作簿到文件
    wb.save('topMethod.xlsx')

def sort_by_times_cited(citation_list):
    # 使用sorted函数和lambda表达式进行排序
    # key参数指定排序依据，reverse=True表示降序排序
    sorted_list = sorted(citation_list, key=lambda x: x['Times Cited, All Databases'], reverse=True)
    return sorted_list

# 示例用法  
# folder_path = 'E:\yuan\work\sjPaper\队列\excel2test'  # 替换为你的文件夹路径  
folder_path = 'E:\yuan\work\sjPaper\队列\excel\wos'  # 替换为你的文件夹路径  
abstracts_lists = read_articles_from_all_excel_files(folder_path)
# print(1)
# print(abstracts_lists)
# print(1)
# 词表
words_list = allFromExcel('forTest2.xlsx')
# words_list = ["LIFE Child study"]
# 各词的统计列表, 列表内含字典
result_dict = list_to_dict_with_empty_lists(words_list) 
# 总计
total_num = len(abstracts_lists)
total_suc = 0
total_fal = 0
fal_list = []
for text in abstracts_lists:
    matches = find_matches_in_text(text['Abstract'], words_list)
    if matches:
        total_suc += 1
        for i in matches:
            result_dict[i].append({'Article Title': text['Article Title'], 'DOI': text['DOI'], 'Times Cited, All Databases': text['Times Cited, All Databases']})
    elif (not matches) and text['Abstract'] != '':
        total_fal += 1
        fal_list.append(text['Article Title'])
    else:
        if total_num > 1:
            total_num -= 1
        else:
            pass

# 各方法计数器
method_numDict = {}

# print(f"成功率：{(total_suc/total_num)*100}%")
# print(f"失败率：{(total_fal/total_num)*100}%")
# # print(result_dict)
tt_num = 0
for k, v in result_dict.items():
    k_num = len(v)
    # print(f"{k}对应论文共{k_num}篇")
    method_numDict[k] = k_num
# print(f"首个未录入论文：{fal_list[:20]}")
# print(f"成功数：{total_suc}")

# print(result_dict)

# 前15键名列表
lisNumtop = get_top_methods(method_numDict, 15)
print(lisNumtop)
create_excel_from_list(lisNumtop)
# 前15复制字典
dicResTop = {}
for i in lisNumtop:
    dicResTop[i] = sort_by_times_cited(result_dict[i])

create_excel_tables(dicResTop)

# print(dicResTop)
for k, v in dicResTop.items():
    k_num = len(v)
    tt_num += k_num
    print(f"{k}对应论文共{k_num}篇")

print(f"目前总计{tt_num}篇")