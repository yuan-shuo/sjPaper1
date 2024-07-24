# import pandas as pd
# import os

# # 读取Excel文件
# excel_path = 'output_dois.xlsx'  # Excel文件的路径
# df = pd.read_excel(excel_path)

# # 获取所有列名
# columns = df.columns

# # 指定文件夹的根路径
# root_folder_path = 'E:\\临时\\自动word部分内容改拟真手写体宏封包\\队列\\人群和指标论文'

# # 定义一个函数来清理列名，使其成为合法的文件夹名称
# def clean_column_name(column_name):
#     # 替换掉Windows系统中不允许的字符
#     invalid_chars = [':','*','?','"','<','>','|']
#     for char in invalid_chars:
#         column_name = column_name.replace(char, '_')
#     return column_name

# # 创建所有列名的文件夹
# for column in columns:
#     cleaned_column_name = clean_column_name(column)
#     column_folder_path = os.path.join(root_folder_path, cleaned_column_name)
    
#     # 检查文件夹是否存在，如果不存在则创建
#     if not os.path.exists(column_folder_path):
#         os.makedirs(column_folder_path)
    
#     # 读取对应列的前两行DOI
#     dois = df[column].head(2).tolist()
    
#     # 在每个文件夹中创建记事本文件
#     note_path = os.path.join(column_folder_path, 'doi.txt')
#     with open(note_path, 'w', encoding='utf-8') as file:
#         for doi in dois:
#             # 确保DOI是字符串类型
#             doi_str = str(doi)
#             file.write(doi_str + '\n')

# print('所有DOI文件已生成。')