import os  
import pandas as pd  
from docx import Document  
  
def extract_data_from_docx(file_path):  
    """从Word文档中提取数据"""  
    data = {}  
    doc = Document(file_path)  
    for para in doc.paragraphs:  
        if para.text.startswith('C'):  # 假设编号总是以C开头  
            parts = para.text.split('，')  
            if len(parts) >= 5:  
                data['编号'] = parts[0]  
                data['姓名'] = parts[1]  
                data['性别'] = parts[2]  
                data['年龄'] = parts[3]  
                data['日期'] = parts[4]  
                data['采访地点'] = ''  # 可能需要更复杂的逻辑来提取  
                  
        # 假设接下来的几段分别包含观察、找寻、思考和畅想  
        if '观察：' in para.text:  
            data['观察'] = para.text.split('观察：', 1)[1].strip()  
        elif '找寻：' in para.text:  
            data['找寻'] = para.text.split('找寻：', 1)[1].strip()  
        elif '思考：' in para.text:  
            data['思考'] = para.text.split('思考：', 1)[1].strip()  
        elif '畅想：' in para.text:  
            data['畅想'] = para.text.split('畅想：', 1)[1].strip()  
              
    return data  
  
def process_files_in_directory(directory):  
    """处理指定目录下的所有Word文档"""  
    results = []  
    for filename in os.listdir(directory):  
        if filename.endswith('.docx'):  
            file_path = os.path.join(directory, filename)  
            data = extract_data_from_docx(file_path)  
            if data:  
                results.append(data)  
      
    # 将结果保存为DataFrame  
    df = pd.DataFrame(results)  
      
    # 保存到Excel  
    excel_path = os.path.join(directory, 'output.xlsx')  
    df.to_excel(excel_path, index=False)  
  
# 更改为您的文档所在的目录  
directory_path = 'E:\临时\自动word部分内容改拟真手写体宏封包\自动小盆友\word'  
process_files_in_directory(directory_path)