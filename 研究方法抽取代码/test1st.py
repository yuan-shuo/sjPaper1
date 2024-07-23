import pandas as pd  
  
def read_abstracts_from_excel(file_path):  
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
  
# 示例用法  
file_path = 'forTest2.xlsx'  # 替换为你的Excel文件路径  
abstracts = read_abstracts_from_excel(file_path)  
print(abstracts)