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
  
# 示例用法  
words_list = ["LIFE Child study", "world"]  
result_dict = list_to_dict_with_empty_lists(words_list)  
print(result_dict)