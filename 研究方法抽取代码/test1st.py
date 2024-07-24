def sort_by_times_cited(citation_list):
    # 使用sorted函数和lambda表达式进行排序
    # key参数指定排序依据，reverse=True表示降序排序
    sorted_list = sorted(citation_list, key=lambda x: x['Times Cited, All Databases'], reverse=True)
    return sorted_list

# 示例使用
if __name__ == "__main__":
    # 假设有以下列表
    citations = [
        {'Times Cited, All Databases': 5, 'Title': 'Paper A'},
        {'Times Cited, All Databases': 10, 'Title': 'Paper B'},
        {'Times Cited, All Databases': 3, 'Title': 'Paper C'}
    ]
    
    # 调用函数并打印结果
    sorted_citations = sort_by_times_cited(citations)
    print(sorted_citations)