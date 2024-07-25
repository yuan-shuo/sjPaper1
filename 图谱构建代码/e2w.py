import os
import pandas as pd
from w2neo import w2neo
# from py2neo import Graph

class ExcelProcessor:
    def __init__(self, folder_path):
        self.folder_path = folder_path

    def process_files(self):
        # al = set()
        # link = Graph("http://localhost:7474", auth=("neo4j", "174235"))
        # graph = link
        # query = "MATCH (u:CoreBook) RETURN u.name AS name"
        # result = graph.run(query)
        # al = {record["name"] for record in result}
        file_list = os.listdir(self.folder_path)

        for file_name in file_list:
            file_name_without_extension = os.path.splitext(file_name)[0]
            # if file_name_without_extension in al:
            #     continue
            if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
                file_path = os.path.join(self.folder_path, file_name)
                
                df = pd.read_excel(file_path, header=None)
                df.columns = ['column1', 'column2', 'column3', 'column4', 'column5', 'column6', 'column7', 'column8']
                result = [tuple(x) for x in df.values]

                print(f"正在处理：{file_name_without_extension}")
                # print(result)
                a = w2neo(result, file_name_without_extension)
                a.go()

if __name__=='__main__':
    # 使用ExcelProcessor类
    # 跑之前确保六列均没有空白格子（用查找直接全选后啥也不输直接点查找）
    # folder_path = 'E:\python_work\pyneo\pdfToData\myPDF\excelSQL'
    # folder_path = r"E:\python_work\pyneo\pdfToData\myPDF\forTest4"
    folder_path = r"E:\python_work\pyneo\pdfToData\myPDF\usedExcel"
    excel_processor = ExcelProcessor(folder_path)
    excel_processor.process_files()