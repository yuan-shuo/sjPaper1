from py2neo import Node, Graph, Relationship

class w2neo:
    def __init__(self, byz, filename, delWarn="n", debug=False):
        self.wyz = byz
        self.debug = debug
        self.name = filename
        self.warn = delWarn
        self.existing_names = set()
        self.ner = {}
        self.rel = []

    # 启动！
    def go(self):
        if self.debug:
            print("目前以debug模式运行ing，仅打印")
        else:
            # 连接数据库并清空所有内容（要先启动数据库不然先注释掉）
            link = Graph("http://localhost:7474", auth=("neo4j", "174235"))
            self.graph = link

            # 执行查询以检索已存在的实体的名称
            query = "MATCH (u) RETURN u.name AS name"
            result = self.graph.run(query)
            
            # 从查询结果中提取名称并返回哈希表
            self.existing_names = {record["name"] for record in result}

            # warn = input('(__!慎重!__)是否删除原有数据库?(y/n)：')
            warn = self.warn

            if warn == 'y':
                # 删库！慎用
                self.graph.delete_all()
                print("已清空原有数据")
            else:
                print("将保留原有数据库")

            # 处理六元组
            self.w2dl()
            # 绘制节点
            self.plot_point()
            # 绘制联系
            self.plot_relation()

            print("构建任务完成！")
            

    # 画点
    def plot_point(self):
        # self.core = Node("CoreBook", name=self.name)
        # self.graph.create(self.core)
        for k, v in self.ner.items():
            if k in self.existing_names:
                # self.graph.nodes.match(name=str(k)).first().update({"source": "new"})
                # subject["source"] + ',' + self.name
                # rel = Relationship(subject, "content", self.core)
                # self.graph.create(rel)
                continue
            node = Node(str(v[0]), name=str(k), source=self.name, method=str(v[1]))
            # 绘制代码，debug时注释掉就行
            # print(f"debug:{k, v}")
            self.graph.create(node)
            # 链接主节点
            # rel = Relationship(node, "content", self.core)
            # self.graph.create(rel)

        print("节点绘制完成！")

    # 联系
    def plot_relation(self):
        for item in self.rel:
            subject = self.graph.nodes.match(name=item['subject']).first()
            object = self.graph.nodes.match(name=item['object']).first()
            properties={'judge':item['relProp']}
            
            rel = Relationship(subject, str(item['relation']), object, **properties)
            self.graph.create(rel)
        
        print("联系绘制完成！")

    # 六元组处理
    def w2dl(self):
        for wyz in self.wyz:
            self.ner[wyz[0]] = (wyz[1], wyz[2])
            self.ner[wyz[5]] = (wyz[6], wyz[7])
            self.rel.append({'subject': wyz[0], 'relation': wyz[3], 'relProp':wyz[4], 'object': wyz[5]})
            
        print("八元组处理完成！")
        # print(self.rel)


if __name__ == '__main__':
    list_wyz = [
        ("王石春", "人物",3, "作者",5,"GB 50218-94", "标准文件",8),
        ("邢念信", "人物",3, "作者",5,"GB 50218-94", "标准文件",8)
        ]
    a = w2neo(list_wyz, "测试规范2")
    a.go()

