import pandas as pd



target = {"【none】":[]}

data = pd.read_excel('test.xlsx').values  #读取数据并转化成列表的形式


for val in data:    #一行一行的读取列表，就是一个一个单元格的读取数据
    cell = val[0].split("\n")   #把每个单元格内容利用换行符划分并保存为列表
    name = []
    num = []
    flag2 = 0
    i = 0
    i2 = len(cell)
    for key in target.keys():   #读取字典中的关键字
        i1 = i
        flag = 0    #记录在这个单元格能不能找到该关键字对应的内容
        for value in cell[i1 : i2]:  #一个一个的读取单元格划成的列表的内容
            i = i+1
            a = value.find("【")
            b = value.find("】")
            if a != -1 and b!=-1:
                name2 = value[a:b+1]
                if name2 not in target.keys():
                    name.append(name2)
                    l = len(target["【none】"])
                    if l > 0:
                        num.append([(len(target["【none】"])-1) * "", value])
                    else:
                        num.append([value])
                    flag2 = 1
            if value.find(key) != -1:   #如果该行文字存在某个关键字则给字典中关键字对应的列表添加该行文字
                target[key].append(value)
                flag = 1    #找到了就标志为1
                break    #这个关键字得到内容了，跳出这个寻找对应值的循环
        if flag == 0:    #如果找遍这个单元格每一行都不存在该关键字就给他添加一个空内容
            target[key].append("")
    if cell[-1].find("【") == -1 and cell[-1].find("】")==-1: #如果该单元格的最后一行不含有【】，则把【none】对应的""替换成该单元格最后一行
        target["【none】"][-1] = cell[-1]
    if flag2 == 1:
        mid_dic = dict(zip(name, num))
        target.update(mid_dic)
output = pd.DataFrame(target)
output.to_csv("message.csv", encoding = "utf_8_sig")   #将分类后的消息保存为csv
