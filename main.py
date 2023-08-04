import re
import pandas as pd
import openpyxl
import glob
import msvcrt
import os


# 文件预处理，去除前端无效信息
def dealtxt(infile):
    input_file = infile
    output_file = 'deal.txt'
    target_field = "# Detailed group assignments are listed below"

    with open(input_file, 'r') as f_in:
        lines = f_in.readlines()

    # 查找目标字段所在行的索引
    start_index = None
    for i, line in enumerate(lines):
        if target_field in line:
            start_index = i
            break

    # 截取目标字段之后的内容
    if start_index is not None:
        new_content = ''.join(lines[start_index + 1:])

        # 将新内容写入输出文件
        with open(output_file, 'w') as f_out:
            f_out.write(new_content)


# 打开文件，i为传入文件名
def open_file(i):
    filenames = i
    with open(filenames) as f:
        files = f.read()
    return files


# 根据字典生成Excel文件
def createxl(filenames):
    # 创建DataFrame
    df = pd.DataFrame.from_dict(filenames, orient='index')

    # 创建Excel Writer对象
    excel_writer = pd.ExcelWriter('example.xlsx', engine='openpyxl')

    # 将 DataFrame 写入 Excel 文件
    df.to_excel(excel_writer, index_label='OTU', header=True)

    # 保存 Excel 文件
    excel_writer.close()


# 修改第一列，第二列的列名
def rename():
    workbook = openpyxl.load_workbook('example.xlsx')
    worksheet = workbook['Sheet1']
    worksheet.cell(row=1, column=1, value='OTU')
    worksheet.cell(row=1, column=2, value='function')
    workbook.save('example.xlsx')


# 根据Excel文件创建字典
def createdict(filename):
    df = pd.read_excel(filename)
    # 创建一个空字典来存储转换后的数据
    data_dict = {}
    # 循环遍历DataFrame的每一行，将第一列作为键，第二列和第三列的值组成列表作为值
    for index, row in df.iterrows():
        key = row['key']
        value = [row['value1'], row['value2'], row['value3']]
        data_dict[key] = value
    return data_dict


# 去重并排序
def sort_remove(inputs, order):
    sort = []
    for i in order:
        if i in inputs:
            sort.append(i)
        else:
            pass
    return sort


# 去除无用OTU
def remove_otu(df):
    rows_to_remove = []
    for index, row in df.iterrows():
        if row[tax.columns[0]] not in otuall:
            rows_to_remove.append(index)
    df = df.drop(rows_to_remove)
    return df


# 建立OTU:功能的键值对
def tranfun(input):
    example = pd.read_excel(input)
    data_dict = {}
    for index, row in example.iterrows():
        data_dict[row['OTU']] = row['function']
    return data_dict


# dataframe转换为excel
def createdfxl(filenames):
    # 创建DataFrame
    df = pd.DataFrame(filenames)
    # 创建Excel Writer对象
    excel_writer = pd.ExcelWriter('example_tax.xlsx', engine='openpyxl')
    # 将 DataFrame 写入 Excel 文件
    df.to_excel(excel_writer, index=False)
    # 保存 Excel 文件
    excel_writer.close()


original_directory = input('请输入工作目录：')
# 获取当前路径下的所有文件夹列表
folders = [f for f in os.listdir(original_directory) if os.path.isdir(os.path.join(original_directory, f))]
if not folders:
    txt_name = glob.glob((os.path.join(original_directory, "*.txt"))) # 如果my_list为空，条件成立
    folders.append(txt_name[0])
for item in folders:
    os.chdir(original_directory)
    if len(folders) == 1:
        dealtxt(item)
    else:
        folder_path = os.path.join(original_directory, item)
        os.chdir(folder_path)  # 切换到文件夹
        txt_files = glob.glob("*.txt")
        txt = txt_files[0]  # 获取第一个txt文件
        dealtxt(txt)  # 对txt文件进行前处理输出为deal.txt

    files = open_file('deal.txt')  # 打开处理后的txt存入file

    # 正则表达式检索出所有的OTU并转换为list
    otuall = list(set(re.findall(r'OTU_[0-9]+', files)))
    # 切片筛选不同功能，以“#”分隔
    otufun = files.split('#')
    del otufun[0]  # 文件处理，删除第一行空白值

    # for循环删除不需要的OTU
    df = pd.read_excel('字段替换.xlsx')
    list1 = df['key'].tolist()
    otufun_to_remove = []
    for remove in otufun:
        if not any(item in remove for item in list1):
            otufun_to_remove.append(remove)
    for remove in otufun_to_remove:
        otufun.remove(remove)

    # 创建空白字典储存OTU：function的键值对
    otufuns = {}
    # 创建功能代码替换字典
    my_dict = createdict('字段替换.xlsx')
    # for循环筛选OTU功能
    order = ['C', 'H', 'O', 'N', 'S', 'Mn', 'Fe', 'As', 'OX', 'RX', 'FX']  # 功能顺序
    for o in otuall:
        key = o
        funf = []
        for otu in otufun:
            if o in otu:
                match = re.search(r'(\w+) \(\d+ records\)', otu)
                match = match.group(1)
                # 下方循环为将功能替换为相应代号
                if match in my_dict:
                    for item in my_dict[match]:
                        if isinstance(item, str):
                            funf.append(item)
            else:
                pass

        remove_string = sort_remove(funf, order)  # 排序并删除，生成list
        result_string = ''.join(remove_string)  # list转换为字符串
        otufuns[key] = result_string  # 将OTU：function的键值对存入字典otufuns

    createxl(otufuns)  # 将otufuns的字典转换为excel
    rename()  # 修改前两列的名字为“OTU”，”function”

    # .xls需转换为.xlsx
    taxname = glob.glob('*taxonomy*.xlsx')  # 使用 glob.glob() 查找带有“taxonomy”字段的xlsx文件
    tax = pd.read_excel(taxname[0])  # 此为存有taxonomy的dataframe
    tax = remove_otu(tax)  # 去除未被匹配到的OTU

    example = tranfun('example.xlsx')  # 建立OTU:功能的键值对

    tax['function'] = tax[tax.columns[0]].map(example)  # 将function与OTU匹配加入到tax中

    # 使用insert()方法将最后一列（function）插入到第2列的位置
    last_column = tax.pop(tax.columns[-1])
    tax.insert(1, last_column.name, last_column)
    # 使用insert()方法将最后一列（taxonomy）插入到第3列的位置
    last_column = tax.pop(tax.columns[-1])
    tax.insert(2, last_column.name, last_column)

    createdfxl(tax)  # tax转换为example_tax.xlsx

    tax.set_index(tax.columns[0], inplace=True)  # 将OTU设置为行索引

    # tax写入excel
    writer = pd.ExcelWriter('report.xlsx')
    result_tax = tax.groupby('function', as_index=False).sum()  # 以function作为标签分类并对OTU数量求和
    result_tax = result_tax.drop(columns='taxonomy')  # 删除taxonomy列
    result_tax.to_excel(writer, '汇总', index_label='NUM', header=True)  # 写入excel
    writer.close()  # 关闭excel
    os.chdir(original_directory)

print("运行成功，请按任意键退出~")
ord(msvcrt.getch())
