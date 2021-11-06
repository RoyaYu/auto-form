import openpyxl
import xlrd
from pathlib import Path

title = ["cn"]
mem_dic = {} ##“key: cn  value: list[goods:price]”

path = Path('test_file', 'test_real.xlsx')
workbook= openpyxl.load_workbook(path,data_only=True) #排表
workbook.active

# for test, remove this
# dum_title = ["cn", "色纸","肾1","表情包","肾2"]
# dump_dic = {"踏雪":{"色纸":["khn1",10],"表情包":["an1",15]},
#             "兔":{"色纸":["mzk2",25]},
#             "延生":{"表情包":["len1rin1",30]}
#             }
# title = dum_title
# mem_dic = dump_dic




# reading part
def addSheet(worksheet, index):

    goods_name = worksheet["b1"].value
    title.append(goods_name)
    title.append("肾"+str(index))
    avg_price = worksheet["c1"].value
    for i in range(1, worksheet.max_row):
        diff = worksheet["a"+str(i+1)].value
        chara = worksheet["c"+str(i+1)].value
        print(chara)
        for col in worksheet.iter_cols(4, worksheet.max_column):
            cn = col[i].value
            if cn not in mem_dic:
                mem_dic[cn] = {goods_name:[[],"",0]}
            elif goods_name not in mem_dic[cn]:
                mem_dic[cn][goods_name] = [[],"",0]
            mem_dic[cn][goods_name][2] += avg_price + diff
            if (len(mem_dic[cn][goods_name][0]) == 0):
                mem_dic[cn][goods_name][0] = [chara, 1]
            else:
                mem_dic[cn][goods_name][0][1] += 1
        clean_up(mem_dic, goods_name)
        # print(mem_dic)

        # print(mem_dic)
            # addValue(avg_price, diff, chara, col[i])

def clean_up(mem_dic, goods_name):
    for i in enumerate(mem_dic):
        if goods_name in mem_dic[i[1]]:
            if len(mem_dic[i[1]][goods_name][0]) != 0:
                one_goods = str(mem_dic[i[1]][goods_name][0][0]) + str(mem_dic[i[1]][goods_name][0][1])
                mem_dic[i[1]][goods_name][1] += one_goods
                mem_dic[i[1]][goods_name][0] = []


# writing part
def write_sheet():
    price_sheet = workbook.create_sheet("肾表")
    # write title
    for til in enumerate(title):
        price_sheet.cell(column = til[0] + 1, row = 1, value = til[1])
    # write value
    for count in enumerate(mem_dic):
        row = count[0] + 2
        cn = count[1]
        price_sheet.cell(column = 1, row = row, value = cn) ## write cn
        personal_goods = mem_dic[cn]
        for goods_name,goods_price in personal_goods.items():
            col = title.index(goods_name) + 1
            price_sheet.cell(column = col, row = row, value = goods_price[0])
            price_sheet.cell(column = col+1, row = row, value = goods_price[1])
    # save file
    workbook.save("output.xlsx")


count = 0
for sheet in workbook.worksheets:
    count += 1
    addSheet(sheet,count)
for i in enumerate(mem_dic):
    for j in enumerate(mem_dic[i[1]]):
        mem_dic[i[1]][j[1]].pop(0)
print(1)
print(mem_dic)

write_sheet();
