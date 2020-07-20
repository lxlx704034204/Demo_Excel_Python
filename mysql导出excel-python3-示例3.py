from openpyxl import Workbook

# 向sheet中写入一行数据
def insertOne(key, sheet):
    row = [key] * 3
    sheet.append(row)

# 新建excel，并创建多个sheet
if __name__ == "__main__":

    book = Workbook()
    # 1 删除默认的sheet
    names = book.get_sheet_names()
    default = book.get_sheet_by_name(names[0])
    book.remove(default)
    # 2 新建自定义的sheet
    for i in range(0, 2):
        # 为每个sheet设置title，插入位置index
        sheet = book.create_sheet("sheet" + str(i), i)
        # 每个sheet里设置列标题
        sheet.append(["title" + str(i)] * 3)

    sheets = book.get_sheet_names()
    count = 0
    # 4 向sheet中插入数据
    for i in range(0, 10):
        insertOne("ni", book.get_sheet_by_name(sheets[1]))
        insertOne("wo", book.get_sheet_by_name(sheets[0]))
        insertOne("ta", book.get_sheet_by_name(sheets[1]))
        count = count + 1

    # 5 保存数据到.xlsx文件
    book.save("test.xlsx")
    print(str(count))
