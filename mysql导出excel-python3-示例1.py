import pymysql
import xlwt

#连接数据库函数1
def get_conn():  
    coon = pymysql.connect(user='root',passwd='root',db='test1',port=3306,host='127.0.0.1',charset='utf8')
    return coon

#执行查询数据函数2
def query_all(cur, sql, args):  
    cur.execute(sql, args)
    return cur.fetchall()

#导出测试用例步聚到export_to_excel_app_casestep函数4  https://www.cnblogs.com/x1you/p/12506369.html
def read_mysql_to_xlsx(filename): 
    list_table_head = ['id', '名称']  #定义表头
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('data', cell_overwrite_ok=True)
    for i in range(len(list_table_head)):
        sheet.write(0,i, list_table_head[i])
    conn = get_conn()   #调用连接数据库函数
    cur = conn.cursor()
    sql = 'SELECT d_id, d_name from test_dept where d_name != "部门1"'  #查询用例及步聚数据
    results = query_all(cur, sql, None)  #调用函数，定义记录查询到的数据
    conn.commit()
    cur.close()
    conn.close()
    row = 1
    for result in results:  #把结果循环写入到sheet
        col = 0
        print(type(result))
        print(result)
        for item in result:
            print(item)
            sheet.write(row, col, item)
            col += 1
        row += 1
    workbook.save(filename)  #保存到excel文件

if __name__ == '__main__':
    #调用导出用例步聚函数
    read_mysql_to_xlsx('export_to_excel_app_casestep.xls')
