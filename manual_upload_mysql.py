'''
将手动添加ERP的物料，加到文件；‘D:\桌面文件\元件\ERP电子元件物料清单_脚本更新版.csv’中后，更新至mysql
'''

import pymysql
from time import time
from pymysql.cursors import DictCursor
from se_URL import gen_bro,login_erp,update_erp_material_list,logout_erp


mysql_conn = pymysql.connect(
        host = 'localhost',  # 服务器地址
        port = 3306,
        user = 'root',
        password = '',  # 密码为空
        database = 'material_info'
    )
# 创建字典游标，查询的结果自动存在字典中
cursor = mysql_conn.cursor(DictCursor)

# 查询数据库的信息并放入字典sql_partNumber_dict中!不好使
sql_partNumber_dict = {}
select_partNumber_sql = "SELECT `Part Number`,code FROM material_list;"
re_num = cursor.execute(select_partNumber_sql)
result = cursor.fetchall()
for re in result:
    sql_part_number = re.get('Part Number')
    sql_code = re.get('code')
    # print(spl_part_number)
    sql_partNumber_dict[sql_part_number] = sql_code

# 手动写入:往mysql里写入已存在的含CODE的物料清单
sql_path = r'D:\桌面文件\元件\ERP电子元件物料清单_脚本更新版.csv'
with open(sql_path, 'r', encoding='gbk') as fp:
    data_line = fp.readlines()
    for data in data_line:
        if data.startswith('item') or data.startswith('Item'):
            continue
        data_list = data.replace('\n', '').split(',')
        print(data_list)
        item, code, type, Part_Number = data_list
        for PN, CODE in sql_partNumber_dict.items():
            if Part_Number in PN:
                print(Part_Number, ':已存在数据库，无需新增')
            else:
                sql = "INSERT INTO material_list(item,`code`,type,`Part Number`) VALUES('%s', '%s', '%s', '%s')"%(item, code, type, Part_Number)

                try:
                    re_num = cursor.execute(sql)
                    mysql_conn.commit()
                    # print('完成录入个数：',re_num)
                except Exception as e:
                    print(e)
                    mysql_conn.rollback()
