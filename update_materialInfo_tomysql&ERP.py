'''
此脚本用途：电子元件CODE查询及新物料编码的录入

注意事项：
1.物料信息不包含：$ 和 @ 符号；编码默认为：N/A
2.文件加密，会乱码；打不开文件，需手动从模板文件<电子物料ERP录入模版>中复制粘贴来。
'''

import pymysql
import os
from time import time
from pymysql.cursors import DictCursor
from se_URL import gen_bro,login_erp,update_erp_material_list,logout_erp,logout_acount

Done_bom_path = r'D:\桌面文件\元件\newBom.csv'
Undone_bom_path = r'D:\桌面文件\元件\undone_Bom.csv'

class My_Mysql:
    host = 'localhost'  # 服务器地址
    port = 3306
    user = 'root'
    password = ''       # 密码为空
    database = 'material_info'

# 查询物料Part Number是否在数据库
def gen_sql_partNumber_dict(cursor):
    sql_partNumber_dict = {}
    select_partNumber_sql = "SELECT `Part Number`,code FROM material_list;"
    re_num = cursor.execute(select_partNumber_sql)
    result = cursor.fetchall()
    for re in result:
        sql_part_number = re.get('Part Number')
        sql_code = re.get('code')
        # print(spl_part_number)
        sql_partNumber_dict[sql_part_number] = sql_code
    return sql_partNumber_dict

# 生成物料编码；return new_code
def gen_Newcode(sql_partNumber_dict,type,cursor):
    sql_partNumber_dict = sql_partNumber_dict
    type_dic = {'LCR类': 'C97-', '半导体类': 'C98-', '其他类': 'C99-'}
    if type == '其他类':
        select_code_sql = "SELECT `code`,MAX(CODE) FROM material_list WHERE `code` LIKE 'C99%' GROUP BY `code`;"
    if type == '半导体类':
        select_code_sql = "SELECT `code`,MAX(CODE) FROM material_list WHERE `code` LIKE 'C98%' GROUP BY `code`;"
    if type == 'LCR类':
        select_code_sql = "SELECT `code`,MAX(CODE) FROM material_list WHERE `code` LIKE 'C97%' GROUP BY `code`;"
    re_num = cursor.execute(select_code_sql)
    result = cursor.fetchall()
    # print(result)
    max_code = result[-1].get('code')
    # print(max_code)
    num = int(max_code.split('-')[-1]) + 1
    # print(num)
    if len(str(num)) == 1:
        new_code = type_dic[type] + '000' + str(num)
    if len(str(num)) == 2:
        new_code = type_dic[type] + '00' + str(num)
    if len(str(num)) == 3:
        new_code = type_dic[type] + '0' + str(num)
    if len(str(num)) == 4:
        new_code = type_dic[type] + '' + str(num)
    # print(new_code)
    return new_code

# 创建mysql连接
def connect_mysql():
    mysql_conn = pymysql.connect(
        host=My_Mysql.host,  # 服务器地址
        port=My_Mysql.port,
        user=My_Mysql.user,
        password=My_Mysql.password,  # 密码为空
        database=My_Mysql.database
    )
    return mysql_conn

def run(target_data):
    # 将数据分割成列表
    data_list = target_data.strip().split('@')

    # 创建数据库连接
    print('连接mysql...')
    mysql_conn = connect_mysql()
    # 创建字典游标，查询的结果自动存在字典中
    cursor = mysql_conn.cursor(DictCursor)
    # 查询数据库的物料清单并转成规格：料号的字典
    sql_partNumber_dict = gen_sql_partNumber_dict(cursor)
    # print(sql_partNumber_dict)

    # 创建链接对象
    print('工作环境准备....')
    bro = gen_bro()

    # 登录ERP
    print('登录ERP...')
    wuliao = login_erp(bro)
    # 确认是否正确进入物料清单列表
    if wuliao != "物料":
        print('进错目录了，进入了：', wuliao)
        print('尝试重新启动程序！')
        return None

    # 打开文件
    fp2 = open(Done_bom_path, 'w', encoding='gbk')
    fp1 = open(Undone_bom_path, 'w', encoding='gbk')
    # 循环遍历出每个物料信息：item, code, name, type, Part_Number
    for data in data_list:
        if data.startswith('item') or data.startswith('Item'):
            continue
        # 分割数据信息
        lis = data.split('$')
        if len(lis) == 1:
            continue
        elif len(lis) == 7:
            kong, item, code, name, type, Part_Number, kong2 = data.split('$')
            print('待操作物料信息>>>', item, code, name, type, Part_Number)

        # 判断物料的Part_Number是否存在于数据库中
        # 若存在code= 数据库中的数据，否则仍未默认值
        for PN, CODE in sql_partNumber_dict.items():
            # 若物料存在数据库break
            if Part_Number in PN:
                print(Part_Number, ':已存在数据库，无需新增')
                code = CODE
                break

        # 若此时code仍未默认值'N/A'，需生成新的物料编码
        if code == 'N/A':
            print(Part_Number, '：无物料CODE')
            # 生成新的料号
            Newcode = gen_Newcode(sql_partNumber_dict, type,cursor)
            print('生成新的CODE:', Newcode)
            code = Newcode

            try:
                # 新CODE写入ERP
                print('新物料信息:', code, ',', Part_Number, ',开始写入ERP物料清单...')
                print('----------------------开始操作---------------------------')
                startin_time = time()
                update_erp_material_list(bro, code, type, Part_Number, name)
                print(code, '完成ERP录入，耗时：', time() - startin_time, '秒')
            except Exception as e:
                print(f'【 {Part_Number} 】写入ERP过程报错了：', e)
                fp1.write(item + ',' + code + ',' + type + ',' + name + ',' + Part_Number + '\n')
            else:
                # 若写入ERP过程未报错，将新code写入数据库
                print('新物料信息开始写入mysql:', code)
                newcode_sql = "INSERT INTO material_list(`code`,type,`Part Number`,name) VALUES( '%s', '%s', '%s', '%s');" % (
                    code, type, Part_Number, name)
                try:
                    re_num = cursor.execute(newcode_sql)
                    mysql_conn.commit()
                    print('完成录MySQL录入！')
                except Exception as e:
                    print(e, '录入MySQL失败')
                    print(code, '录入MySQL失败')
                    mysql_conn.rollback()

        # 将物料信息写入csv文件中
        print(f'{Part_Number}:写入CSV文件')
        fp2.write(item + ',' + code + ',' + type + ',' + name + ',' + Part_Number + '\n')
        print(f'----------------------{Part_Number}已结束操作---------------------------')

    # 循环结束，关闭文件
    print('生成完整的Bom，文件路径:', Done_bom_path)
    fp1.close()
    fp2.close()
    # 若不存在ERP写入失败的料号，则删掉Undone_bom_path文件
    if os.path.getsize(Undone_bom_path) == 0:
        os.remove(Undone_bom_path)
    else:
        print(r'有物料写入ERP过程报错了，见->D:\桌面文件\元件\undone_Bom.csv')
    # 退出erp账号
    logout_acount(bro)
    # 关闭bro
    logout_erp(bro)


if __name__ == '__main__':
    target_data = input('从电子物料ERP录入模版中复制待编码或录入的物料信息>>>>>>>>>>')
    # 执行主程序
    run(target_data)

