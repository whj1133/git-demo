# -*- coding: utf-8 -*-

try:
    from xmlrpc import client as xmlrpclib
except ImportError:
    import xmlrpclib

ip = '127.0.0.1:8069'
user = 'admin'  # 用户名
pwd = '123456'  # 密码
dbname = 'erp14-xb'  # 数据库

sock = xmlrpclib.ServerProxy('http://%s/xmlrpc/common' % ip)
uid = sock.login(dbname, user, pwd)
sock = xmlrpclib.ServerProxy('http://%s/xmlrpc/object' % ip)

import xlrd

index = 0

xlsfile = "C:\\Users\\wang39\\Desktop\\zjxb_wl.xlsx"
book = xlrd.open_workbook(xlsfile)
sheet1 = book.sheet_by_index(index)
sheet_name1 = book.sheet_names()[index]
table = book.sheet_by_index(0)

nrows = sheet1.nrows  # 获取行总数
ncols = sheet1.ncols  # 获取列总数

start = 0
for row in range(nrows):
    if table.row_values(row)[0] == "物料编码":
        start = row + 1
        break

name = 1
if start:
    for row in range(start, nrows):
        name += 1
        material_info = dict()
        material_info['default_code'] = table.row_values(row)[0] or ''
        material_info['name'] = table.row_values(row)[1] or ''
        # name = table.row_values(row)[5] or ''
        material_info['huo_hao'] = table.row_values(row)[2] or ''
        material_info['gui_ge'] = table.row_values(row)[3] or ''
        material_info['chang_jia'] = table.row_values(row)[4] or ''
        material_info['wen_du'] = table.row_values(row)[8] or ''
        # if table.row_values(row)[0]=='Y.X.001' or table.row_values(row)[0]=='Y.X.002' or table.row_values(row)[0]=='Y.X.003' or table.row_values(row)[0]=='Y.X.004':

        if table.row_values(row)[7]:
            type_id = sock.execute_kw(dbname, uid, pwd, 'product.category', 'search',
                                      [[('name', '=', table.row_values(row)[7])]])
            print(type_id)
    #         material_info.update({"categ_id": type_id[0]})
    #
    #         material = sock.execute_kw(dbname, uid, pwd, 'product.template', 'create', [material_info])
    #     else:
    #         material = sock.execute_kw(dbname, uid, pwd, 'product.template', 'create', [material_info])
    #     print(material_info)
    # print(name)

