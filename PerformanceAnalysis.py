# -*- coding:utf-8 -*-
import pandas as pd
import pymysql
from settings import Settings
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta
import pymysql
import win32com
from win32com.client import Dispatch
from sitecaculation import Site_caculation, Caculation
from sqlalchemy import create_engine
from ReadData import read_data

def isClearance(x, y, z, clearance_goods):
    if clearance_goods != None:
        for i in range(0, len(clearance_goods)):
            clearance_goods_each = clearance_goods[i]
            if (x >= clearance_goods_each[0]) and (y in clearance_goods_each):
                return '当月转清仓'
    else:
        return z

def instant_order_deal(plat, special_product, clearance_goods, new_product, orders):
    # 读取特定库龄并转为列表格式
    if plat == 'amazon':
        special_product = read_data().read_special_product(plat)

    # 目前只有wish要判断
    if plat == 'wish':
        newproduct = read_data().read_newproduct()  # 读取新品数据
    # 退款订单处理
    orders['订单总金额(包含客户运费、平台补贴)'] = orders.apply(lambda x: 0 if (x['订单类型'] == 'refund') else x['订单总金额(包含客户运费、平台补贴)'], axis=1)
    orders['平均采购价'] = orders.apply(lambda x: 0 if (x['订单类型'] == 'resend') else x['平均采购价'], axis=1)
    # 中英仓处理
    orders['仓库分类'] = orders.apply(lambda x: '中仓' if (x['发运仓库'] =='SH [上海奉贤仓]') | (x['发运仓库'] =='WZC [温州仓]') | (x['发运仓库'] =='SZC [深圳仓]') else '海外仓', axis=1)
    # orders['仓库分类'] = orders.apply(lambda x: np.where(((x['发运仓库'] =='SH [上海奉贤仓]') | (x['发运仓库'] =='WZC [温州仓]') | (x['发运仓库'] =='SZC [深圳仓]')), '中仓', '海外仓'), axis=1)
    # 当月转清仓处理
    orders['仓库分类'] = orders.apply(lambda x: isClearance(x['付款时间'], x['产品代码'], x['仓库分类'], clearance_goods), axis=1)
    # 处理新品
    if plat == 'wish':
        orders['仓库分类'] = orders.apply(lambda x: '新品' if (x['产品代码'] in newproduct) else x['仓库分类'], axis=1)
        # orders['仓库分类'] = orders.apply(lambda x: np.where((x['产品代码'] in newproduct), '新品', x['仓库分类']), axis=1)
    # 特定库龄处理
    if plat == 'amazon':
        orders['仓库分类'] = orders.apply(
            lambda x: '特定库龄' if ((x['发运仓库'] + x['产品代码']) in special_product) else x['仓库分类'], axis=1)
        # orders['仓库分类'] = orders.apply(lambda x: np.where(((x['发运仓库'] + x['产品代码']) in special_product), '特定库龄', x['仓库分类']), axis=1)
    else:
        orders['仓库分类'] = orders.apply(lambda x: '特定库龄' if (x['产品代码'] in special_product) else x['仓库分类'], axis=1)
        # orders['仓库分类'] = orders.apply(lambda x: np.where((x['产品代码'] in special_product), '特定库龄', x['仓库分类']), axis=1)
    # 处理好仓库分类，接下来判断是否是开发新品
    orders['开发新品'] = orders.apply(lambda x: '开发新品' if x['产品代码'] in new_product else '非开发新品', axis=1)
    # 然后处理货值
    orders['货值'] = orders['数量'] * orders['平均采购价']
    return orders

'''订单分配到人'''
def order_to_sale(plat, orders_deal, account):
    '''分配到人，以左边的表为准'''
    if plat == 'ebay':
        result = pd.merge(orders_deal, account, on='产品代码', how='left')
    else:
        result = pd.merge(orders_deal, account, on='平台账号', how='left')
    return result

def main():

    # 读取哪些平台需要计算业绩，ebay的站点分配情况
    plat_list = Settings().plat_list
    ebay_site = Settings().ebay_site
    # 本月订单处理
    # 定义一个大容器
    special_product = read_data().read_special_product('其他平台')
    # 读取当月转清仓数据
    clearance_goods = read_data().read_clearance_goods()
    # 读取本月开发新品数据
    new_product = read_data().read_new_product('本月')

    all_plat_result_instant = pd.DataFrame()
    for plat in plat_list:
        # 处理好仓库分类和货值的和是否是开发新品的订单
        # 读取该plat订单
        orders = read_data().read_order('instant_order', plat)
        if orders.empty == True:
            break
        orders = instant_order_deal(plat, special_product, clearance_goods, new_product, orders)
        if plat != 'ebay':
            account = read_data().read_account_sales(plat)
            result = order_to_sale(plat, orders, account)
        else:
            # 先定义一个容器
            result = pd.DataFrame()
            for site, value in ebay_site.items():
                # 读取订单信息
                data_site = orders[orders['站点'] == site]
                # 读取ebay产品分配信息
                product_sale = read_data().read_product_sale(value)
                result_site = order_to_sale(plat, data_site, product_sale)
                result = pd.concat([result, result_site], axis=0, sort=False)
        result.to_csv('个人业绩/' + plat + str(datetime.datetime.strftime(datetime.datetime.now(), '%Y%m%d')) + '.csv',
                      index=None, encoding='utf_8_sig')
        print(result)
        all_plat_result_instant = pd.concat([all_plat_result_instant, result], axis=0, sort=False)
    all_plat_result_instant['销售员'] = all_plat_result_instant['销售员'].fillna('未指定')
    all_plat_result_instant['销售员'] = all_plat_result_instant['销售员'].str.lower()

    all_plat_result_instant.to_csv('全平台业绩本月.csv', index=None, encoding='utf_8_sig')

    # 销售和仓库销售业绩
    wh_result, sale_result = Caculation().caculation_sale_instant(all_plat_result_instant)
    # 店铺业绩计算
    ress = Caculation().caculation_by_account(all_plat_result_instant)
    # 开发业绩核算
    result_developer = Caculation().caculation_developer(all_plat_result_instant)
    result_developer = result_developer.fillna(0)
    # 站点本期计算
    Site_caculation(all_plat_result_instant, 'w', '本期')
    ress.to_csv('店铺销售业绩.csv', encoding='utf_8_sig', index=None)
    wh_result.to_csv('各仓库销售业绩.csv', encoding='utf_8_sig', header=None)
    sale_result.to_csv('销售业绩本月.csv', index=None, encoding='utf_8_sig')
    result_developer.to_csv('开发业绩本月.csv', encoding='utf_8_sig')

    # 读取产品名称
    product_title = read_data().read_productTitle()
    # 加上产品名称
    all_plat_result_instant = pd.merge(all_plat_result_instant, product_title, on='产品代码', how='left')
    # 保存全平台业绩
    all_plat_result_instant.to_csv('全平台业绩本月.csv', index=None, encoding='utf_8_sig')
    # 顺道写入数据库,目前用来计算个人业绩
    order_instant(all_plat_result_instant)


    # 处理上月订单
    new_product = read_data().read_new_product('上月')
    new_product['开发新品'] = '开发新品'
    all_plat_result_last = pd.DataFrame()
    for plat in plat_list:
        # 处理好仓库分类和货值的和是否是开发新品的订单
        # 读取该plat订单
        orders = read_data().read_order('last_order', plat)
        orders = orders[orders['付款时间'] < datetime.datetime.strftime(datetime.datetime.now() -
                                                                    relativedelta(months=1, days=1), '%Y%m%d')]

        # 读取plat订单
        orders = pd.merge(orders, new_product, on='产品代码', how='left')
        orders['开发新品'] = orders['开发新品'].fillna('非开发新品')

        if orders.empty == True:
            break
        orders['订单总金额(包含客户运费、平台补贴)'] = orders.apply(lambda x: 0 if (x['订单类型'] == 'refund') else x['订单总金额(包含客户运费、平台补贴)'],
                                                    axis=1)
        if plat != 'ebay':
            account = read_data().read_account_sales(plat)
            result = order_to_sale(plat, orders, account)
        else:
            # 先定义一个容器
            result = pd.DataFrame()
            for site, value in ebay_site.items():
                # 读取订单信息
                data_site = orders[orders['站点'] == site]
                # 读取ebay产品分配信息
                product_sale = read_data().read_product_sale(value)
                result_site = order_to_sale(plat, data_site, product_sale)
                result = pd.concat([result, result_site], axis=0, sort=False)

        print(result)
        all_plat_result_last = pd.concat([all_plat_result_last, result], axis=0, sort=False)
    all_plat_result_last['销售员'] = all_plat_result_last['销售员'].fillna('未指定')
    all_plat_result_last['销售员'] = all_plat_result_last['销售员'].str.lower()

    '''站点数据计算'''

    Site_caculation(all_plat_result_last, 'a', '上期')
    '''核算开发业绩'''
    result_developer_last = Caculation().caculation_developer(all_plat_result_last)
    result_developer_last = result_developer_last.fillna(0)

    sale_result.to_csv('销售业绩上月.csv', index=None, encoding='utf_8_sig')
    result_developer_last.to_csv('开发业绩上月.csv', encoding='utf_8_sig')

# 当月处理好的订单写入数据库
def order_instant(all_plat_result_instant):
    # 订单存入数据库删除
    '''清空数据库'''
    DBKWARGS = Settings().DBKWARGS
    conn = pymysql.connect(**DBKWARGS)
    sql = """truncate table custom_statement"""
    cur = conn.cursor()
    cur.execute(sql)
    conn.commit()
    conn.close()
    # 重新引用一个变量，防止之后加其他的东西

    # 重新指定列名，和数据库的列名一一对应
    all_plat_result_instant.columns = ['orderType', 'platform', 'site', 'account', 'warehouse', 'orderCode', 'sku', 'seller',
                           'developer',
                           'saleroom', 'itemID', 'profits', 'number', 'purchasePrice', 'paytime', 'country',
                           'sortingdepot',
                           'newon', 'value', 'sales', 'productitle']
    # 要有这一句，随便设置哪列都可以
    all_plat_result_instant.set_index(['orderType'], inplace=True)
    # 创建引擎
    conn = create_engine('***')
    # 插入数据库
    pd.io.sql.to_sql(all_plat_result_instant, 'custom_statement', conn, schema='skykey', if_exists='append')
    conn.dispose()

def use_vba(file_path, VBA):
    '''
    :param file_path: 要运行excel宏的工作簿的地址
    :param VBA: 宏的名字
    :return: None
    '''
    '''定义excel参数'''
    print('-----------------------------------------------------------------------')
    print('正在运行excel宏')
    vba = win32com.client.DispatchEx('Excel.Application')
    vba.Visible = True
    vba.DisplayAlerts = 0
    vbabook = vba.Workbooks.Open(file_path, False)
    # 运行VBA
    vbabook.Application.Run(VBA)
    # 关闭excel， True为保存excel改动
    vbabook.Close(True)
    vba.quit()
    print('-----------------------------------------------------------------------')
    print('运行宏完毕')

if __name__ == '__main__':
    # 程序的开始时间
    start = datetime.datetime.now()
    main()
    # 指定要使用VBA的excel的绝对路径，一定要是绝对的
    result_file_path = 'D:/众结资料/1日常工作内容/每日销售开发业绩(Python)/' + str(datetime.datetime.strftime(
        datetime.datetime.now(), '%Y%m%d')) + '/业绩.xlsm'
    # 要使用的宏名称
    VBA = '业绩.xlsm!模块1.计算'
    # 使用VBA
    use_vba(result_file_path, VBA)
    # 结束时间
    end = datetime.datetime.now()
    print('----------------------------------------------------')
    print('程序运行时间为:')
    print(end - start)
    print('----------------------------------------------------')