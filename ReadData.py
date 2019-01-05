import pandas as pd
import pymysql
from settings import Settings
import datetime
from dateutil.relativedelta import relativedelta

class read_data(object):

    def __init__(self):
        pass

    def read_order(self, sql_path, plat):
        DBKWARGS = Settings().DBKWARGS
        conn = pymysql.connect(**DBKWARGS)
        data = pd.read_sql("select * from " + sql_path + " where platform='" + plat + "';", con=conn)
        conn.close()
        data.columns = ['订单类型', '平台', '站点', '平台账号', '发运仓库', '订单号', '产品代码',
                        '销售负责人', '产品开发人', '订单总金额(包含客户运费、平台补贴)', 'ITEM_ID(ASIN)',
                        '毛利', '数量', '平均采购价', '付款时间', '国家']
        # 去掉两个sku
        data = data[(data['产品代码'] != 'S3148030710') & (data['产品代码'] != 'TEST1')]
        # 数据在数据库中都是字符格式，需转成对应的格式
        data['付款时间'] = pd.to_datetime(data['付款时间'])
        data['ITEM_ID(ASIN)'] = pd.to_numeric(data['ITEM_ID(ASIN)'], errors='ignore')
        data['订单总金额(包含客户运费、平台补贴)'] = pd.to_numeric(data['订单总金额(包含客户运费、平台补贴)'], errors='ignore')
        data['毛利'] = pd.to_numeric(data['毛利'], errors='ignore')
        data['数量'] = pd.to_numeric(data['数量'], errors='ignore')
        data['平均采购价'] = pd.to_numeric(data['平均采购价'], errors='ignore')
        return data

    '''读取账号分配信息'''
    def read_account_sales(self, plat):
        DBKWARGS = Settings().DBKWARGS
        conn = pymysql.connect(**DBKWARGS)
        data = pd.read_sql(
            "SELECT * FROM account_assign WHERE platform = '" + plat + "' AND performance != '未指定' AND performance != '未分配' AND createtime = '" + str(
            datetime.datetime.strftime(datetime.datetime.now() - relativedelta(days=2), "%Y%m")) + "';",
            con=conn)
        conn.close()
        data = data[['account', 'performance']]
        data.columns = ['平台账号', '销售员']
        return data

    '''读取ebay产品分配信息'''
    def read_product_sale(self, site):
        DBKWARGS = Settings().DBKWARGS
        conn = pymysql.connect(**DBKWARGS)
        data = pd.read_sql(
            "SELECT * FROM product WHERE " + site + " != '未指定' AND " + site + " != '不做' AND " + site + " != '侵权' AND " + site + " != '危险品';",
            con=conn)
        conn.close()
        data = data[['productSku', site]]
        data.columns = ['产品代码', '销售员']
        return data

    '''读取产品名称信息'''
    def read_productTitle(self, ):
        DBKWARGS = Settings().DBKWARGS
        conn = pymysql.connect(**DBKWARGS)
        sql = """select productSku,productTitle from product;"""
        product_title = pd.read_sql(sql, con=conn)
        product_title.columns = ['产品代码', '产品名称']
        return product_title

    '''读取月初定下的新品数据'''
    def read_newproduct(self):
        fh = open('配置信息/新品.xlsx', 'rb')
        newproduct = pd.read_excel(fh)
        newproduct = newproduct['产品代码'].tolist()
        return newproduct

    '''读取当月转清仓数据'''
    def read_clearance_goods(self):
        fh = open('配置信息/当月转清仓.xlsx', 'rb')
        sheets = pd.read_excel(fh, sheet_name='读取批次', usecols=[0, 1], index_col=[0])
        sheets = sheets[sheets['时间'] != '无']
        if len(sheets['时间'].tolist()) <= 0:
            return None
        else:
            clearance_goods = []
            for index in sheets.index:
                sheet_each = pd.read_excel(fh, sheet_name=index)
                sheet_each = sheet_each['产品代码'].tolist()
                sheet_each.insert(0, sheets.loc[index].values)
                clearance_goods.append(sheet_each)
        return clearance_goods

    '''读取当月开发新品'''
    def read_new_product(self, sheets):
        fh = open('配置信息/开发新品表.xlsx', 'rb')
        data = pd.read_excel(fh, sheet_name=sheets)
        if sheets == '本月':
            data = data['产品代码'].tolist()
        return data

    '''读取特定库龄数据'''
    def read_special_product(self, plat):
        fh = open('配置信息/特定库龄.xlsx', 'rb')
        data = pd.read_excel(fh, sheet_name=plat)
        data = data['产品代码'].tolist()
        return data

# data = _read_data().read_productTitle()
# print(data)