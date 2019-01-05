class Settings():

    def __init__(self):
        # 配置平台信息
        self.plat_list = ('***')
        # 配置ebay站点分配信息
        self.ebay_site = {'***'}
        # 数据库端口，用户密码
        self.DBKWARGS = {'host': '***', 'user': '***', 'passwd': '***', 'db': '***'}
        # 存放上月订单数据库名
        self.last_order = 'last_order'
        # 存放本月订单数据库名
        self.instant_order = 'instant_order'


