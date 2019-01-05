import pandas as pd
import numpy as np
class Site_caculation(object):
    def __init__(self, result, type, date):
        self.result = result
        self.type = type
        self.date = date
    '''站点销售数据拆分'''
    def __site_caculation(self, result):
        '''三大平台各站点毛利计算'''
        data_eas = result[(result['平台'] == 'ebay') | (result['平台'] == 'amazon') | (result['平台'] == 'shopee')]
        data_eas_re = data_eas[data_eas['订单类型'] == 'refund']
        res_eas = self.__site_pivot_table(data_eas, data_eas_re)
        # UK&非UK站点计算
        data_uk = result[(result['站点'] == 'Amazon.co.uk') | (result['站点'] == 'UK')]
        data_uk.loc[:, '站点'] = '`UK'
        data_uk_re = data_uk[data_uk['订单类型'] == 'refund']
        res_uk = self.__site_pivot_table(data_uk, data_uk_re)

        data_unuk = result[(result['站点'] != 'Amazon.co.uk') & (result['站点'] != 'UK')]
        data_unuk.loc[:, '站点'] = '`非UK'
        data_unuk_re = data_unuk[data_unuk['订单类型'] == 'refund']
        res_unuk = self.__site_pivot_table(data_unuk, data_unuk_re)

        resite = pd.concat([res_eas, res_uk, res_unuk], axis=0, sort=False)
        return resite
    '''站点销售情况计算方法'''
    def __site_pivot_table(self, data, data_re):
        res_eas = pd.pivot_table(data, index=['站点'], values=['订单总金额(包含客户运费、平台补贴)', '毛利'],
                                 aggfunc=np.sum, fill_value=0).reset_index()
        res_eas.columns = ['站点', '毛利', '销售额']
        res_eas_re = pd.pivot_table(data_re, index=['站点'], values=['毛利'],
                                    aggfunc=np.sum, fill_value=0).reset_index()
        res_eas_re.columns = ['站点', '退款金额']
        res_eas = pd.merge(res_eas, res_eas_re, on='站点', how='left')
        res_eas = res_eas.fillna(0)
        return res_eas
    '''仓库销售数据拆分'''
    def __warehouse_pivot_table(self, data, data_re):
        res_eas = pd.pivot_table(data, index=['发运仓库'], values=['订单总金额(包含客户运费、平台补贴)', '毛利'],
                                 aggfunc=np.sum, fill_value=0).reset_index()
        res_eas.columns = ['站点', '毛利', '销售额']
        res_eas_re = pd.pivot_table(data_re, index=['发运仓库'], values=['毛利'],
                                    aggfunc=np.sum, fill_value=0).reset_index()
        res_eas_re.columns = ['站点', '退款金额']
        res_eas = pd.merge(res_eas, res_eas_re, on='站点', how='left')
        res_eas = res_eas.fillna(0)
        return res_eas
    '''仓库销售情况计算方法'''
    def __warehouse_cacolation(self, result):
        '''计算amazonFBA&非FBA的'''
        result_fba = result[result['平台'] == 'amazon']
        data_am_fba = result_fba[result_fba['发运仓库'].str.contains(r'FBA')]
        data_am_fba.loc[:, '发运仓库'] = 'amazonFBA'
        data_am_fba_re = data_am_fba[data_am_fba['订单类型'] == 'refund']
        res_am_fba = self.__warehouse_pivot_table(data_am_fba, data_am_fba_re)

        data_am_unfba = result_fba[~result_fba['发运仓库'].str.contains(r'FBA')]
        data_am_unfba.loc[:, '发运仓库'] = 'amazon非FBA'
        data_am_unfba_re = data_am_unfba[data_am_unfba['订单类型'] == 'refund']
        res_am_unfba = self.__warehouse_pivot_table(data_am_unfba, data_am_unfba_re)

        '''仓库细分计算'''
        # 中&温
        data_cn = result[(result['发运仓库'] == 'SHZZ [上海中转仓]') | (result['发运仓库'] == 'WZC [温州仓]') | (result['发运仓库'] == 'SH [上海奉贤仓]')]
        data_cn.loc[:, '发运仓库'] = '中&温'
        data_cn_re = data_cn[data_cn['订单类型'] == 'refund']
        res_cn = self.__warehouse_pivot_table(data_cn, data_cn_re)
        # FBA仓
        data_fba = result[result['发运仓库'].str.contains(r'FBA')]
        data_fba.loc[:, '发运仓库'] = 'FBA仓'
        data_fba_re = data_fba[data_fba['订单类型'] == 'refund']
        res_fba = self.__warehouse_pivot_table(data_fba, data_fba_re)
        # UK仓
        data_uk = result[
            (result['发运仓库'] == 'UK [英国仓]') | (result['发运仓库'] == '4PX-UK [4px英国仓]') | (result['发运仓库'] == '4PX-UK-LONDON [4px英国伦敦仓]')]
        data_uk.loc[:, '发运仓库'] = 'UK仓'
        data_uk_re = data_uk[data_uk['订单类型'] == 'refund']
        res_uk = self.__warehouse_pivot_table(data_uk, data_uk_re)
        # US仓
        data_us = result[
            (result['发运仓库'] == 'GSW [古斯美西仓]') | (result['发运仓库'] == 'GSE [古斯美东仓]')]
        data_us.loc[:, '发运仓库'] = 'US仓'
        data_us_re = data_us[data_us['订单类型'] == 'refund']
        res_us = self.__warehouse_pivot_table(data_us, data_us_re)

        # SZ仓
        data_sz = result[
            (result['发运仓库'] == 'SZC [深圳仓]')]
        data_sz.loc[:, '发运仓库'] = '深圳仓'
        data_sz_re = data_sz[data_sz['订单类型'] == 'refund']
        res_sz = self.__warehouse_pivot_table(data_sz, data_sz_re)

        res = pd.concat([res_am_unfba, res_am_fba, res_cn, res_fba, res_uk, res_us, res_sz], axis=0, sort=False)
        return res

    def site_sale(self, result, type, date):

        resite = self.__site_caculation(result)
        res = self.warehouse_cacolation(result)
        res_site = pd.concat([resite, res], axis=0, sort=False)
        res_site['期数'] = date
        res_site.to_csv('站点销售数据.csv', encoding='utf_8_sig', index=None, mode=type)
        return res_site

class Caculation(object):
    def __init__(self):
        pass

    def caculation_sale_instant(self, result):
        '''按仓库分类&销售员计算结果'''
        wh_result = pd.pivot_table(result, index=['销售员', '仓库分类'], values=['订单总金额(包含客户运费、平台补贴)', '毛利'],
                                   aggfunc=np.sum, fill_value=0)
        '''按销售员计算结果'''
        sale_result = pd.pivot_table(result, index=['销售员'], values=['订单总金额(包含客户运费、平台补贴)', '毛利'],
                                     aggfunc=np.sum, fill_value=0).reset_index()
        sale_result.columns = ['销售员', '毛利', '订单总金额(包含客户运费、平台补贴)']
        '''特定库龄货值计算'''
        try:
            result_sp = result[result['仓库分类'] == '特定库龄']
            sale_result_sp = pd.pivot_table(result_sp, index=['销售员'], values=['货值'],
                                            aggfunc=np.sum, fill_value=0).reset_index()
            sale_result_sp.columns = ['销售员', '货值']
            sale_result = pd.merge(sale_result, sale_result_sp, on='销售员', how='left')
        except:
            sale_result['货值'] = 0
        '''计算负毛利订单和正常产品的负毛利订单'''
        result_minus_normal = result[(result['订单类型'] != 'refund') & (result['订单类型'] != 'resend')]
        result_minus_normal = result_minus_normal[result_minus_normal['毛利'] < 0]
        result_minus = result_minus_normal[result_minus_normal['仓库分类'] != '特定库龄']
        '''可能为空的情况'''
        try:
            sale_result_m = pd.pivot_table(result_minus_normal, index=['销售员'], values=['毛利'],
                                           aggfunc=np.sum, fill_value=0).reset_index()
            sale_result_m.columns = ['销售员', '负订单毛利']
            sale_result = pd.merge(sale_result, sale_result_m, on='销售员', how='left')
            try:
                sale_result_n = pd.pivot_table(result_minus, index=['销售员'], values=['毛利'],
                                               aggfunc=np.sum, fill_value=0).reset_index()
                sale_result_n.columns = ['销售员', '正常产品负订单毛利']
                sale_result = pd.merge(sale_result, sale_result_n, on='销售员', how='left')
            except:
                return wh_result, sale_result
            return wh_result, sale_result
        except:
            return wh_result, sale_result

    def caculation_sale_last(self, all_plat_result_last):
        sale_result = pd.pivot_table(all_plat_result_last, index=['销售员'], values=['订单总金额(包含客户运费、平台补贴)', '毛利'],
                                     aggfunc=np.sum, fill_value=0).reset_index()
        sale_result.columns = ['销售员', '毛利', '订单总金额(包含客户运费、平台补贴)']
        '''业绩保存'''
        sale_result = sale_result.fillna(0)
        return sale_result

    def caculation_developer(self, data):

        data = data[data['开发新品'] == '开发新品']
        result = pd.pivot_table(data, index=['产品开发人'], values=['订单总金额(包含客户运费、平台补贴)', '毛利'],
                                aggfunc=np.sum, fill_value=0)
        result = result.fillna(0)
        return result

    '''核算店铺销售额毛利'''
    def caculation_by_account(self, result):
        # 店铺的销售情况
        res = pd.pivot_table(result, index=['平台账号'], values=['订单总金额(包含客户运费、平台补贴)', '毛利'],
                             aggfunc=np.sum, fill_value=0).reset_index()
        res.columns = ['平台账号', '订单总金额(包含客户运费、平台补贴)', '毛利']
        # 店铺的FBA销售情况
        result_fba = result[result['发运仓库'].str.contains(r'FBA')]
        res_fba = pd.pivot_table(result_fba, index=['平台账号'], values=['订单总金额(包含客户运费、平台补贴)', '毛利'],
                                 aggfunc=np.sum, fill_value=0).reset_index()
        # 重新指定列名
        res_fba.columns = ['平台账号', '销售额(FBA)', '毛利(FBA)']

        # 结果合并
        ress = pd.merge(res, res_fba, on='平台账号', how='left')
        ress = ress.fillna(0)
        return ress