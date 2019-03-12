import warnings
import requests
import xlrd
from lxml import etree
from pymysql import connect


class ComparyInfo:
    def __init__(self):
        # cookies 模拟登陆
        self.headers = {
            "Host": "www.qichacha.com",
            "Connection": "keep-alive",
            "Accept":"*/*",
            "X-Requested-With": "XMLHttpRequest",
            "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36",
            "Referer": "https://www.qichacha.com/",
            "Accept-Encoding": "gzip, deflate, br",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Cookie": "QCCSESSID=uohcqrm7mhgh2bc7l2kfbcvj94; UM_distinctid=169575b1f0692f-0f221724388d04-333b5602-100200-169575b1f0734c; CNZZDATA1254842228=607646511-1551947251-https%253A%252F%252Fwww.baidu.com%252F%7C1551947251; zg_did=%7B%22did%22%3A%20%22169575b212a217-05ccd87bf22a5a-333b5602-100200-169575b212b72b%22%7D; Hm_lvt_3456bee468c83cc63fb5147f119f1075=1551948784; hasShow=1; _uab_collina=155194878416882669474286; acw_tc=701919a515519486418937540ed5b3c34e9fab732f5c2e70ac1a5e254c; Hm_lpvt_3456bee468c83cc63fb5147f119f1075=1551948873; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201551948783919%2C%22updated%22%3A%201551948920892%2C%22info%22%3A%201551948783923%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%5C%22%24utm_source%5C%22%3A%20%5C%22baidu%5C%22%2C%5C%22%24utm_medium%5C%22%3A%20%5C%22cpc%5C%22%2C%5C%22%24utm_term%5C%22%3A%20%5C%22%E6%9F%A5%E5%85%AC%E5%8F%B8%E5%A4%9A%E8%AF%8D1%5C%22%7D%22%2C%22referrerDomain%22%3A%20%22www.baidu.com%22%2C%22cuid%22%3A%20%224fe56043271280809a8c3346eca5da85%22%7D"}
        self.baseurl = "https://www.qichacha.com/search?key="
        self.pageurl = "https://www.qichacha.com"
        self.proxies = {"http": "http://183.247.152.98:53281"} # 普通代理
        # 连接数据库，创建游标
        self.db = connect("localhost", "root", "123456", charset="utf8")
        self.cursor = self.db.cursor()

    def getPageHtm(self,name):
        # 获取企业页面链接
        baseurl = self.baseurl + str(name)
        res = requests.get(baseurl, proxies = self.proxies, headers = self.headers)
        res.encoding = "utf-8"
        html = res.text
        compHtml = etree.HTML(html)
        compUrl = compHtml.xpath('//*[@id="search-result"]/tr[1]/td[3]/a/@href')
        return compUrl

    def comParyInformation(self,compUrl):
        # 提取公司页面信息
        url = self.pageurl + compUrl[0]
        res = requests.get(url, proxies = self.proxies, headers=self.headers)
        res.encoding = "utf-8"
        html = res.text
        parseHtml = etree.HTML(html)
        # 公司名称
        info1 = parseHtml.xpath('//*[@id="company-top"]/div[2]/div[2]/div[1]/h1/text()')

        # 法人
        info2 = parseHtml.xpath('//*[@id="Cominfo"]/table[1]/tr[2]/td[1]/div/div[1]/div[2]/a/h2/text()')

        # 注册资本
        info3 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[1]/td[2]/text()')
        # 成立时间
        info4 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[2]/td[4]/text()')

        # 邮箱
        info5 = parseHtml.xpath('//*[@id="company-top"]/div[2]/div[2]/div[3]/div[2]/span[1]/span[2]/a/text()')
        # 电话
        info6 = parseHtml.xpath('//*[@id="company-top"]/div[2]/div[2]/div[3]/div[1]/span[1]/span[2]/span/text()')
        # 地址
        info7 = parseHtml.xpath('//*[@id="company-top"]/div[2]/div[2]/div[3]/div[2]/span[3]/a[1]/text()')

        # 官网
        info8 = parseHtml.xpath('//*[@id="company-top"]/div[2]/div[2]/div[3]/div[1]/span[3]/a/text()')
        if info8 == []:
            info8 = ["-"]
        # 经验状态
        info9 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[2]/td[2]/text()')
        # 统一社会信用代码
        info10 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[3]/td[2]/text()')
        # 纳税人识别号
        info11 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[3]/td[4]/text()')
        # 注册号
        info12 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[4]/td[2]/text()')
        # 组织机构代码
        info13 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[4]/td[4]/text()')
        # 公司类型
        info14 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[5]/td[2]/text()')
        # 所属行业
        info15 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[5]/td[4]/text()')
        # 所属地
        info16 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[7]/td[2]/text()')
        # 曾用名
        info17 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[8]/td[2]/text()')
        # 人员规模
        info18 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[9]/td[2]/text()')
        # 营业期限
        info19 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[9]/td[4]/text()')
        # 营业范围
        info20 = parseHtml.xpath('//*[@id="Cominfo"]/table[2]/tr[11]/td[2]/text()')
        info_list = info1 + info2 + info3 + info4 +info5 +info6 + info7 + info8 + info9 +info10 + info11+ info12 + info13 +info14 + info15 +info16 + info17 + info18 + info19 + info20
        # print(info_list)
        return info_list

    #存入本地数据库
    def writeToMysql(self,info_list):
        c_db = "create database if not exists ComparyInfom;"
        u_db = "use ComparyInfom;"
        c_tab ="create table if not exists info(" \
               "id int primary key auto_increment," \
               "compary_name varchar(100)," \
               "legal_person varchar(30)," \
               "money varchar(20)," \
               "ounding_time varchar(20)," \
               "email varchar(50)," \
               "phone_number varchar(30), " \
               "address varchar(30)," \
               "websit varchar(50)," \
               "status varchar(30)," \
               "credit_code varchar(30)," \
               "ratepayer_num varchar(30)," \
               "regest_num varchar(20)," \
               "Organization_code varchar(20)," \
               "compary_type varchar(20)," \
               "tmt varchar(30)," \
               "location varchar(10)," \
               "uesed_name varchar(30)," \
               "satff varchar(10)," \
               "operating_period varchar(30)," \
               "introduce varchar(512))charset=utf8;"
        warnings.filterwarnings("error")
        try:
            self.cursor.execute(c_db)
        except Warning:
            pass
        self.cursor.execute(u_db)
        try:
            self.cursor.execute(c_tab)
        except Warning:
            pass

        info_insert = "insert into info(compary_name,legal_person,money,ounding_time,email,phone_number,address,websit,status,credit_code,ratepayer_num,regest_num,Organization_code,compary_type,tmt,location,uesed_name,satff,operating_period,introduce) values('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')" % \
                   (info_list[0].strip(), info_list[1].strip(),info_list[2].strip(),info_list[3].strip(),info_list[4].strip(),info_list[5].strip(),info_list[6].strip(),info_list[7].strip(),info_list[8].strip(),info_list[9].strip(),info_list[10].strip(),info_list[11].strip(),info_list[12].strip(),info_list[13].strip(),info_list[14].strip(),info_list[15].strip(),info_list[16].strip(),info_list[17].strip(),info_list[18].strip(),info_list[19].strip())
        self.cursor.execute(info_insert)
        self.db.commit()

    #主函数
    def workOn(self):
        data = xlrd.open_workbook('companyname.xlsx')  # 打开一个excel
        sheet = data.sheet_by_index(0)  # 根据顺序获取sheet
        for i in range(sheet.ncols):
            compary_list = sheet.col_values(i)
            for name in compary_list:
                compUrl = self.getPageHtm(name.strip())
                info_list = self.comParyInformation(compUrl)
                self.writeToMysql(info_list)
        print("爬取成功")


if __name__ == "__main__":
    spader = ComparyInfo()
    spader.workOn()
