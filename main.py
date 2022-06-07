###
# 作者：程梦飞
# 联系电话：就不联系了
# 声明：此代码仅作为分享交流所用，不许未经作者同意作为商业盈利
###

from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep
import pandas as pd


def Select(id,xpath):
    driver.find_element_by_id(id).click()
    driver.find_element_by_id(id).find_element_by_xpath(xpath).click()

options = EdgeOptions()
options.use_chromium = True
options.binary_location = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" # 浏览器的位置
driver = Edge(options=options, executable_path=r"C:\Users\Cmf\Downloads\edgedriver_win64\msedgedriver.exe") # 相应的浏览器的驱动位置
driver.get("http://zjt.icitpower.com/pc/Console/Login.aspx")#打开地址网页
driver.maximize_window()#最大化网页
## login
driver.find_element_by_id("txtUserName").send_keys("xxxxx")#账号
driver.find_element_by_id("txtUserPwd").send_keys("xxxx")#密码
driver.find_element_by_id("btnLogin").click()


df = pd.read_excel("data.xlsx",index_col=0)

for num in range(3,300):
    row_data = df.loc[[num]].values
    for cel in row_data:
        print(cel)

    address = str(cel[4])
    name = cel[10]
    p = cel[14]
    area = cel[16]
    year = cel[17]
    print(num,address,name,p,area,year)

    driver.switch_to.default_content()
    driver.find_element_by_xpath('//*[@id="menu"]/li[1]/dl/dd[1]/a').click()
    driver.switch_to.frame(driver.find_element_by_id('iframe1'))
    #建筑位置
    driver.find_element_by_id("ddlDist4").click()
    driver.find_element_by_id("ddlDist4").find_element_by_xpath('//*[@id="ddlDist4"]/option[18]').click()

    #详细地址
    driver.find_element_by_xpath('//*[@id="tbRoomAddress"]').click()
    driver.find_element_by_xpath('//*[@id="tbRoomAddress"]').send_keys(address)

    #土地性质
    driver.find_element_by_id("sel_type").click()
    driver.find_element_by_id("sel_type").find_element_by_xpath('//*[@id="sel_type"]/option[3]').click()

    #所在区域
    driver.find_element_by_xpath('//*[@id="cbArea_3"]').click()

    #建筑名称
    temp = "房屋"+name
    driver.find_element_by_xpath('//*[@id="tbName"]').send_keys(temp)

    #排查对象
    driver.find_element_by_id("sel_pc_object").click()
    driver.find_element_by_id("sel_pc_object").find_element_by_xpath('//*[@id="sel_pc_object"]/option[2]').click()

    #设计方式
    driver.find_element_by_id("sel_design").find_element_by_xpath('//*[@id="sel_design"]/option[3]').click()

    #结构类型
    driver.find_element_by_id("sel_stuct").find_element_by_xpath('//*[@id="sel_stuct"]/option[2]').click()

    #建造方式
    Select("sel_ways",'//*[@id="sel_ways"]/option[3]')

    #自建房屋建筑
    Select("sel_pc_object_zj",'//*[@id="sel_pc_object_zj"]/option[3]')

    #建筑层数（地上）
    if p == 1:
        Select("sel_floor", '//*[@id="sel_floor"]/option[1]')
    elif p == 2:
        Select("sel_floor",'//*[@id="sel_floor"]/option[2]')

    #所有权人
    driver.find_element_by_id("btnAdd1").click()
    driver.switch_to.frame(driver.find_element_by_id("layui-layer-iframe1"))
    WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id("tbName")).send_keys(name)
    #driver.find_element_by_id("tbName").send_keys(name)
    driver.find_element_by_id("btnSave").click()

    driver.switch_to.default_content()
    driver.switch_to.frame(driver.find_element_by_id('iframe1'))
    driver.find_element_by_class_name("layui-layer-ico.layui-layer-close.layui-layer-close1").click()

    #使用人
    driver.find_element_by_id("btnAdd2").click()
    driver.switch_to.frame(driver.find_element_by_id("layui-layer-iframe2"))

    WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id("tbName")).send_keys(name)
    #driver.find_element_by_id("tbName").send_keys(name)
    driver.find_element_by_id("btnSave").click()

    driver.switch_to.default_content()
    driver.switch_to.frame(driver.find_element_by_id('iframe1'))
    driver.find_element_by_class_name("layui-layer-ico.layui-layer-close.layui-layer-close1").click()
    #建筑面积
    driver.find_element_by_id("tbArea").send_keys(area)
    #建成时间
    driver.find_element_by_id("tbYear").send_keys(year)
    #建造年代
    if year <= 1980:
        Select("sel_year", '//*[@id="sel_year"]/option[2]')
    elif year >= 1981 and year <= 1990:
        Select("sel_year", '//*[@id="sel_year"]/option[3]')
    elif year >= 1991 and year <= 2000:
        Select("sel_year", '//*[@id="sel_year"]/option[4]')
    elif year >= 2001 and year <= 2010:
        Select("sel_year", '//*[@id="sel_year"]/option[5]')
    elif year >= 2011 and year <= 2015:
        Select("sel_year", '//*[@id="sel_year"]/option[6]')
    elif year >=2016:
        Select("sel_year", '//*[@id="sel_year"]/option[7]')

    #已取得的行政许可
    #driver.find_element_by_xpath('//*[@id="cb_license_2"]').click()
    btn = driver.find_element_by_xpath('//*[@id="cb_license_2"]')
    driver.execute_script("arguments[0].click();", btn)
    #违法建设和违法违规审批
    btn = driver.find_element_by_xpath('//*[@id="radio_wf2"]')
    driver.execute_script("arguments[0].click();", btn)


    #实际用途
    btn = driver.find_element_by_xpath('//*[@id="cbUseType_8"]')
    driver.execute_script("arguments[0].click();", btn)

    #签字排查 - 政府工作人员
    btn = driver.find_element_by_xpath('//*[@id="btnAdd3"]')
    driver.execute_script("arguments[0].click();", btn)
    sleep(0.5)
    driver.switch_to.frame(driver.find_element_by_id("layui-layer-iframe3"))

    WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id("tbName")).send_keys("程展召")
    #driver.find_element_by_id("tbName").send_keys("程展召")
    driver.find_element_by_id("tbUnit").send_keys("曹集乡政府")
    driver.find_element_by_id("btnSave").click()

    driver.switch_to.default_content()
    driver.switch_to.frame(driver.find_element_by_id('iframe1'))
    driver.find_element_by_class_name("layui-layer-ico.layui-layer-close.layui-layer-close1").click()

    #签字排查 - 技术专家
    btn = driver.find_element_by_xpath('//*[@id="btnAdd4"]')
    driver.execute_script("arguments[0].click();", btn)
    sleep(0.5)
    driver.switch_to.frame(driver.find_element_by_id("layui-layer-iframe4"))
    WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id("tbName")).send_keys("梁东建")
    #driver.find_element_by_id("tbName").send_keys("梁东建")
    driver.find_element_by_id("tbUnit").send_keys("夏邑县住房保障服务中心")
    driver.find_element_by_id("btnSave").click()

    driver.switch_to.default_content()
    driver.switch_to.frame(driver.find_element_by_id('iframe1'))
    driver.find_element_by_class_name("layui-layer-ico.layui-layer-close.layui-layer-close1").click()

    #排查结论
    driver.find_element_by_id("sel_result").find_element_by_xpath('//*[@id="sel_result"]/option[2]').click()
    #排查时间
    driver.find_element_by_xpath('//*[@id="tb_pc_date"]').click()
    driver.switch_to.default_content()
    iframe = driver.find_element('xpath', '//*[@id="_my97DP"]/iframe')
    driver.switch_to.frame(iframe)
    driver.find_element_by_xpath('/html/body/div/div[3]/table/tbody/tr[6]/td[3]').click()
    driver.switch_to.default_content()
    driver.switch_to.frame(driver.find_element_by_id("iframe1"))
    #保存
    WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_xpath('//*[@id="btnSave"]')).click()