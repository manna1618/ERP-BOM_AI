
'''
1.将新增的物料信息，录入ERP物料清单
2.ERP基础设置： 开启保存即审核；关闭重复性检查功能
3.chromedriver的版本与chrome对应：https://registry.npmmirror.com/binary.html?path=chromedriver/
'''
# 导包
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from selenium.webdriver import ActionChains
from time import sleep,time
def gen_bro():
    # 创建对象
    opt = Options()

    # opt.add_argument(UA)
    # # 无头浏览器
    # opt.add_argument("--headless")
    # opt.add_argument('--disable-gpu')
    # 设置窗口大小，防止出现堆叠问题
    opt.add_argument("--window-size=1900,1030")
    # 实例化浏览器对象
    print('创建browser对象...')
    bro = Chrome(options=opt)
    # 全局设置监测elements，如果元素加载出来了 就继续. 如果没加载出来. 会最多等待10s；限find操作，增加点击操作未操作的话 需sleep
    bro.implicitly_wait(10)
    return bro

def logout_erp(bro):
    bro.quit()
    print('关闭浏览器')

def login_erp(bro):
    # 金蝶云星空登录界面
    url = 'http://218.68.31.6:8888/k3cloud'
    bro.get(url)
    # 选择数据中心
    select_tag = bro.find_element(By.XPATH,'/html/body/div[1]/form/div[2]/div/div[2]/div[2]/div[3]/div[1]/div[3]/span[1]')
    sleep(3)
    select_tag.click()
    for i in range(5):
        select_tag.send_keys(Keys.DOWN)

    select_tag.send_keys(Keys.ENTER)
    sleep(3)

    print('登录ERP账号...')
    # 解析登录
    user_input_tag = bro.find_element(By.ID, 'user').send_keys('李洋')
    sleep(1)
    password_input_tag = bro.find_element(By.ID, 'password').send_keys('111111')
    sleep(1)
    btnLogin_tag = bro.find_element(By.ID, 'btnLogin')
    btnLogin_tag.click()
    sleep(8)
    print('完成ERP登录！准备跳转至物料清单界面...')

    # 解析详情页
    # 点击主界面,未设置常用功能-物料新增
    # 点击四个方形
    try:
        kd_SystemMenu = bro.find_element(By.XPATH,'//*[@class = "k-link kd_SystemMenu"]')
    except:
        print('账号异常或已登录，强制退出中...')
        # 点击警告提示：账号已登录
        try:
            jinggaotishi_tag = bro.find_element(By.XPATH,'/html/body/div[15]/div[2]/div[2]/button[1]')
            sleep(1)
            jinggaotishi_tag.click()
            sleep(3)
            kd_SystemMenu = bro.find_element(By.XPATH, '//*[@class = "k-link kd_SystemMenu"]')
        except Exception as e:
            print('报错了：',e,'未跳转界面成功')

    sleep(1)
    kd_SystemMenu.click()
    print('点击基础资料')
    # 点击基础管理下的基础资料
    jichuguanli_tag = bro.find_element(By.XPATH,
                                              '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/div[2]/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[1]/div/ul[1]/li[8]')
    sleep(1)
    # print("基础管理->",jichuguanli_tag.text)

    jichuguanli_tag.click()

    # 点击基础管理下的基础资料
    jichuziliao_tag = bro.find_element(By.XPATH, '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/div[2]/div/div/div/div[2]/div/div[1]/div/div[1]/div/div[1]/div/ul[1]/li[8]/ul/li/div/div[2]')
    sleep(1)
    # print("基础资料->",jichuziliao_tag.text)

    jichuziliao_tag.click()
    # 物料标签
    insert_btn_tag = bro.find_element(By.XPATH, '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/div[2]/div/div/div/div[2]/div/div[1]/div/div[2]/div/div[1]/div/div[2]/div/div[3]/div[1]/div[1]/a[1]')

    insert_btn_tag.click()
    wuliao = insert_btn_tag.text
    print("物料->",insert_btn_tag.text)

    sleep(5)
    print('已打开物料清单！待录入物料信息')
    # 直接点击常用功能里的-物料新增
    # insert_btn_tag = bro.find_element(By.ID, '369937')
    # insert_btn_tag.click()
    # sleep(3)

    return wuliao

# 点击注销账号
def logout_acount(bro):
    # 选择注销按钮
    zhuxiao_tag = bro.find_element(By.XPATH,
                                  '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[1]/div'
                                  '/div[2]/div/div[1]/div[1]/div/div[1]/a')
    sleep(3)
    zhuxiao_tag.click()
    sleep(3)
    print('已正常退出账号！')

def write_jiben_info(bro,bianma,mingcheng,guige,wuliaofenzu):
    # 新增

    xinzeng_tag = bro.find_element(By.XPATH,
                                  '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                  'div[3]/div/div/div/div[2]/div/div[1]/ul/li[2]/span/span[1]')

    sleep(3)
    xinzeng_tag.click()
    sleep(5)
    # 录入编码
    bianma_tag = bro.find_element(By.XPATH,
                                  '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                  'div[3]/div/div/div/div[3]/div/div[1]/div[2]/div/div/div[1]/div[2]/div/input')

    sleep(3)
    bianma_tag.send_keys(bianma)
    # 录入名称
    mingcheng_tag = bro.find_element(By.XPATH,
                                     '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                     'div[3]/div/div/div/div[3]/div/div[1]/div[2]/div/div/div[2]/div[2]/div/div/span/span/input')
    sleep(3)
    mingcheng_tag.send_keys(mingcheng)
    # 录入规格型号
    guige_tag = bro.find_element(By.XPATH,
                                 '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                 'div[3]/div/div/div/div[3]/div/div[1]/div[3]/div[2]/div/div[5]/div[2]/div/div/span/span/input')
    sleep(3)
    guige_tag.send_keys(guige)
    # 录入物料分组
    wuliaofenzu_tag = bro.find_element(By.XPATH,
                                 '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                 'div[3]/div/div/div/div[3]/div/div[1]/div[3]/div[2]/div/div[7]/div[2]/div/span/span/input')
    sleep(3)
    wuliaofenzu_tag.send_keys(wuliaofenzu)

    #点击属性
    yunxushengchan_tag = bro.find_element(By.XPATH,
                                          '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                          'div[3]/div/div/div/div[3]/div/div[1]/div[3]/div[2]/div/div[15]/div[2]/div/div/input')
    sleep(3)
    # yunxushengchan_tag.click()
    bro.execute_script('arguments[0].click();',yunxushengchan_tag)
    yunxuweiwai_tag = bro.find_element(By.XPATH,
                                          '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                          'div[3]/div/div/div/div[3]/div/div[1]/div[3]/div[2]/div/div[13]/div[2]/div/div/input')
    sleep(3)
    # yunxuweiwai_tag.click()
    bro.execute_script('arguments[0].click();', yunxuweiwai_tag)

def write_kucun_info(bro):
    # 点击启用批号管理
    qiyongpihaoguanli_tag = bro.find_element(By.XPATH,
                                          '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                          'div[3]/div/div/div/div[3]/div/div[1]/div[4]/div[2]/div/div[19]/div[2]/div/div/label')

    sleep(3)
    qiyongpihaoguanli_tag.click()
    # 点击启用保质期管理
    qiyongbaozhiqiguanli_tag = bro.find_element(By.XPATH,
                                             '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                             'div[3]/div/div/div/div[3]/div/div[1]/div[4]/div[2]/div/div[24]/div[2]/div/div/label')

    sleep(3)
    qiyongbaozhiqiguanli_tag.click()
    baozhiqidanwei_tag = bro.find_element(By.XPATH,
                                             "/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/"
                                             "div[3]/div/div/div/div[3]/div/div[1]/div[4]/div[2]/div/div[23]/div[2]/div/span/span/span[1]")

    # 录入保质期
    guige_tag = bro.find_element(By.XPATH,
                                 '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                 'div[3]/div/div/div/div[3]/div/div[1]/div[4]/div[2]/div/div[2]/div[2]/div/span/span/input')
    sleep(3)
    guige_tag.send_keys('1825')

def write_zhiliang_info(bro):
    #点击属性
    lailiaojianyan_tag = bro.find_element(By.XPATH,
                                          '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                          'div[3]/div/div/div/div[3]/div/div[1]/div[7]/div[2]/div/div/div[1]/div[2]/div/div/input')

    sleep(3)
    # lailiaojianyan_tag.click()
    bro.execute_script('arguments[0].click();', lailiaojianyan_tag)

    kucunjianyan_tag = bro.find_element(By.XPATH,
                                          '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                          'div[3]/div/div/div/div[3]/div/div[1]/div[7]/div[2]/div/div/div[7]/div[2]/div/div/input')
    sleep(3)
    # kucunjianyan_tag.click()
    bro.execute_script('arguments[0].click();', kucunjianyan_tag)

def save(bro):
    save_tag = bro.find_element(By.XPATH,
                                  '/html/body/form/div[1]/div/div/div/div[1]/div/div[1]/div/div/div[1]/div/div[2]/div[2]/div/'
                                  'div[3]/div/div/div/div[2]/div/div[1]/ul/li[5]/span')
    sleep(3)
    save_tag.click()

def update_erp_material_list(bro,code,type,Part_Number,name):
    # 要录入的数据
    bianma = code
    mingcheng = name
    guige = Part_Number

    if bianma:
        wuliaofenzu = bianma.split('-')[0]
    else:
        wuliaofenzu = ''
        print('未定义编码或编码为空')

    print(bianma,'：开始写入基本信息')
    write_jiben_info(bro,bianma,mingcheng,guige,wuliaofenzu)

    print(bianma, '：开始写入库存信息')
    write_kucun_info(bro)

    print(bianma, '：开始写入质量信息')
    write_zhiliang_info(bro)

    print(bianma, '开始保存数据')
    save(bro)
    print(f'完成{bianma}信息的录入')



# 调试用代码
if __name__ == '__main__':
    bro = gen_bro()
    wuliao =login_erp(bro)
    print('================================================')
    code = 'C97-9994'
    type = 'LCR类'
    Part_Number = 'xxxxxx'
    name = "xxx"
    print(code,':开始填入物料信息==============>')
    update_erp_material_list(bro,code,type,Part_Number,name)
