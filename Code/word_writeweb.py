from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from .word_separate import WordFileClean
from . import config
import time


class WordWriteWeb:

    def __init__(self):
        self.url = config.url
        self.status = False
        self.chrome = webdriver.Chrome()
        self.chrome.implicitly_wait(30)
        self.chrome.get(self.url)
        self.chrome.maximize_window()
        time.sleep(20)

    def __str__(self):
        if self.status:
            msg = '合同录入结果:' + 'Sucess'
        else:
            msg = '合同录入结果:' + 'Failure'
        return msg

    '''
    def driver(self):
        self.chrome = webself.chromeChrome()
        self.chrome.implicitly_wait(30)
    '''

    def isElement(self, identifyBy, element):
        '''判断元素是否存在'''
        flag = False
        try:
            if identifyBy == 'id':
                self.chrome.find_element_by_id(element)
                flag = True
            elif identifyBy == 'xpath':
                self.chrome.find_element_by_xpath(element)
                flag = True
        except Exception as e:
            print('该元素不存在：',element)
            flag = False
        if flag:
            print('该元素存在为：',element)
        else:
            pass
        return flag

    def calendar(self,str_):
        '''
        日历插件中：日期对应的行列排列序号
        :param str_:日期中的日
        :return:行列排列顺序
        '''
        row = 1
        col = 0
        number = int(str_)
        if number < 7:
            col = number
        elif number == 7 or number == 14 or number == 21 or number == 8:
            row = number // 7
            col = 7
        else:
            row += number // 7
            col = number % 7
        return row,col

    def updateTitle(self, n, key):
        '''
        修改章节名字
        :param n: 章节序号
        :param key: 修改的内容
        :return: 无返回
        '''
        self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_%d_a"]' % (n)).click()
        e = self.chrome.find_element_by_xpath('//*[@id="insChapterTree"]//div[text()="章节"]')
        #time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="修改章节名称"]').click()
        self.chrome.find_element_by_xpath(
            '//*[@id="updateChapterForm"]/div[2]/fieldset/div/div/input').clear()
        self.chrome.find_element_by_xpath(
            '//*[@id="updateChapterForm"]/div[2]/fieldset/div/div/input').send_keys(key)
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

    def section_alter(self, data):
        '''
        web和Word的章节数一样，修改名字
        :return:
        '''
        self.chrome.switch_to.parent_frame()
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_2')
        sectionNumlocal = len(data)
        sectionNumweb = len(self.chrome.find_elements_by_xpath('//*[@id="insChapterTree_tree"]/li'))
        print('local端章节数为：%d' % sectionNumlocal)
        print('web端章节数为：%d' % sectionNumweb)
        if sectionNumlocal == sectionNumweb:
            # 修改每个section名字
            print('*' * 8, '修改每个章节的名字：开始', '*' * 8)
            n = 1
            for key in data:
                self.updateTitle(n,key)
                n += 1
            print('*' * 8, '修改每个章节的名字：结束', '*' * 8)
        else:
            print('Error：web端、Word文件两者的段落数不一样，请检查...')
            sys.exit()

    def section_judge(self, data):
        '''进行章节增删改功能'''
        self.chrome.switch_to.parent_frame()
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_2')
        sectionNumlocal = len(data)
        sectionNumweb = len(self.chrome.find_elements_by_xpath('//*[@id="insChapterTree_tree"]/li'))
        print('local端章节数为：%d'%sectionNumlocal)
        print('web端章节数为：%d' % sectionNumweb)
        if sectionNumlocal == sectionNumweb:
            # 修改每个section名字
            print('*' * 8, '修改每个章节的名字：开始', '*' * 8)
            n = 1
            for key in data:
                self.updateTitle(n,key)
                n += 1
            print('*' * 8, '修改每个章节的名字：结束', '*' * 8)
        elif sectionNumweb > sectionNumlocal:
            #从章节最后往后面删除
            num = sectionNumweb - sectionNumlocal
            print('*'*4,'web章节多了%d个,进行删除' % (num),'*'*4)
            for i in reversed(range(sectionNumlocal+1,sectionNumweb+1)):
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_%d_a"]'%i).click()
                e = self.chrome.find_element_by_xpath('//*[@id="insChapterTree"]//div[text()="章节"]')
                time.sleep(0.5)
                e.click()
                self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="删除章节"]').click()
                self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()
                self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

            print('*' * 8, '修改每个章节的名字：开始', '*' * 8)
            n = 1
            for key in data:
                self.updateTitle(n, key)
                n += 1
            print('*' * 8, '修改每个章节的名字：结束', '*' * 8)

        else:
            #从最后开始增加章节
            num = sectionNumlocal - sectionNumweb
            print('*' * 4, 'web章节少了%d个,进行添加' % (num), '*' * 4)
            for i in range(num):
                e = self.chrome.find_element_by_xpath('//*[@id="insChapterTree"]//div[text()="章节"]')
                time.sleep(0.5)
                e.click()
                self.chrome.find_element_by_id('//div[@class="kui-split-menu-text" and text()="增加章节"]').click()
                self.chrome.find_element_by_xpath('//*[@id="addChapterForm"]/div[2]/fieldset/div/div[1]/input').send_keys(i)
                self.chrome.find_element_by_xpath(
                    '//*[@id="addChapterForm"]/div[2]/fieldset/div/div[2]/div/ul/li[1]/input').click()
                self.chrome.find_element_by_id('popupTreeItem_btn').click()
                self.chrome.find_element_by_id('popupTree_tree_%d_span'%(sectionNumweb)).click()
                self.chrome.find_element_by_xpath('//*[@id="addChapterFormDialog_btn"]/button[2]/div[2]').click()
                frame = self.chrome.find_element_by_xpath(
                    '//*[@id="addChapterForm"]/div[2]/fieldset/div/div[4]/div/div[2]/iframe')
                self.chrome.switch_to.frame(frame)
                self.chrome.find_element_by_xpath('/html/body').send_keys(i)
                self.chrome.switch_to.parent_frame()
                self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()
                self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()
            print('*' * 8, '修改每个章节的名字：开始', '*' * 8)
            #修改章节名字
            n = 1
            for key in data:
                self.updateTitle(n, key)
                n += 1
            print('*' * 8, '修改每个章节的名字：结束', '*' * 8)


    def paragraph_judge(self, localNumber):
        '''对章节里的段落进行增删功能'''
        webNumber = len(self.chrome.find_elements_by_xpath('//*[@id="contentBlock"]/div[2]/div'))
        print('web的段落数：',webNumber)
        print('Word文件中的段落数量：', localNumber)
        if webNumber == localNumber:
            return True
        elif webNumber < localNumber:
            num = localNumber - webNumber
            print('添加不足的%d个段落'%num)
            for i in range(num):
                try:
                    e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
                    time.sleep(0.5)
                    e.click()
                    self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="添加段落"]').click()
                except Exception :
                    pass
            #可能会出现添加失败，再次添加
            webNumber = len(self.chrome.find_elements_by_xpath('//*[@id="contentBlock"]/div[2]/div'))
            num = localNumber - webNumber
            if webNumber == localNumber:
                return True
            else:
                for i in range(num):
                    try:
                        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
                        time.sleep(0.5)
                        e.click()
                        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="添加段落"]').click()
                    except Exception:
                        pass
            return True
        else:
            num = webNumber - localNumber
            print('删除多余的%d个段落' % num)
            for i in range(num):
                e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
                time.sleep(0.5)
                e.click()
                self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="删除段落"]').click()
                self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()
            return True

    def alter_information(self, information):
        '''
        修改基本信息---产品存续
        :param information: 要修改信息的list集合
        :return:
        '''
        self.chrome.find_element_by_xpath('//*[@id="kui-3"]/div[2]').click()
        frame = self.chrome.find_element_by_xpath('//*[@id="updateDialog_content"]/div/div[2]/iframe')
        self.chrome.switch_to.frame(frame)

        # 修改律师事务所
        e = self.chrome.find_element_by_xpath('//*[@id="kui-12_btn"]')
        time.sleep(0.5)
        e.click()

        print(information['律师事务所'])
        self.chrome.find_element_by_xpath('//*[@class="kui-select-list-grp"]/li[text()="%s"]' % (information['律师事务所'])).click()

        # 修改会计事务所
        self.chrome.find_element_by_xpath('//*[@id="kui-13_btn"]').click()
        self.chrome.find_element_by_xpath('//*[@class="kui-select-list-grp"]/li[text()="%s"]' % (information['会计师事务所'])).click()

        # 点击保存按钮
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="保存"]').click()

        # 点击确定按钮
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 点击取消按钮
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="取消"]').click()

        self.chrome.switch_to.parent_frame()

    def alter_information2(self, information):
        '''
        修改基本信息---产品开发、产品募集
        :param information: 要修改信息的list集合
        :return:
        '''
        self.chrome.find_element_by_xpath('//*[@id="kui-3"]/div[2]').click()
        frame = self.chrome.find_element_by_xpath('//*[@id="updateDialog_content"]/div/div[2]/iframe')
        self.chrome.switch_to.frame(frame)

        # 修改律师事务所
        e = self.chrome.find_element_by_xpath('//*[@id="kui-15_btn"]')
        time.sleep(0.5)
        e.click()

        self.chrome.find_element_by_xpath('//*[@class="kui-select-list-grp"]/li[text()="%s"]' % (information['律师事务所'])).click()

        # 修改会计事务所
        self.chrome.find_element_by_xpath('//*[@id="kui-16_btn"]').click()
        self.chrome.find_element_by_xpath('//*[@class="kui-select-list-grp"]/li[text()="%s"]' % (information['会计师事务所'])).click()

        # 点击保存按钮
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="保存"]').click()

        # 点击确定按钮
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 点击取消按钮
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="取消"]').click()

        self.chrome.switch_to.parent_frame()

    def has_cunxu(self, web_filename):
        '''
         产品存续-- “存续期限”为空，退出写入，有就继续
        :param web_filename:
        :return:
        '''
        self.chrome.get(self.url)
        time.sleep(2)

        # 产品管理
        self.chrome.find_element_by_xpath('//*[@class="kui-inline-block" and text()="产品管理"]').click()
        time.sleep(0.5)

        # 产品开发
        self.chrome.find_element_by_xpath('//*[@id="kui-11"]//span[text()="产品存续"]').click()
        time.sleep(0.5)

        # 产品名字：博时鑫禧灵活配置混合型证券投资基金
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_1')
        self.chrome.find_element_by_xpath('//*[@id="kui-13"]').send_keys(web_filename)
        time.sleep(0.5)
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="搜索"]').click()
        time.sleep(0.5)

        # 修改基本信息
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="修改"]').click()
        frame = self.chrome.find_element_by_xpath('//*[@id="updateDialog_content"]/div/div[2]/iframe')
        self.chrome.switch_to.frame(frame)

        try:
            # 点击保存按钮
            self.chrome.find_element_by_xpath('//*[@id="but_info"]/div[2]').click()

            # 点击确定按钮
            self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

            # 点击取消按钮
            self.chrome.find_element_by_xpath('//*[@id="info_Button"]/div[2]').click()

            self.chrome.switch_to.parent_frame()
            return True
        except Exception :
            return False

    def has_cunxu2(self, web_filename):
        '''
        产品开发----“存续期限”为空，退出写入，有就继续
        :param web_filename:
        :return:
        '''
        self.chrome.get(self.url)
        time.sleep(2)

        # 产品管理
        self.chrome.find_element_by_xpath('//*[@class="kui-inline-block" and text()="产品管理"]').click()
        time.sleep(0.5)

        # 产品开发
        self.chrome.find_element_by_xpath('//*[@id="kui-11"]//span[text()="产品开发"]').click()
        time.sleep(0.5)

        # 产品名字：博时鑫禧灵活配置混合型证券投资基金
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_1')
        self.chrome.find_element_by_xpath('//*[@id="kui-12"]').send_keys(web_filename)
        time.sleep(0.5)
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="搜索"]').click()
        time.sleep(0.5)

        # 修改基本信息
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="修改"]').click()
        frame = self.chrome.find_element_by_xpath('//*[@id="updateDialog_content"]/div/div[2]/iframe')
        self.chrome.switch_to.frame(frame)

        try:
            # 点击保存按钮
            self.chrome.find_element_by_xpath('//*[@id="but_info"]/div[2]').click()

            # 点击确定按钮
            self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

            # 点击取消按钮
            self.chrome.find_element_by_xpath('//*[@id="info_Button"]/div[2]').click()

            self.chrome.switch_to.parent_frame()
            return True
        except Exception :
            return False

    def has_cunxu3(self, web_filename):
        '''
        产品募集----“存续期限”为空，退出写入，有就继续
        :param web_filename:
        :return:
        '''
        self.chrome.get(self.url)
        time.sleep(2)

        # 产品管理
        self.chrome.find_element_by_xpath('//*[@class="kui-inline-block" and text()="产品管理"]').click()
        time.sleep(0.5)

        # 产品开发
        self.chrome.find_element_by_xpath('//*[@id="kui-11"]//span[text()="产品募集"]').click()
        time.sleep(0.5)

        # 产品名字：博时鑫禧灵活配置混合型证券投资基金
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_1')
        self.chrome.find_element_by_xpath('//*[@id="kui-13"]').send_keys(web_filename)
        time.sleep(0.5)
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="搜索"]').click()
        time.sleep(0.5)

        # 修改基本信息
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="修改"]').click()
        frame = self.chrome.find_element_by_xpath('//*[@id="updateDialog_content"]/div/div[2]/iframe')
        self.chrome.switch_to.frame(frame)

        try:
            # 点击保存按钮
            self.chrome.find_element_by_xpath('//*[@id="but_info"]/div[2]').click()

            # 点击确定按钮
            self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

            # 点击取消按钮
            self.chrome.find_element_by_xpath('//*[@id="info_Button"]/div[2]').click()

            self.chrome.switch_to.parent_frame()
            return True
        except Exception:
            return False

    def write1(self, web_filename, data):
        '''
        业务流程：产品开发-产品名-基金合同
        :param web_filename: 产品名字
        :param data: 录入的字典数据
        :return: 录入成功，文件下载成功
        '''

        # 链接URL
        print('打开：URL')
        self.chrome.get(self.url)
        time.sleep(2)

        # 产品管理
        print('点击：产品管理')
        # self.chrome.find_element_by_xpath('//*[@id="kui-11"]/div[2]/table/tbody/tr/td/ul/li[1]/span').click()
        self.chrome.find_element_by_xpath('//*[@class="kui-inline-block" and text()="产品管理"]').click()
        time.sleep(0.5)

        # 产品开发
        print('点击：产品开发')
        self.chrome.find_element_by_xpath('//*[@id="kui-11"]//span[text()="产品开发"]').click()
        time.sleep(0.5)

        # 产品名字
        print('点击：%s' % web_filename)
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_1')
        self.chrome.find_element_by_xpath('//*[@id="kui-12"]').send_keys(web_filename)
        time.sleep(0.5)
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="搜索"]').click()
        time.sleep(0.5)

        # 报会文件制作
        print('点击：报会文件制作')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="报会文件制作"]').click()


        # 基金合同：合同文件序号参数
        print('点击：基金合同')
        self.chrome.find_element_by_xpath('//*[@id="fileTable_listContainer"]//span[text()="基金合同"]').click()
        time.sleep(0.5)

        # 下一步
        print('点击：下一步')
        self.chrome.find_element_by_xpath('//*[@id="button_next"]/div[2]').click()
        time.sleep(0.5)


        #对websection数量进行判断
        print('对websection数量进行判断，并修改章节名字')
        self.section_alter(data)
        #print(2)

        # 录入所有部分
        i = 1
        for c in data:
            #第i章节
            print('*' * 10, '写入第%d章节:  开始' % (i), '*' * 10)
            if i == 1:
                #第一章节
                self.chrome.switch_to.parent_frame()
                self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_2')
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
                #time.sleep(0.5)
                # 判断web段落数是否满足<=local段落数
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：参数
                    for j in range(localNumber):
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath('//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()#当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath('//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span'%(j+1)).click()
                        self.chrome.find_element_by_xpath('//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span'%(j+1)).click()
                else:
                    continue
            else:
                # 以后的第二章节
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_%d_a"]' % (i)).click()
                time.sleep(0.5)
                # 对段落数进行增补
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：对每个段落复制文字
                    for j in range(localNumber):
                        # 第一段不需要退出当前frame，第二段之后需要先退出当前frame
                        #if j == 0:
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath('//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()#当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            print('*' * 10, '写入第%d章节:  结束' % (i), '*' * 10)
            i += 1  # 完成一个章节+1


        # 保存
        print('保存录入：点击【保存】按钮')
        self.chrome.find_element_by_id('//*[@id="contract_layout_vertical_south"]//div[text()="保存"]').click()

        print('保存录入：点击弹窗【确定】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        print('保存录入：点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        #更新网页
        self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="添加段落"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="删除段落"]').click()
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        # 生成Word文件
        print('生成Word文件')
        self.chrome.find_element_by_id('folderadd').click()
        time.sleep(30)
        print('生成成功，点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 下载Word文件
        print('下载Word文件')
        self.chrome.find_element_by_id('//*[@id="contract_layout_vertical_south"]//div[text()="下载"]').click()
        time.sleep(3)
        self.status = True
        #self.tearDown()

    def write2(self,  web_filename, data, information):
        '''
        业务流程：产品存续-产品名-招募说明书
        :param web_filename: 产品名字
        :param data: 录入的字典数据
        :param
        :return:文件录入成功，下载成功
        '''
        # 链接URL
        print('打开：URL')
        self.chrome.get(self.url)
        time.sleep(2)

        # 产品管理
        print('点击：产品管理')
        # self.chrome.find_element_by_xpath('//*[@id="kui-11"]/div[2]/table/tbody/tr/td/ul/li[1]/span').click()
        self.chrome.find_element_by_xpath('//*[@id="kui_frame_vertical_north"]/div/ul/li[3]').click()
        time.sleep(0.5)

        # 产品开发
        print('点击：产品存续')
        self.chrome.find_element_by_xpath('//*[@id="kui-11"]//span[text()="产品存续"]').click()
        time.sleep(0.5)

        # 产品名字：博时鑫禧灵活配置混合型证券投资基金
        print('点击：%s' % web_filename)
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_1')
        self.chrome.find_element_by_xpath('//*[@id="kui-13"]').send_keys(web_filename)
        time.sleep(0.5)
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="搜索"]').click()
        time.sleep(0.5)

        # 修改基本信息
        print('修改基本信息')
        self.alter_information(information)

        # 报会文件制作
        print('点击：法律文件修改')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="法律文件修改"]').click()
        # time.sleep(3)

        # 合同文件类型：招募说明书
        print('点击：招募说明书')
        self.chrome.find_element_by_xpath(
            '//*[@id="fileTable_listContainer"]//span[text()="招募说明书"]').click()
        time.sleep(0.5)

        # 下一步
        print('点击：下一步')
        self.chrome.find_element_by_xpath('//*[@id="button_next"]/div[2]').click()
        time.sleep(0.5)

        # 对websection数量进行判断
        print('对websection数量进行判断，并修改章节名字')
        self.section_alter(data)
        # print(2)

        # 录入所有部分
        i = 1
        for c in data:
            # 第i章节
            print('*' * 10, '写入第%d章节:  开始' % (i), '*' * 10)
            if i == 1:
                # 第一章节
                self.chrome.switch_to.parent_frame()
                self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_2')
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
                # time.sleep(0.5)
                # 判断web段落数是否满足<=local段落数
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：参数
                    for j in range(localNumber):
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()  # 当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            else:
                # 以后的第二章节
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_%d_a"]' % (i)).click()
                time.sleep(0.5)
                # 对段落数进行增补
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：对每个段落复制文字
                    for j in range(localNumber):
                        # 第一段不需要退出当前frame，第二段之后需要先退出当前frame
                        # if j == 0:
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()  # 当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            print('*' * 10, '写入第%d章节:  结束' % (i), '*' * 10)
            i += 1  # 完成一个章节+1

        # 保存
        print('保存录入：点击【保存】按钮')
        self.chrome.find_element_by_xpath('//*[@id="contract_layout_vertical_south"]//div[text()="保存"]').click()

        print('保存录入：点击弹窗【确定】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        print('保存录入：点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 更新网页
        self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="添加段落"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="删除段落"]').click()
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        # 生成Word文件
        print('生成Word文件')
        self.chrome.find_element_by_id('folderadd').click()
        time.sleep(30)
        print('生成成功，点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 下载Word文件
        print('下载Word文件')
        self.chrome.find_element_by_xpath('//*[@id="contract_layout_vertical_south"]//div[text()="下载"]').click()
        time.sleep(3)
        self.status = True
        # self.tearDown()

    def write3(self,  web_filename, data, information, date):
        '''
        业务流程：产品存续-产品名-更新招募说明书
        :param web_filename: 产品名字
        :param data: 录入的字典数据
        :param
        :return:文件录入成功，下载成功
        '''
        # 链接URL
        print('打开：URL')
        self.chrome.get(self.url)
        time.sleep(2)

        # 产品管理
        print('点击：产品管理')
        # self.chrome.find_element_by_xpath('//*[@id="kui-11"]/div[2]/table/tbody/tr/td/ul/li[1]/span').click()
        self.chrome.find_element_by_xpath('//*[@class="kui-inline-block" and text()="产品管理"]').click()
        time.sleep(0.5)

        # 产品开发
        print('点击：产品存续')
        self.chrome.find_element_by_xpath('//*[@id="kui-11"]//span[text()="产品存续"]').click()
        time.sleep(0.5)

        # 产品名字：博时鑫禧灵活配置混合型证券投资基金
        print('点击：%s' % web_filename)
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_1')
        self.chrome.find_element_by_xpath('//*[@id="kui-13"]').send_keys(web_filename)
        time.sleep(0.5)
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="搜索"]').click()
        time.sleep(0.5)

        # 修改基本信息
        print('修改基本信息')
        self.alter_information(information)

        # 报会文件制作
        print('点击：招募文件更新')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="招募文件更新"]').click()
        # time.sleep(3)

        # 填充更新招募说明书的时间
        print('填充更新招募说明书的时间')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="招募制作"]').click()

        try:
            self.chrome.find_element_by_xpath('//*[@id="applyyearid_btn"]').click()
            self.chrome.find_element_by_xpath(
                '//li[@class="kui-select-list-opt-item" and text()="%s"]' % (date[0][0])).click()

            self.chrome.find_element_by_id('number_btn').click()
            self.chrome.find_element_by_xpath('//*[@id="number_list"]/ul/li[%s]' % (date[0][1])).click()

            self.chrome.find_element_by_xpath('//*[@id="financeendday_btn"]').click()
            time.sleep(5)

            self.chrome.find_element_by_xpath('//*[@id="financeendday_text"]').clear()
            self.chrome.find_element_by_xpath('//*[@id="financeendday_text"]').send_keys(date[1][0])

            if len(date[1]) == 2:
                self.chrome.find_element_by_xpath('//*[@id="infoendday_text"]').clear()
                self.chrome.find_element_by_xpath('//*[@id="infoendday_text"]').send_keys(date[1][1])
            else:
                pass
            # 点击确定按钮
            time.sleep(3)
            self.chrome.find_element_by_xpath('//*[@id="kui-59"]/div[2]').click()
            time.sleep(3)
        except NoSuchElementException as e:
            print('不用填充更新招募说明书的时间：' + '\n' + str(e))
            self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()
            self.chrome.find_element_by_id('editbtns').click()

        # 对websection数量进行判断
        print('对websection数量进行判断，并修改章节名字')
        self.section_alter(data)
        # print(2)

        # 录入所有部分
        i = 1
        for c in data:
            # 第i章节
            print('*' * 10, '写入第%d章节:  开始' % (i), '*' * 10)
            if i == 1:
                # 第一章节
                self.chrome.switch_to.parent_frame()
                self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_2')
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
                # time.sleep(0.5)
                # 判断web段落数是否满足<=local段落数
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：参数
                    for j in range(localNumber):
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()  # 当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            else:
                # 以后的第二章节
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_%d_a"]' % (i)).click()
                time.sleep(0.5)
                # 对段落数进行增补
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：对每个段落复制文字
                    for j in range(localNumber):
                        # 第一段不需要退出当前frame，第二段之后需要先退出当前frame
                        # if j == 0:
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()  # 当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            print('*' * 10, '写入第%d章节:  结束' % (i), '*' * 10)
            i += 1  # 完成一个章节+1

        # 保存
        print('保存录入：点击【保存】按钮')
        self.chrome.find_element_by_xpath('//*[@id="contract_layout_vertical_south"]//div[text()="保存"]').click()

        print('保存录入：点击弹窗【确定】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        print('保存录入：点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 更新网页
        self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="添加段落"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="删除段落"]').click()
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        # 生成Word文件
        print('生成Word文件')
        self.chrome.find_element_by_id('folderadd').click()
        time.sleep(60)
        print('生成成功，点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 下载Word文件
        print('下载Word文件')
        self.chrome.find_element_by_xpath('//*[@id="contract_layout_vertical_south"]//div[text()="下载"]').click()
        time.sleep(3)
        self.status = True
        # self.tearDown()

    def write4(self, web_filename, data):
        '''
        业务流程：产品存续-产品名-基金合同
        :param web_filename: 产品名字
        :param data: 录入的字典数据
        :return: 录入成功，文件下载成功
        '''

        #链接URL
        print('打开：URL')
        self.chrome.get(self.url)
        time.sleep(2)

        # 产品管理
        print('点击：产品管理')
        # self.chrome.find_element_by_xpath('//*[@id="kui-11"]/div[2]/table/tbody/tr/td/ul/li[1]/span').click()
        self.chrome.find_element_by_xpath('//*[@class="kui-inline-block" and text()="产品管理"]').click()
        time.sleep(0.5)

        # 产品开发
        print('点击：产品存续')
        self.chrome.find_element_by_xpath('//*[@id="kui-11"]//span[text()="产品存续"]').click()
        time.sleep(0.5)

        # 产品名字
        print('点击：%s' % web_filename)
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_1')
        self.chrome.find_element_by_xpath('//*[@id="kui-13"]').send_keys(web_filename)
        time.sleep(0.5)
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="搜索"]').click()
        time.sleep(0.5)

        # 报会文件制作
        print('点击：法律文件修改')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="法律文件修改"]').click()
        # time.sleep(3)

        # 基金合同：合同文件序号参数
        print('点击：基金合同')
        self.chrome.find_element_by_xpath('//*[@id="fileTable_listContainer"]//span[text()="基金合同"]').click()
        time.sleep(0.5)

        # 下一步
        print('点击：下一步')
        self.chrome.find_element_by_xpath('//*[@id="button_next"]/div[2]').click()
        time.sleep(0.5)


        #对websection数量进行判断
        print('对websection数量进行判断，并修改章节名字')
        self.section_alter(data)
        #print(2)

        # 录入所有部分
        i = 1
        for c in data:
            #第i章节
            print('*' * 10, '写入第%d章节:  开始' % (i), '*' * 10)
            if i == 1:
                #第一章节
                self.chrome.switch_to.parent_frame()
                self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_2')
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
                #time.sleep(0.5)
                # 判断web段落数是否满足<=local段落数
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：参数
                    for j in range(localNumber):
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath('//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()#当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath('//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span'%(j+1)).click()
                        self.chrome.find_element_by_xpath('//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span'%(j+1)).click()
                else:
                    continue
            else:
                # 以后的第二章节
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_%d_a"]' % (i)).click()
                time.sleep(0.5)
                # 对段落数进行增补
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：对每个段落复制文字
                    for j in range(localNumber):
                        # 第一段不需要退出当前frame，第二段之后需要先退出当前frame
                        #if j == 0:
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath('//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()#当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            print('*' * 10, '写入第%d章节:  结束' % (i), '*' * 10)
            i += 1  # 完成一个章节+1


        # 保存
        print('保存录入：点击【保存】按钮')
        self.chrome.find_element_by_id('//*[@id="contract_layout_vertical_south"]//div[text()="保存"]').click()

        print('保存录入：点击弹窗【确定】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        print('保存录入：点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        #更新网页
        self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="添加段落"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="删除段落"]').click()
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        # 生成Word文件
        print('生成Word文件')
        self.chrome.find_element_by_id('folderadd').click()
        time.sleep(30)
        print('生成成功，点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 下载Word文件
        print('下载Word文件')
        self.chrome.find_element_by_id('//*[@id="contract_layout_vertical_south"]//div[text()="下载"]').click()
        time.sleep(3)
        self.status = True
        #self.tearDown()

    def write5(self, web_filename, data):
        '''
        业务流程：产品募集-产品名-基金合同
        :param web_filename: 产品名字
        :param data: 录入的字典数据
        :return: 录入成功，文件下载成功
        '''

        # 链接URL
        print('打开：URL')
        self.chrome.get(self.url)
        time.sleep(2)

        # 产品管理
        print('点击：产品管理')
        # self.chrome.find_element_by_xpath('//*[@id="kui-11"]/div[2]/table/tbody/tr/td/ul/li[1]/span').click()
        self.chrome.find_element_by_xpath('//*[@class="kui-inline-block" and text()="产品管理"]').click()
        time.sleep(0.5)

        # 产品开发
        print('点击：产品募集')
        self.chrome.find_element_by_xpath('//*[@id="kui-11"]//span[text()="产品募集"]').click()
        time.sleep(0.5)

        # 产品名字
        print('点击：%s' % web_filename)
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_1')
        self.chrome.find_element_by_xpath('//*[@id="kui-13"]').send_keys(web_filename)
        time.sleep(0.5)
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="搜索"]').click()
        time.sleep(0.5)

        # 报会文件制作
        print('点击：法律文件修改')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="法律文件修改"]').click()
        # time.sleep(3)

        # 基金合同：合同文件序号参数
        print('点击：基金合同')
        self.chrome.find_element_by_xpath('//*[@id="fileTable_listContainer"]//span[text()="基金合同"]').click()
        time.sleep(0.5)

        # 下一步
        print('点击：下一步')
        self.chrome.find_element_by_xpath('//*[@id="button_next"]/div[2]').click()
        time.sleep(0.5)

        # 对websection数量进行判断
        print('对websection数量进行判断，并修改章节名字')
        self.section_alter(data)
        # print(2)

        # 录入所有部分
        i = 1
        for c in data:
            # 第i章节
            print('*' * 10, '写入第%d章节:  开始' % (i), '*' * 10)
            if i == 1:
                # 第一章节
                self.chrome.switch_to.parent_frame()
                self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_2')
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
                # time.sleep(0.5)
                # 判断web段落数是否满足<=local段落数
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：参数
                    for j in range(localNumber):
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()  # 当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            else:
                # 以后的第二章节
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_%d_a"]' % (i)).click()
                time.sleep(0.5)
                # 对段落数进行增补
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：对每个段落复制文字
                    for j in range(localNumber):
                        # 第一段不需要退出当前frame，第二段之后需要先退出当前frame
                        # if j == 0:
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()  # 当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            print('*' * 10, '写入第%d章节:  结束' % (i), '*' * 10)
            i += 1  # 完成一个章节+1

        # 保存
        print('保存录入：点击【保存】按钮')
        self.chrome.find_element_by_id('//*[@id="contract_layout_vertical_south"]//div[text()="保存"]').click()

        print('保存录入：点击弹窗【确定】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        print('保存录入：点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 更新网页
        self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="添加段落"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="删除段落"]').click()
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        # 生成Word文件
        print('生成Word文件')
        self.chrome.find_element_by_id('folderadd').click()
        time.sleep(30)
        print('生成成功，点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 下载Word文件
        print('下载Word文件')
        self.chrome.find_element_by_id('//*[@id="contract_layout_vertical_south"]//div[text()="下载"]').click()
        time.sleep(3)
        self.status = True
        # self.tearDown()

    def write6(self,  web_filename, data, information):
        '''
        业务流程：产品开发-产品名-招募说明书
        :param web_filename: 产品名字
        :param data: 录入的字典数据
        :param
        :return:文件录入成功，下载成功
        '''
        # 链接URL
        print('打开：URL')
        self.chrome.get(self.url)
        time.sleep(2)

        # 产品管理
        print('点击：产品管理')
        # self.chrome.find_element_by_xpath('//*[@id="kui-11"]/div[2]/table/tbody/tr/td/ul/li[1]/span').click()
        self.chrome.find_element_by_xpath('//*[@class="kui-inline-block" and text()="产品管理"]').click()
        time.sleep(0.5)

        # 产品开发
        print('点击：产品开发')
        self.chrome.find_element_by_xpath('//*[@id="kui-11"]//span[text()="产品开发"]').click()
        time.sleep(0.5)

        # 产品名字：博时鑫禧灵活配置混合型证券投资基金
        print('点击：%s' % web_filename)
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_1')
        self.chrome.find_element_by_xpath('//*[@id="kui-12"]').send_keys(web_filename)
        time.sleep(0.5)
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="搜索"]').click()
        time.sleep(0.5)

        # 修改基本信息
        print('修改基本信息')
        self.alter_information2(information)

        # 报会文件制作
        print('点击：报会文件制作')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="报会文件制作"]').click()


        # 合同文件类型：招募说明书
        print('点击：招募说明书')
        self.chrome.find_element_by_xpath(
            '//*[@id="fileTable_listContainer"]//span[text()="招募说明书"]').click()
        time.sleep(0.5)

        # 下一步
        print('点击：下一步')
        self.chrome.find_element_by_xpath('//*[@id="button_next"]/div[2]').click()
        time.sleep(0.5)

        # 对websection数量进行判断
        print('对websection数量进行判断，并修改章节名字')
        self.section_alter(data)
        # print(2)

        # 录入所有部分
        i = 1
        for c in data:
            # 第i章节
            print('*' * 10, '写入第%d章节:  开始' % (i), '*' * 10)
            if i == 1:
                # 第一章节
                self.chrome.switch_to.parent_frame()
                self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_2')
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
                # time.sleep(0.5)
                # 判断web段落数是否满足<=local段落数
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：参数
                    for j in range(localNumber):
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()  # 当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            else:
                # 以后的第二章节
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_%d_a"]' % (i)).click()
                time.sleep(0.5)
                # 对段落数进行增补
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：对每个段落复制文字
                    for j in range(localNumber):
                        # 第一段不需要退出当前frame，第二段之后需要先退出当前frame
                        # if j == 0:
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()  # 当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            print('*' * 10, '写入第%d章节:  结束' % (i), '*' * 10)
            i += 1  # 完成一个章节+1

        # 保存
        print('保存录入：点击【保存】按钮')
        self.chrome.find_element_by_xpath('//*[@id="contract_layout_vertical_south"]//div[text()="保存"]').click()

        print('保存录入：点击弹窗【确定】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        print('保存录入：点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 更新网页
        self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="添加段落"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="删除段落"]').click()
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        # 生成Word文件
        print('生成Word文件')
        self.chrome.find_element_by_id('folderadd').click()
        time.sleep(30)
        print('生成成功，点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 下载Word文件
        print('下载Word文件')
        self.chrome.find_element_by_xpath('//*[@id="contract_layout_vertical_south"]//div[text()="下载"]').click()
        time.sleep(3)
        self.status = True
        # self.tearDown()

    def write7(self,  web_filename, data, information):
        '''
        业务流程：产品募集-产品名-招募说明书
        :param web_filename: 产品名字
        :param data: 录入的字典数据
        :param
        :return:文件录入成功，下载成功
        '''
        # 链接URL
        print('打开：URL')
        self.chrome.get(self.url)
        time.sleep(2)

        # 产品管理
        print('点击：产品管理')
        # self.chrome.find_element_by_xpath('//*[@id="kui-11"]/div[2]/table/tbody/tr/td/ul/li[1]/span').click()
        self.chrome.find_element_by_xpath('//*[@id="kui_frame_vertical_north"]/div/ul/li[3]').click()
        time.sleep(0.5)

        # 产品开发
        print('点击：产品募集')
        self.chrome.find_element_by_xpath('//*[@id="kui-11"]//span[text()="产品募集"]').click()
        time.sleep(0.5)

        # 产品名字：博时鑫禧灵活配置混合型证券投资基金
        print('点击：%s' % web_filename)
        self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_1')
        self.chrome.find_element_by_xpath('//*[@id="kui-13"]').send_keys(web_filename)
        time.sleep(0.5)
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="搜索"]').click()
        time.sleep(0.5)

        # 修改基本信息
        print('修改基本信息')
        self.alter_information2(information)

        # 报会文件制作
        print('点击：法律文件修改')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="法律文件修改"]').click()
        # time.sleep(3)

        # 合同文件类型：招募说明书
        print('点击：招募说明书')
        self.chrome.find_element_by_xpath(
            '//*[@id="fileTable_listContainer"]//span[text()="招募说明书"]').click()
        time.sleep(0.5)

        # 下一步
        print('点击：下一步')
        self.chrome.find_element_by_xpath('//*[@id="button_next"]/div[2]').click()
        time.sleep(0.5)

        # 对websection数量进行判断
        print('对websection数量进行判断，并修改章节名字')
        self.section_alter(data)
        # print(2)

        # 录入所有部分
        i = 1
        for c in data:
            # 第i章节
            print('*' * 10, '写入第%d章节:  开始' % (i), '*' * 10)
            if i == 1:
                # 第一章节
                self.chrome.switch_to.parent_frame()
                self.chrome.switch_to.frame('ui-frame-main-frameTab_tabContent_2')
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
                # time.sleep(0.5)
                # 判断web段落数是否满足<=local段落数
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：参数
                    for j in range(localNumber):
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()  # 当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            else:
                # 以后的第二章节
                self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_%d_a"]' % (i)).click()
                time.sleep(0.5)
                # 对段落数进行增补
                localNumber = len(data[c])
                flag = self.paragraph_judge(localNumber)
                if flag:
                    # 复制文字：对每个段落复制文字
                    for j in range(localNumber):
                        # 第一段不需要退出当前frame，第二段之后需要先退出当前frame
                        # if j == 0:
                        print('第%d段：数据录入web' % (j + 1))
                        frame = self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[2]/iframe' % (j + 1))
                        self.chrome.switch_to.frame(frame)
                        self.chrome.find_element_by_xpath('/html/body').clear()
                        self.chrome.find_element_by_xpath('/html/body').send_keys(data[c][j])
                        self.chrome.switch_to.parent_frame()  # 当前iframe:ui-frame-main-frameTab_tabContent_2
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[13]/span' % (j + 1)).click()
                        self.chrome.find_element_by_xpath(
                            '//*[@id="%d_formcontrol"]/div/div[1]/span[14]/span' % (j + 1)).click()
                else:
                    continue
            print('*' * 10, '写入第%d章节:  结束' % (i), '*' * 10)
            i += 1  # 完成一个章节+1

        # 保存
        print('保存录入：点击【保存】按钮')
        self.chrome.find_element_by_xpath('//*[@id="contract_layout_vertical_south"]//div[text()="保存"]').click()

        print('保存录入：点击弹窗【确定】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        print('保存录入：点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 更新网页
        self.chrome.find_element_by_xpath('//*[@id="insChapterTree_tree_1_a"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="添加段落"]').click()
        e = self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="段落"]')
        time.sleep(0.5)
        e.click()
        self.chrome.find_element_by_xpath('//div[@class="kui-split-menu-text" and text()="删除段落"]').click()
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="确定"]').click()

        # 生成Word文件
        print('生成Word文件')
        self.chrome.find_element_by_id('folderadd').click()
        time.sleep(30)
        print('生成成功，点击【关闭】按钮')
        self.chrome.find_element_by_xpath('//div[@class="kui-button-text" and text()="关闭"]').click()

        # 下载Word文件
        print('下载Word文件')
        self.chrome.find_element_by_xpath('//*[@id="contract_layout_vertical_south"]//div[text()="下载"]').click()
        time.sleep(3)
        self.status = True
        # self.tearDown()

if __name__ == "__main__":
    filename = '博时中证500指数增强型证券投资基金招募说明书.docx'
    d = WordFileClean(filename)
    data = d.get_data()
    information = d.alter_information()
    print(d)

    a = WordWriteWeb()
    a.write2(config.web_filename[1], data, information)
