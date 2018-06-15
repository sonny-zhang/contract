from Code.word_separate import WordFileClean
from Code.word_writeweb import WordWriteWeb
from Code import config
import logging.config
import os
import copy
import sys
import time
import shutil


path_localfile = os.path.join(os.path.abspath(os.path.dirname(os.path.dirname(__file__))), 'localfile')
path_contras_container = os.path.join(os.path.abspath(os.path.dirname(os.path.dirname(__file__))), 'contras_container')
path_downloadfile = os.path.join(os.path.abspath(os.path.dirname(os.path.dirname(__file__))), 'downloadfile')

class TwoWords:
    '''两个Word文件信息比对'''

    def __init__(self):
        self.path_local = path_localfile
        self.path_web = config.filepath_web
        self.path_download = path_downloadfile
        self.path_contras_container = path_contras_container
        self.web_filename = config.web_filename
        self.product_stage = config.product_stage
        self.numture = 0
        self.numfalss = 0
        self.numinvalid = 0
        self.name = ''
        logpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logger.conf')
        logging.config.fileConfig(logpath)
        self.logger = logging.getLogger('cse')
        


    def __str__(self):

        msg = ' sucess：' + str(self.numture)
        msg += '\n'+ ' failure：'+ str(self.numfalss)
        msg += '\n' + ' invalid：' + str(self.numinvalid)
        return msg

    def fiel_together_path(self, path, fileName):
        '''
        文件名和路径拼接功能
        :param path: 路径
        :param fileName:文件名
        :return: 文件路径
        '''
        pathName = path + '\\' + fileName
        return pathName

    def selectFlashback_allfiles(self, path):
        '''
        获得文件夹里的所有文件,并倒叙排列
        :param path: 文件夹路径
        :return: 所有文件名的list
        '''
        list = os.listdir(path)
        list.sort(key=lambda x: int(x[:3]))
        return list

    def file_time(self):
        '''
        获取文件家里的时间最新的文件名
        :param DIR: 文件夹路径
        :return: 文件名字
        '''
        iterms = os.listdir(self.path_web)
        iterms = sorted(iterms, key=lambda x: os.stat(self.path_web + "/" + x).st_atime, reverse=True)
        new_file = iterms[0]
        return new_file

    def del_element(self, list):
        '''
        删除list里面不需要的元素
        :param list:
        :return:
        '''
        n = list.count('')
        for i in range(n):
            list.remove('')
        n = list.count('\n')
        for i in range(n):
            list.remove('\n')
        return list

    def interception_list(self, list):
        '''
        对list里面的元素进行截取并去重，只要前3个字符
        :param str: 要截取的字符
        :return: 只有三位字符的list
        '''
        newList = []
        for s in list:
            newList.append(s[:3])
        newList = set(newList)  #消除重复元素
        return str(newList)

    def is_same(self, Afiles, Bfiles):
        '''
        判断list1和list2里面的元素是否一样，且各自的list没有相同的
        :param Afiles:
        :param Bfiles:
        :return: True;False
        '''
        A = self.interception_list(Afiles)
        B = self.interception_list(Bfiles)
        self.logger.info('localfile： ' + str(Afiles))
        self.logger.info('downloadfile： ' + str(Bfiles))
        if len(A) == len(B):
            # 先判断数量是否一致
            for i in range(len(A)):
                if A[i] == B[i]:
                    pass
                else:
                    self.logger.error('Error： ' + '两个文件夹里的文件名，有不对应的，请修改...')
                    sys.exit()
                    # 判断Afiles里每个str前三个字符是否和Bfiles一致，有不一致的就要提示，并退出
        else:
            self.logger.error('Error： ' + 'localfile和downloadfile文件数量不一致，请检查...')
            sys.exit()


    def allFile(self, localFath, webFath):
        '''
        :param localFath: Word合同存放路径
        :param webFath: 下载合同的路径
        :return: listLocar,listWeb
        '''
        pass

    def data(self, file_path):
        '''
        第一次写入：文件写入页面功能
        :param file_path: 要录入的.docx文件路径
        :param webname: web网页的产品名字
        :return:
        '''
        self.clean = WordFileClean(file_path)
        data = self.clean.get_data()
        print(self.clean)
        return data

    def information(self):
        '''
        要更新的数据：律师事务所，会计事务所
        :return:list
        '''
        mation = self.clean.alter_information()
        return mation

    def date_information(self):
        '''
        获得要填入的日期数据：文件名里的日期，财务截止日期
        :return: list集合里两个list
        '''
        date = self.clean.time_information()
        return date

    def contrast_file(self, fileApath, fileBpath):
        '''
        对两个文件进行比对
        :param fileA:文件A路径
        :param fileB:文件B路劲
        :return: 将错误信息写入日志
        '''
        fileA = WordFileClean(fileApath).get_data()
        fileB = WordFileClean(fileBpath).get_data()
        i = 1
        for p in fileA:
            # print('*' * 8, '第%d章节比对：开始' % (i), '*' * 8)
            self.logger.info('*' * 20+'第%d章节比对：开始' % (i)+'*' * 20)
            Avalues = copy.deepcopy(fileA[p])  # 深copy
            try:
                Bvalues = fileB[p]
            except KeyError:
                # print('Error:')
                # print('       没有在【%s】找到该章节：%s'%(fileBpath, p))
                # print('*' * 8, '第%d章节比对：结束' % (i), '*' * 8)
                self.logger.error('Error:'+'\n'+'\t\t，没有在【%s】找到该章节：%s'%(fileBpath, p))
                self.logger.info('*' * 20+'第%d章节比对：结束' % (i)+'*' * 20)
                i += 1
                continue

            for j in range(len(Avalues)):
                # 该章节的一个段落
                Aonevalue = self.del_element(Avalues[j])
                Bonevalue = self.del_element(Bvalues[j])
                if len(Aonevalue) == len(Bonevalue):
                    # 段落里的行数一致进行比较
                    # print('*' * 5, '第%d段，每一行对比: start' % (j + 1), '*' * 5)
                    self.logger.info('*' * 10+ '第%d段，每一行对比: start' % (j + 1)+ '*' * 10)
                    for r in range(len(Aonevalue)):
                        # 该段落的每一行
                        if Aonevalue[r] == Bonevalue[r]:
                            self.numture += 1
                        else:
                            # print('第%d行，两者数据不一致，请检查下...' % (r + 1))
                            # print(' local:', '\n', Aonevalue[r])
                            # print(' download:', '\n', Bonevalue[r])
                            self.logger.error('第%d行，两者数据不一致，请检查下...' % (r + 1))
                            self.logger.error(' local:' + '\n' + Aonevalue[r])
                            self.logger.error(' download:'+ '\n'+ Bonevalue[r])
                            self.numfalss += 1

                    self.logger.info('*' * 10+ '第%d段，每一行对比: end' % (j + 1)+ '*' * 10)
                else:
                    if i == len(fileA) and j == len(Avalues)-1:
                        # 最后一章节，最后一段需要处理
                        end = len(Aonevalue)
                        Bonevalue = Bonevalue[0:end]

                        self.logger.info('*' * 10+ '第%d段，每一行对比: start' % (j + 1)+ '*' * 10)
                        for r in range(len(Aonevalue)):
                            # 该段落的每一行
                            if Aonevalue[r] == Bonevalue[r]:
                                self.numture += 1
                            else:
                                # print('第%d行，两者数据不一致，请检查下...' % (r + 1))
                                # print(' local:', '\n', Aonevalue[r])
                                # print(' download:', '\n', Bonevalue[r])
                                self.logger.error('第%d行，两者数据不一致，请检查下...' % (r + 1))
                                self.logger.error(' local:' + '\n' + Aonevalue[r])
                                self.logger.error(' download:' + '\n' + Bonevalue[r])
                                self.numfalss += 1
                        # print('*' * 5, '第%d段，每一行对比: end' % (j + 1), '*' * 5)
                        self.logger.info('*' * 10 + '第%d段，每一行对比: start' % (j + 1) + '*' * 10)
                    else:
                        self.numinvalid += 1
                        # print('Error: 第%d段，两者的行数不一致，无法进行对比，请查看...' % (j + 1))
                        # print(' local:', len(Aonevalue), '\n', Aonevalue)
                        # print(' download:', len(Bonevalue), '\n', Bonevalue)
                        self.logger.error('Error: 第%d段，两者的行数不一致，无法进行对比，请查看...' % (j + 1))
                        self.logger.error(' local:'+ str(len(Aonevalue))+ '\n'+ str(Aonevalue))
                        self.logger.error(' download:'+ str(len(Bonevalue))+ '\n'+ str(Bonevalue))

            print('*' * 8, '第%d章节比对：结束' % (i), '*' * 8)
            self.logger.info('*' * 20+ '第%d章节比对：结束' % (i)+ '*' * 20)
            i += 1

    def write_only(self):
        '''
        只对合同进行录入web页面
        :return:
        '''
        files = self.selectFlashback_allfiles(self.path_local)
        self.logger.info('要录入的合同数量为：%d个' % (len(self.web_filename)))

        driver = WordWriteWeb()#创建写入类
        for i in range(len(self.web_filename)):
            self.logger.info('*' * 20 + '第%d次录入web页面: 开始' % (i + 1) + '*' * 20)
            name = files[i]
            self.logger.info('当前使用的Word文件：%s' % (name))
            self.logger.info('当前要录入的web为：%s' % (self.web_filename[i]))
            file_path = self.fiel_together_path(self.path_local, name)
            webname = self.web_filename[i]
            data = self.data(file_path)
            if '基金合同' in name:
                if self.product_stage == '产品开发':
                    driver.write1(webname, data)  # 录入合同
                elif self.product_stage == '产品存续':
                    driver.write4(webname, data)
                elif self.product_stage == '产品募集':
                    driver.write5(webname, data)
                else:
                    self.logger.warning('config配置文件的product_stage参数错误，请检查...')
            elif '基金招募说明书' in name:
                information = self.information()


                if self.product_stage == '产品开发':
                    cunxu = driver.has_cunxu2(webname)
                    if cunxu:
                        driver.write6(webname, data, information)
                    else:
                        self.logger.warning('Warning：*存续期限*为空，请先填写...')
                        self.logger.info('*' * 20 + '第%d次录入web页面: 结束' % (i + 1) + '*' * 20)
                        self.numture += 1
                        continue
                elif self.product_stage == '产品存续':
                    cunxu = driver.has_cunxu(webname)
                    if cunxu:
                        driver.write2(webname, data, information)
                    else:
                        self.logger.warning('Warning：*存续期限*为空，请先填写...')
                        self.logger.info('*' * 20 + '第%d次录入web页面: 结束' % (i + 1) + '*' * 20)
                        self.numture += 1
                        continue
                elif self.product_stage == '产品募集':
                    cunxu = driver.has_cunxu3(webname)
                    if cunxu:
                        driver.write7(webname, data, information)
                    else:
                        self.logger.warning('Warning：*存续期限*为空，请先填写...')
                        self.logger.info('*' * 20 + '第%d次录入web页面: 结束' % (i + 1) + '*' * 20)
                        self.numture += 1
                        continue
                else:
                    self.logger.warning('config配置文件的product_stage参数错误，请检查...')


            elif '基金更新招募说明书' in name:
                cunxu = driver.has_cunxu(webname)
                if cunxu:

                    information = self.information()
                    date = self.date_information()
                    driver.write3(webname, data, information, date)
                else:
                    self.logger.warning('Warning：*存续期限*为空，请先填写...')
                    self.logger.info('*' * 20 + '第%d次录入web页面: 结束' % (i + 1) + '*' * 20)
                    self.numture += 1
                    continue

            else:
                self.logger.warning('Warning: 没有匹配到要写入的web业务流程，请联系脚本维护人员...')
                sys.exit()


            self.numture += 1
            self.logger.info('*' * 20, '第%d次录入web页面: 结束' % (i + 1), '*' * 20)

    def contrast_only(self):
        '''
        仅对文件进行对比
        :return: 对比结果
        '''
        Afiles = self.selectFlashback_allfiles(path_localfile)
        Bfiles = self.selectFlashback_allfiles(path_downloadfile)
        # 获得要对比的文件名
        self.is_same(Afiles, Bfiles)
        # 对文件进行是否合法性判断
        for i in range(len(Afiles)):
            fileApath = self.fiel_together_path(self.path_local, Afiles[i])
            fileBpath = self.fiel_together_path(self.path_download, Bfiles[i])

            self.logger.info('='*30+'第%d次对比： 开始'%(i+1)+'='*30)

            self.contrast_file(fileApath, fileBpath)

            self.logger.info('='*30+'第%d次对比： 结束'%(i+1)+'='*30)
            # 开始比对文件数据是否一致

    def write_and_contrast(self):
        '''
        1,从web_filename获取文件个数，遍历执行录入和对比
        2，先执行录入，将加载的文件放进downloadfile
        3,再将两个文件放进contras_container文件夹
        4，对contras_container里面的文件进行对比
        5，将对比结果记录进文件result.txt文件里
        :return:
        '''
        files = self.selectFlashback_allfiles(self.path_local)
        self.logger.info('要录入的合同数量为：%d个' % (len(self.web_filename)))

        # 判断web_filename和loadfile数量是否一致
        if len(self.web_filename) == len(os.listdir(self.path_local)):
            pass
        else:
            self.logger.error('Error：web_filename配置参数和loadfile文件夹里的数量不一样,请检查哈...')
            sys.exit()

        # downloadfile文件夹是空的，才能进行
        downloaddir = os.listdir(self.path_download)
        if downloaddir:
            self.logger.error('Error： downloadfile文件夹里有文件，请先清理掉...')
            sys.exit()
        else:
            pass

        driver = WordWriteWeb()  # 创建写入类

        # 文件写入页面
        for i in range(len(self.web_filename)):
            self.logger.info('*' * 20 + '第%d次录入web页面: 开始' % (i + 1) + '*' * 20)
            name = files[i]
            self.logger.info('当前使用的Word文件：%s' % (name))
            self.logger.info('当前要录入的web为：%s' % (self.web_filename[i]))
            file_path = self.fiel_together_path(self.path_local, name)
            webname = self.web_filename[i]
            data = self.data(file_path)
            if '基金合同' in name:
                if self.product_stage == '产品开发':
                    driver.write1(webname, data)  # 录入合同
                elif self.product_stage == '产品存续':
                    driver.write4(webname, data)
                elif self.product_stage == '产品募集':
                    driver.write5(webname, data)
                else:
                    self.logger.warning('config配置文件的product_stage参数错误，请检查...')
            elif '基金招募说明书' in name:
                information = self.information()
                if self.product_stage == '产品开发':
                    cunxu = driver.has_cunxu2(webname)
                    if cunxu:
                        driver.write6(webname, data, information)
                    else:
                        self.logger.warning('Warning：*存续期限*为空，请先填写...')
                        self.logger.info('*' * 20 + '第%d次录入web页面: 结束' % (i + 1) + '*' * 20)
                        self.numture += 1
                        continue
                elif self.product_stage == '产品存续':
                    cunxu = driver.has_cunxu(webname)
                    if cunxu:
                        driver.write2(webname, data, information)
                    else:
                        self.logger.warning('Warning：*存续期限*为空，请先填写...')
                        self.logger.info('*' * 20 + '第%d次录入web页面: 结束' % (i + 1) + '*' * 20)
                        self.numture += 1
                        continue
                elif self.product_stage == '产品募集':
                    cunxu = driver.has_cunxu3(webname)
                    if cunxu:
                        driver.write7(webname, data, information)
                    else:
                        self.logger.warning('Warning：*存续期限*为空，请先填写...')
                        self.logger.info('*' * 20 + '第%d次录入web页面: 结束' % (i + 1) + '*' * 20)
                        self.numture += 1
                        continue
                else:
                    self.logger.warning('config配置文件的product_stage参数错误，请检查...')


            elif '基金更新招募说明书' in name:
                cunxu = driver.has_cunxu(webname)
                if cunxu:

                    information = self.information()
                    date = self.date_information()
                    driver.write3(webname, data, information, date)
                else:
                    self.logger.warning('Warning：*存续期限*为空，请先填写...')
                    self.logger.info('*' * 20 + '第%d次录入web页面: 结束' % (i + 1) + '*' * 20)
                    self.numture += 1
                    continue

            else:
                self.logger.warning('Warning: 没有匹配到要写入的web业务流程，请联系脚本维护人员...')
                sys.exit()


            self.numture += 1
            self.logger.info('*' * 20, '第%d次录入web页面: 结束' % (i + 1), '*' * 20)

            time.sleep(5)

            # 将下载的文件copy到download文件夹
            fileName = self.file_time()
            pathA = self.path_web + '\\' + fileName
            pathB = self.path_download + '\\00%d' % (i + 1) + fileName
            shutil.copy(pathA, pathB)
        
        # 文件开始对比
        self.write_only()



if __name__ == "__main__":
    status = TwoWords()
    status.contrast_only()
