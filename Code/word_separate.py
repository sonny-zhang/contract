from Code import config
import docx
import re,copy
import sys, gc
from Code import replaceParameters




class WordFileClean:
    '''将本地Word文件进行数据分离'''

    def __init__(self, filename):
        self.filename = filename
        self.sp = config.separateParameter
        self.data = {}

    def __str__(self):
        if   self.data.keys():
            msg = self.filename + '-' * 6 + 'Word数据已经清理成功'
        else:
            msg = self.filename + '-' * 6 + 'Word数据还没有清理'
        return msg

    def open_wordfile(self):
        '''获得Word文件的paragraphs'''

        self.file = docx.Document(self.filename)
        self.data_false = self.file.paragraphs
        return self.data_false

    def oddNumber(self, count):
        '''获得奇数list
        count参数：奇数的个数
        return：相应n个奇数的list
        '''
        number = 1
        li = []
        for i in range(1, count * 2):
            if number % 2 == 1:
                list.append(number)
            else:
                pass
            number += 1
        return li

    def valueClean(self,list1):
        '''
        对每个字符串进行处理,去除字符串中的换行符\n，去除字符串为‘目录’
        :param list1: list集合
        :return: 插入\n
        '''
        list = []
        for s in list1:
            if '\n' in s:
                list.append(s.replace('\n',''))
            elif s == '目录':
                pass
            else:
                list.append(s)
        return list

    def valueClean2(self,list1):
        '''
        :param list1: list集合
        :return: 遍历后，追加\n
        '''
        v = []

        for i in range(list1.count('')):
            #删除列表里的空格
            list1.remove('')

        for i in range(len(list1)):
            if '\t' in list1[i]:
                #将制表符\t替换成四个英文空格
                list1[i] = list1[i].replace('\t','    ')
            else:
                pass

            if i != len(list1)-1:
                #不是最后一个字符串，需要加换行符\n
                v.append(list1[i].strip())
                v.append('\n')
            else:
                v.append(list1[i].strip())
        return v


    def half_dict(self):
        '''对Word数据进行初次处理，提取章节dict，返回minDict'''

        self.open_wordfile()

        list1 = list(map(lambda x: x.text, self.data_false))
        initialDict = self.valueClean(list1)
        self.minorDict = {}
        endDict = ''

        r1 = re.compile(r'第.*部分')
        r2 = re.compile(r'重要提示|.*【重要提示】.*')
        numberChapter = []
        #numberChapter，段落序号列表
        for i in range(len(initialDict)):
            if r1.match(initialDict[i].strip()):
                numberChapter.append(i)
            elif r2.match(initialDict[i].strip()):
                numberChapter.append(i)
            else:
                continue
        numberChapter.append(len(initialDict))
        #numberChapter：包含所有key所在的序列+最后的一个段落序列
        r3 = re.compile(r'【.*】')
        for i in range(len(numberChapter)-1):
            #章节数
            key = initialDict[numberChapter[i]]
            if r3.match(key):
                #对key进行处理
                key = key[1:-1]
            elif '部分' in key:
                key = ''.join(key.split())    #去除中间空格
            value = initialDict[numberChapter[i]+1:numberChapter[i+1]]
            #value = self.valueClean2(value)
            self.minorDict[key] = value
        return self.minorDict

    def isAllchapter(self):
        '''判断参数separateParameter是否配置正确'''
        helf_data = copy.deepcopy(self.half_dict())

        for key in helf_data:
            if key in config.separateParameter.keys():
                pass
            else:
                print('ERROR:')
                print('     Word文件：【%s】'%(self.filename))
                print('     其中的“%s”章节名 没有在config.py的separateParameter参数中'%(key))
                sys.exit()

        #是否有替换参数
        has_parameter = config.has_parameter
        if has_parameter:
            helf_data = replaceParameters.data(helf_data)
            replaceParameters.printData(helf_data)
        else:
            pass

        return helf_data

    def keyClean(self, key):
        if '部分' in key:
            k = key.split('部分')
            link = '部分  '.join(k)
        else:
            link = key
        return link

    def isInlist(self, str, list):
        '''判断str是否在list中，返回所在list的序号'''
        for sn in range(len(list)):
            if str in list[sn]:
                return sn
            else:
                pass

    def oneC_moreP(self):
        '''一个章节多个段落清洗，返回一个章节的value'''
        oneChapter = []
        if self.oneChapternumber == 1:
            for i in range(2):
                if i == 0:
                    start = 0
                    end = self.isInlist(config.separateParameter[self.t][0], self.value)
                    list = self.value[start:end]
                    list = self.valueClean2(list)
                    oneChapter.append(list)
                    continue
                else:
                    start = self.isInlist(config.separateParameter[self.t][0], self.value)
                    end = len(self.value)
                    list = self.value[start:end]
                    list = self.valueClean2(list)
                    oneChapter.append(list)
                    continue
        else:
            for i in range(self.oneChapternumber+1):
                if i == 0:
                    start = 0
                    end = self.isInlist(config.separateParameter[self.t][0], self.value)
                    list = self.value[start:end]
                    list = self.valueClean2(list)
                    oneChapter.append(list)
                    continue
                elif i == self.oneChapternumber:
                    start = self.isInlist(config.separateParameter[self.t][i - 1], self.value)
                    end = len(self.value)
                    list = self.value[start:end]
                    list = self.valueClean2(list)
                    oneChapter.append(list)
                    continue
                else:
                    start = self.isInlist(config.separateParameter[self.t][i - 1], self.value)
                    end = self.isInlist(config.separateParameter[self.t][i], self.value)
                    list = self.value[start:end]
                    list = self.valueClean2(list)
                    oneChapter.append(list)
                    continue
        return oneChapter

    def oneC_oneP(self, list):
        '''
        对章节只有一个段落进行处理
        :param list:参数是整段的list
        :return: 返回list
        '''
        value = self.valueClean2(list)
        return value

    def is_str(self, str):
        di = {}
        if '通力律师事务所' in str:
            di['律师事务所'] = '上海市通力律师事务所'
        elif '国浩律师集团' in str:
            di['律师事务所'] = '国浩律师集团(北京)事务所'
        elif '上海源泰律师事务所' in str:
            di['律师事务所'] = '上海源泰律师事务所'
        elif '普华永道中天会计师事务所' in str:
            di['会计师事务所'] = '普华永道中天会计师事务所(特殊普通合伙）'
        else:
            pass
        return  di

    def get_data(self):
        '''得到最终的清洗后的data'''
        helf_data = self.isAllchapter()
        chapterNumber = len(helf_data)
        #章节数
        n = 0
        for self.t in helf_data:
            #一个章节的标题
            n += 1
            #n章节序数
            self.oneChapternumber = len(config.separateParameter[self.t])
            #一个章节中的段落数
            self.value = helf_data[self.t]
            if self.oneChapternumber == 0:
                #段落数不是1个,走特定方法
                list2 = []
                self.t = self.keyClean(self.t)
                list2.append(self.oneC_oneP(self.value))
                self.data[self.t] = list2
            else:
                #段落数不是1个
                oneChapter = self.oneC_moreP()
                self.t = self.keyClean(self.t)
                self.data[self.t] = oneChapter
        return self.data

    def alter_information(self):
        if self.minorDict:
            data = self.minorDict
        else:
            data = self.half_dict()

        information = {}
        for k in data:
            if '相关服务机构' in k:
                paragraph = data[k]
                #paragraph是所有的段落
                for one in paragraph:
                    #对一个章节里的所有字符串进行遍历
                    di = self.is_str(one)
                    information.update(di)
            else:
                continue
        if '律师事务所' and '会计师事务所' in information.keys():
            del data
            gc.collect()
            return information
        elif '律师事务所' or '会计师事务所' in information.keys():
            print('Error：只找到“律师事务所”，“会计师事务所”其中一个')
            sys.exit()
        else:
            print('Error：没有找到“律师事务所”和“会计师事务所”')
            sys.exit()

    def time_information(self):
        a1 = re.findall(r'\d\d*', self.filename)
        a2 = []
        #print(self.filename, a1)

        b1 = re.findall(r'所载内容截止日为(.*)日', str(self.half_dict()['重要提示']))
        b2 = re.findall(r'\d\d*', str(b1))
        b3 = []


        list_ = []

        i = False
        for n in a1:
            if len(n)==4:
                a2.append(n)
                i = True
                continue
            if i:
                a2.append(n)
                break

        if len(b2) == 6:
            t1 = b2[0] + '-' + b2[1] + '-' + b2[2]
            t2 = b2[3] + '-' + b2[4] + '-' + b2[5]
            b3.append(t2)       #先添加财务截止日
            b3.append(t1)       #在添加内容截止日
        elif len(b2) == 3:
            t1 = b2[0] + '-' + b2[1] + '-' + b2[2]
            b3.append(t1)
        else:
            print('日期替换没有获得正确的数据')

        list_.append(a2)
        list_.append(b3)
        return list_



if __name__ == '__main__':
    filename = '博时鑫泽灵活配置混合型证券投资基金更新招募说明书2017年第2号（正文） (1).docx'
    a = WordFileClean(filename)
    #h = a.half_dict()
    # print(h)
    d = a.get_data()
    c = a.alter_information()
    '''
    for i in d.keys():
        #print(i)
        print(i,len(d[i]))

    #print(d['第三部分  基金管理人'])
    #print(d['第三部分  基金托管人'][3])
    #print(d['第一部分  绪言'][0])
    '''
    e = a.time_information()
    print(e)



