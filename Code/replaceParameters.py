#招募合同替换参数
import logging.config
import os



ins = {
    '相关服务机构':{
        '三、出具法律意见书的律师事务所': {
            '名称':'{%laworgfullname%}',
            '注册地址': '{%lawregisteraddress%}',
            '办公地址': '{%lawregisteraddress%}',
            '电话': '{%lawtel%}',
            '传真': '{%lawfax%}',
            '负责人': '{%lawcharger%}',
            },
        '四、审计基金财产的会计师事务所': {
            '名称':'{%accountingorgfullname%}',
            '住所': '{%accountingregisteraddress%}',
            '办公地址': '{%accountingregisteraddress%}',
            '执行事务合伙人': '{%accountingcharger%}',
            '联系电话': '{%accountingtel%}',
            '传真': '{%accountingfax%}',
            }

    }
}

one_map = {
    '基金管理人': [
        ['一、基金管理人概况','二、主要成员情况','{%managercompanyinfo%}'],
        ['二、主要成员情况','三、基金管理人的职责','{%managerpersoninfo%}'],
    ],
    '基金托管人': [
        ['一、基本情况', '二、基金托管部门及主要人员情况', '{%trustcompanyinfo%}'],
        ['二、基金托管部门及主要人员情况', '三、证券投资基金托管情况', '{%trustpersoninfo%}'],
        ['三、证券投资基金托管情况', '四、托管业务的内部控制制度', '{%trusttransinfo%}'],
        ['四、托管业务的内部控制制度', '五、托管人对管理人运作基金进行监督的方法和程序', '{%trustcontrollawinfo%}'],
        ]
}

empty_map = {
    '序言': '{%eeeeee%}'

}

def logger_():
    logpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logger.conf')
    print(logpath)
    logging.config.fileConfig(logpath)
    logger = logging.getLogger('cse')
    return logger

def replace1(text, kv):
    '''
    以冒号分割的字符串进行替换
    :param text:字符串
    :param kv:替换规则
    :return:替换后的字符串
    '''
    for k,v in kv.items():
        if text.startswith(k+"：") or text.startswith(k+":"):
            #字符串满足：kv的键+“：”，就替换
            text = k + '：' + v
    return text

def parameter1(header, sub, text):
    '''
    过滤ins参数，进行替换
    :param header: 文章的章节名字
    :param sub:小标题名字
    :param text:一段内容
    :return:小标题，替换后的内容
    '''
    if text.strip() in ins[header]:
        sub = text.strip()
    elif sub:
        text = replace1(text, ins[header][sub])
    return sub, text


def data(helfData):
    sub = ''
    for c in helfData:
        h = c.split('部分')[-1]
        #print(h)
        if h in ins:
            #根据冒号分割替换
            for i in range(len(helfData[c])):
                sub, helfData[c][i] = parameter1(h, sub, helfData[c][i])
                #print(helfData[c][i])
        elif h in one_map:
            #根据前后标题替换
            for i in range(len(one_map[h])):
                start = one_map[h][i][0]
                end = one_map[h][i][1]
                for j in range(len(helfData[c])):
                    if start in helfData[c][j]:
                        s = j+1
                    if end in helfData[c][j]:
                        d = j
                helfData[c] = helfData[c][0:s] + [one_map[h][i][2]] +helfData[c][d:]
        else:
            pass
    return helfData

def printData(dat):
    logger = logger_()
    logger.info('*' * 20+ '这是被替换的数据'+ '*' * 20)
    for d in dat:
        h = d.split('部分')[-1]
        if h in ins or h in one_map:
            for p in dat[d]:
                logger.info(p)
    logger.info('*' * 20 + '这是被替换的数据' + '*' * 20)


if __name__ == "__main__":
    file_name = '博时中证500指数增强型证券投资基金招募说明书.docx'
    from Code.word_separate import WordFileClean
    helfData = WordFileClean(file_name).half_dict()
    dat = data(helfData)
    printData(dat)
