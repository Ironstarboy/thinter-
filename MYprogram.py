import pandas as pd
import xlrd as xd
import re
import jieba    
import datetime
import time
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.pipeline import make_pipeline

#利用jieba库定义分词函数
def chinese_word_cut(sentences):
    return " ".join(jieba.cut(sentences))
#定义停用词识别函数
def get_custom_stopwords(stop_words_file):
    with open(stop_words_file,encoding='utf-8') as f:
        stopwords = f.read()
    stopwords_list = stopwords.split('\n')
    custom_stopwords_list = [i for i in stopwords_list]
    return custom_stopwords_list
#定义找到第一个数字的函数,找到则输出index，找不到则输出-1
def index_first_num(string):
    k = len(string)
    if k == 0:
        return -1
    else:
        for i in range(k):
            if string[i].isdigit():
                return i
        return -1
#判断是否为工作日，是的话输出-1；否则周六输出2，周日输出1
def workday(string):
    tt=time.strptime(string,'%Y年%m月%d日')
    if tt.tm_wday < 5:
        return -1
    else:
        return 7 - tt.tm_wday

#设置分语段的模型
#导入训练数据


def run(file_name):

    df = pd.read_csv('data_label.csv')


    X = df[['sentences']]
    y = df[['attribute']]
    X['cutted_sentences'] = X.sentences.apply(chinese_word_cut)
    X['cutted_sentences_zero'] = X.sentences.apply(chinese_word_cut)
    #用零代表每个位置上的数字
    for i in range(len(X['cutted_sentences_zero'])):
        for j in range(len(X.cutted_sentences_zero[i])):
            if (X.cutted_sentences_zero[i][j].isdigit()):
                X.cutted_sentences_zero[i] = X.cutted_sentences_zero[i].replace(X.cutted_sentences_zero[i][j],'0')
    #载入停用词表
    stop_words_file = "停用词表.txt"
    stopwords = get_custom_stopwords(stop_words_file)
    #在特征中除去停用词
    vect = CountVectorizer(stop_words=frozenset(stopwords))
    nb = MultinomialNB()
    pipe = make_pipeline(vect, nb)
    pipe.fit(X.cutted_sentences_zero, y)

    #导入需要测试的数据
    print('读取中...')
    data = xd.open_workbook (file_name) #打开excel表所在路径
    print('读取成功！')
    sheet = data.sheet_by_name('test')  #读取数据，以excel表名来打开
    list_test = []
    #将表中数据按行逐步添加到列表中，最后转换为list结构
    for r in range(sheet.nrows):
        data = []
        for c in range(sheet.ncols):
            data.append(sheet.cell_value(r,c))
        list_test.append(list(data))
    str_list = []
    for i in range(len(list_test)):
        str_list.append(re.split('[ ,，。\n]',str(list_test[i][0])))
        while '' in str_list[i]:
            str_list[i].remove('')

    #利用模型预测test
    output = pd.DataFrame(columns=('structure','underlying','principal','coupon','sum_guarantee','guarantee rate','期限','锁定期','初始日期','终止日期','strike_1','strike_2','strike_3'))
    #对输入excel中的每一行进行处理
    for k in range(len(str_list)):
        #对输入的一行预处理：分词，将数字转化为零
        str_series = pd.Series(str_list[k])
        cutted_str_zero = str_series.apply(chinese_word_cut)
        for i in range(len(cutted_str_zero)):
            for j in range(len(cutted_str_zero[i])):
                if (cutted_str_zero[i][j].isdigit()):
                    cutted_str_zero[i] = cutted_str_zero[i].replace(cutted_str_zero[i][j],'0')
        #输出每一个语段的预测结果
        predict = pipe.predict(cutted_str_zero)
        str_list_out = ['','','','','','','']
        #对于每种一行一句的语段特殊处理，除去冒号前的内容
        if '\n' in list_test[k][0]:
            for j in range(len(str_list[k])):
                if '：' in str_list[k][j]:
                    index = str_list[k][j].index('：')
                    str_list[k][j] = str_list[k][j][index + 1 :]
                elif ':' in str_list[k][j]:
                    index = str_list[k][j].index(':')
                    str_list[k][j] = str_list[k][j][index + 1 :]
        #输出预测的语段结果
        m = 0
        for i in predict:
            str_list_out[i] = str_list_out[i].replace(str_list_out[i],str_list_out[i] + str_list[k][m])
            m = m + 1

        #本金部分缩约
        for p in range(len(str_list_out[4])):
            if str_list_out[4][p].isdigit():
                continue
            elif str_list_out[4][p] in ['万','w','百万','W']:
                continue
            else:
                str_list_out[4] = str_list_out[4].replace(str_list_out[4][p],'F')
        str_list_out[4] = str_list_out[4].replace('F','')
        for i in ['W','w','万']:
            str_list_out[4] = str_list_out[4].replace(i,'0000')

        #票息缩约
        for p in range(len(str_list_out[6])):
            if str_list_out[6][p].isdigit():
                continue
            elif str_list_out[6][p] in ['%','%','.']:
                continue
            else:
                str_list_out[6] = str_list_out[6].replace(str_list_out[6][p],'F')
        str_list_out[6] = str_list_out[6].replace('F','')

        #提取保证金率和总保证金
        str_list_out_guarantee = ['','']
        index_1 = index_first_num(str_list_out[5])
        if index_1 != -1:
            index_2 = index_1
            while(index_2 > -1):
                if index_2 > len(str_list_out[5]) - 1:
                    break
                elif str_list_out[5][index_2].isdigit():
                    index_2 = index_2 + 1
                elif str_list_out[5][index_2] in ['%','%','.','w','万','W']:
                    index_2 = index_2 +1
                else:
                    break
            if str_list_out[5][index_2 - 1] in ['%','%']:
                str_list_out_guarantee[1] = str_list_out[5][index_1:index_2]
            elif str_list_out[index_2 - 1 ] in ['w','W','万']:
                str_list_out_guarantee[0] = str_list_out[5][index_1:index_2]
            else :
                str_list_out_guarantee[0] = str_list_out[5][index_1:index_2]
            index_3 = index_first_num(str_list_out[5][index_2:])
            if index_3 != -1:
                index_3 = index_3 + index_2
                index_4 = index_3
                while(index_4 > -1):
                    if index_4 > len(str_list_out[5]) - 1:
                        break
                    elif str_list_out[5][index_4].isdigit():
                        index_4 = index_4 + 1
                    elif str_list_out[5][index_4] in ['%','%','.','w','万','W']:
                        index_4 = index_4 +1
                    else:
                        break
                if str_list_out[5][index_4 - 1] in ['%','%']:
                    str_list_out_guarantee[1] = str_list_out[5][index_3:index_4]
                elif str_list_out[5][index_4 - 1 ] in ['w','W','万']:
                    str_list_out_guarantee[0] = str_list_out[5][index_3:index_4]
                else :
                    str_list_out_guarantee[0] = str_list_out[5][index_3:index_4]
        for i in ['W','w','万']:
            str_list_out_guarantee[0] = str_list_out_guarantee[0].replace(i,'0000')
        for i in str_list_out_guarantee:
            str_list_out.append(i)

        #期限缩约
        str_list_out_term = ['','']
        index_1 = index_first_num(str_list_out[2])
        term_1 = ''
        term_2 = ''
        index_2 = index_1
        if index_2 != -1:
            while (str_list_out[2][index_2].isdigit()):
                index_2 = index_2 + 1;
            term_1 = term_1.replace('',str_list_out[2][index_1:index_2])
            index_3 = index_first_num(str_list_out[2][index_2:])
            if index_3 != -1:
                index_3 = index_3 + index_2
                index_4 = index_3
                while(index_4 > -1):
                    if index_4 > len(str_list_out[2]) - 1:
                        break
                    elif str_list_out[2][index_4].isdigit():
                        index_4 = index_4 + 1
                    else:
                        break
                term_2 = term_2.replace('',str_list_out[2][index_3:index_4])
        if term_2 != '':
            str_list_out_term[0] = str_list_out_term[0].replace('',term_1 if int(term_1)>int(term_2) else term_2)
            str_list_out_term[1] = str_list_out_term[1].replace('',term_1 if int(term_1)<=int(term_2) else term_2)
        else:
            str_list_out_term[0] = str_list_out_term[0].replace('',term_1)
        for i in str_list_out_term:
            str_list_out.append(i)

        #起始日和终止日
        now = datetime.datetime.now()
        start = str(now.year) + '年' + str(now.month) + '月' + str(now.day) + '日'
        str_list_out.append(start)
        termination_month = now.month + int(str_list_out[9])
        if termination_month == 12:
            termination = str(now.year) + '年' + str(12) + '月' + str(now.day) + '日'
        else:
            temp = int(termination_month / 12)
            termination_month = termination_month % 12
            termination_year = now.year + temp
            termination = str(termination_year) + '年' + str(termination_month) + '月' + str(now.day) + '日'
        temp =  workday(termination)
        if temp == -1:
            str_list_out.append(termination)
        else:
            termination = datetime.datetime.strptime(termination,'%Y年%m月%d日') + datetime.timedelta(days=workday(termination))
            termination = str(termination.year) + '年' + str(termination.month) + '月' + str(termination.day) + '日'
            str_list_out.append(termination)
        output.loc[k] = str_list_out

        #获得strike的中所有数字
        num_strike = []
        strike = str_list_out[3]
        while(index_first_num(strike) != -1):
            index = index_first_num(strike)
            for i in range(len(strike[index:])):
                if strike[i + index].isdigit():
                    continue
                elif strike[i + index] in ['.']:
                    continue
                else:
                    break
            if (i+1) == len(strike[index:]) and strike[-1].isdigit():
                num_strike.append(strike[index:])
                break
            num_strike.append(strike[index:i + index])
            strike = strike[i + index:]
        num_list = [num_strike[0],'','']
        for i in range(len(num_strike)):
            if float(num_strike[i]) < 50 or float(num_strike[i]) > 150:
                num_strike[i] = ''
        num_strike = sorted(list(set(num_strike)))
        index = 0
        while(num_strike[index] == ''):
            index = index + 1
            if index > len(num_strike) - 1:
                break
        num_strike = num_strike[index:]
        for i in range(len(num_strike)):
            num_strike[i] = float(num_strike[i])
        num_strike = sorted(num_strike)
        for i in range(3 - len(num_strike)):
            num_strike.append('')
        for i in num_strike:
            str_list_out.append(i)

        str_list_out[5:] = str_list_out[6:]
        str_list_out[2:] = str_list_out[4:]

        output.loc[k] = str_list_out
    #输出到excel
    output.to_csv('output_test.csv',encoding="utf_8_sig")


        

        
        
        
        
        
    