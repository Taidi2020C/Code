# 方法2：关键词比重法
# 算得F1-Score = 0.6971705168744503(未将主题*2)
# 将主题*2后：0.7063620285496864
# 因此需要探索最佳的主题和正文比例，使得准确率最高

import pandas as pd
import jieba
import jieba.analyse
import re

'''
提取出的关键词数分别为（因此把topK设为100）
98
43
26
41
28
30
54
23
55
83
34
45
74
59
64
'''


# 数据处理，提取关键词和权重返回tuple list
# jieba进行关键词提取和权重计算，仍会受到停用词和数字的影响
# 因此先打破句子为list，去除停用词后再组成'伪句子'进行关键词权重提取
def sentence_process(sentence):
    punc = u'0123456789.'
    nodigit_sentence = re.sub(r'[{}]'.format(punc), '', sentence)
    sentence_list = list(jieba.cut(nodigit_sentence))
    processed_str = ''
    stopwords_list = [
        line.strip() for line in open(r'D:\AA大学资料\文件 - 个人\大一下 数据挖掘\停用词表.txt',
                                      encoding='UTF-8').readlines()
    ]
    for word in sentence_list:
        if word not in stopwords_list:
            processed_str = processed_str + word
    sentence_key_list = jieba.analyse.extract_tags(processed_str, withWeight=True, topK=30)
    return sentence_key_list


# 获取意见对应各一级大类的特征值dic
def eigenvalues(key_dic, sentence_processed):
    key_rate_multiple_result = {}
    for key in key_dic.keys():
        key_rate_multiple = 0
        for i in range(len(key_dic[key])):
            key_name, key_rate = key_dic[key][i]
            for j in range(len(sentence_processed)):
                sentence_key_name, sentence_key_rate = sentence_processed[j]
                if sentence_key_name == key_name:
                    key_rate_multiple += sentence_key_rate * key_rate
        key_rate_multiple_result[key] = key_rate_multiple
    return key_rate_multiple_result


# 返回分类结果
def result(key_rate_multiple_result):
    key_rate_max = 0
    result_class = 'Wrong'
    for key in key_rate_multiple_result.keys():
        if key_rate_multiple_result[key] > key_rate_max:
            key_rate_max = key_rate_multiple_result[key]
            result_class = key
    return result_class


# 获取一级分类关键词库（关键词+权重tuple）
key_ex = pd.read_excel(r'D:\AA大学资料\文件 - 个人\大一下 数据挖掘\示例数据\附件1.xlsx',
                       usecols=['一级分类', '二级分类', '三级分类'])
key_str_dic = {}
key_str = ''
key_df = pd.DataFrame(index=['一级分类', '关键词', '权重'])
key_1 = key_ex['一级分类']
key_1_only = pd.DataFrame(key_1.drop_duplicates())  # 对一级分类去重并由数组转化为sentenceframe
key_1_only = key_1_only.reset_index(drop=True)  # 重置索引
for i in range(len(key_1_only)):
    for n in range(len(key_ex)):
        if key_ex.values[n][0] == key_1_only.values[i][0]:
            key_str = key_str + str(key_ex.values[n][1]) + str(
                key_ex.values[n][2])
    key_str_dic[str(key_1_only.values[i][0])] = key_str
    key_str = ''
key_dic = {}
for key in key_str_dic.keys():
    key_list = jieba.analyse.extract_tags(key_str_dic[key], withWeight=True, topK=100)
    key_dic[key] = key_list

advice_sentence_ex = pd.read_excel(
    r'D:\AA大学资料\文件 - 个人\大一下 数据挖掘\示例数据\清洗后.xlsx',
    usecols=['留言编号', '留言主题', '留言详情', '一级分类'])
advice_class_list = []
for i in range(len(advice_sentence_ex)):
    advice_processed = sentence_process(str(advice_sentence_ex.values[i][1])*2 + str(advice_sentence_ex.values[i][2]))
    advice_class_list.append(result(eigenvalues(key_dic, advice_processed)))
advice_sentence_ex['分类结果'] = advice_class_list
advice_sentence_ex.to_excel(r'D:\AA大学资料\文件 - 个人\大一下 数据挖掘\示例数据\result3.xlsx', sheet_name='1')
