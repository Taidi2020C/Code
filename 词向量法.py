# 方法1：词向量法
# 算得F1-Score = 0.7168017776026979（未将主题*2）
# 将主题*2后：0.7299544042933424
# 因此需要探索最佳的主题和正文比例，使得准确率最高

import pandas as pd
import jieba
import jieba.analyse
import re


# 返回意见在各一级分类下的文本向量
def set_word2vec(key_dic, sentence_cut_list):
    return_vec_dic = {}
    for classic_1 in key_dic.keys():
        classic_1_list = key_dic[classic_1]
        return_vec = [0] * len(classic_1_list)
        for word in sentence_cut_list:
            if word in classic_1_list:
                return_vec[classic_1_list.index(word)] += 1
        return_vec_dic[classic_1] = return_vec
    return return_vec_dic


# 数据处理，返回去除数字、停用词后的意见关键词List
def data_process(sentence):
    punc = u'0123456789.'
    nodigit_sentence = re.sub(r'[{}]'.format(punc), '', sentence)
    sentence_list = list(jieba.cut(nodigit_sentence))
    output_list = []
    stopwords_list = [
        line.strip() for line in open(r'D:\AA大学资料\文件 - 个人\大一下 数据挖掘\停用词表.txt',
                                      encoding='UTF-8').readlines()
    ]
    for word in sentence_list:
        if word not in stopwords_list:
            output_list.append(word)
    return output_list


# 返回向量模最大的分类
def result(vec_dic):
    max_len = 0
    max_class = 'Wrong'
    for key in vec_dic.keys():
        if sum(vec_dic[key]) > max_len:
            max_class = key
            max_len = sum(vec_dic[key])
    return max_class


# 获取一级分类关键词库
key_ex = pd.read_excel(r'D:\AA大学资料\文件 - 个人\大一下 数据挖掘\示例数据\附件1.xlsx',
                       usecols=['一级分类', '二级分类', '三级分类'])
key1 = key_ex['一级分类']
key1_only = pd.DataFrame(key1.drop_duplicates())
key1_only = key1_only.reset_index(drop=True)
key_list = []
key_dic = {}
for i in range(len(key1_only)):
    for n in range(len(key_ex)):
        if key_ex.values[n][0] == key1_only.values[i][0]:
            key_str_list = set(jieba.cut(key_ex.values[n][1])) | set(
                jieba.cut(key_ex.values[n][2]))
            key_list = list(set(key_list) | key_str_list)
    key_dic[str(key1_only.values[i][0])] = key_list
    key_list = []
advice_sentence_ex = pd.read_excel(
    r'D:\AA大学资料\文件 - 个人\大一下 数据挖掘\示例数据\清洗后.xlsx',
    usecols=['留言编号', '留言主题', '留言详情', '一级分类'])
advice_class_list = []
for i in range(len(advice_sentence_ex)):
    advice_processed = data_process(
        str(advice_sentence_ex.values[i][1])*2 +
        str(advice_sentence_ex.values[i][2]))
    advice_vec = set_word2vec(key_dic, advice_processed)
    advice_class_list.append(result(advice_vec))
advice_sentence_ex['分类结果'] = advice_class_list
advice_sentence_ex.to_excel(r'D:\AA大学资料\文件 - 个人\大一下 数据挖掘\示例数据\result1-1.xlsx', sheet_name='1')
