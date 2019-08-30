# coding: utf-8

import os
import ahocorasick
import jieba
import numpy as np
import re
import string
from gensim.models import KeyedVectors
from gensim.models import word2vec
from docx import Document
from win32com import client as wc
from tqdm import tqdm
import datetime

class Word_embedding:
    """
    注意:对于专业领域的词汇---分词时将专业词汇添加到结巴默认词库中去
    功能：对任意doc,txt,docx格式的训练语料进行词向量训练
    附加功能：可实现在原有的词向量模型的基础上进行增量训练
    输入：将未处理doc,txt,docx格式的语料放入指定文件夹,以及输入词向量及模型存储位置
    输出：在指定输出文件夹位置输出词向量模型及词向量库
    """
    
    def __init__(self):
        
        # doc训练语料存放位置
        self.file_path = 'C:\\Users\\z50004593\\Desktop\\训练语料\\word_doc'
        
        # 停用词汇库txt文件存储位置
        self.stop_word_path = 'C:\\Users\\z50004593\\Desktop\\词库\\停用词库\\word_stop.txt'
        
        # 已处理的训练语料存放位置
        self.save_path = 'C:\\Users\\z50004593\\Desktop\\训练语料\\预处理后语料'
        
        # 训练后词向量以及模型的存储位置
        self.embedding_path = 'C:\\Users\\z50004593\\Desktop\\训练语料\\word_embedding_model'
        
        # word2vec参数设置
        self.parms={"sg":1, "hs":1,"size":300,"window":10, "min_count":5,
                    "iter":10, "min_count":7, "batch_words":5000, "seed":1}
               
    # 读取文件夹中的某种格式文件
    def filepath_read(self, ext):
        doc_file = [(self.file_path + '\\' + i) for i in os.listdir(self.file_path)if i.split('.')[-1] == ext]
        # print(doc_file)
        return doc_file
    
    
    # 文件位置p:list(doc文件位置)---对所有doc文件转换为docx文件
    def Savedocas(self):
        p = self.filepath_read('doc')
        print("doc 转 docx start:")
        word = wc.Dispatch('Word.Application')
        for path in tqdm(p):
            doc = word.Documents.Open(path)        
            # 目标路径下的文件
            filepath,fullflname = os.path.split(path)
            fname,ext = os.path.splitext(fullflname)
            # 转化后路径下的文件
            doc.SaveAs(filepath + '\\' + fname + '.docx', 12, False, "", True, "", False, False, False, False)      
            doc.Close()
        print("#已完成# doc文件转为docx格式")
    
    
    # 将docx文件中只读取正文内容(剔除图片表格)存储到txt文件
    def docx_to_txt(self):
        filepath = self.filepath_read('docx')
        print("docx 转 txt start:")
        for path in tqdm(filepath):
            doc = Document(path)
            filepath1, fullflname1 = os.path.split(path)
            fname1, ext1 = os.path.splitext(fullflname1)
            with open(filepath1 + '\\' + fname1 +'.txt', 'a') as f:
                for p in doc.paragraphs:
                    if p.text != '':
                        try:
                            f.write(str(p.text).strip())
                            f.write('\n')
                        except UnicodeEncodeError:
                            print('*'*10)
                            
            f.close()
        print("#已完成# docx文件转为txt格式")
       
    # 对txt文件进行预处理然后分词
    def pre_txt(self):
        path_list = self.filepath_read('txt')
        print("txt语料预处理 start:")
        stopwords = [w.strip() for w in open(self.stop_word_path, 'r',encoding='utf-8-sig') if w.strip().strip('\n')]
        for path in tqdm(path_list):
            with open(path,'r') as f:
                document = f.read()
                document_cut = [w.strip() for w in jieba.cut(document) if w.strip() not in stopwords ]
                result = ' '.join(document_cut)
            f.close()
            filepath3, fullflname3 = os.path.split(path)
            fname3, ext3 = os.path.splitext(fullflname3)
            with open(self.save_path + '\\' + fname3 + '_已清洗.txt', 'w', encoding='utf-8') as f2:
                f2.write(result)
            f2.close()
        print("#已完成# txt语料预处理")
        
        
    # 对多个已处理的txt文件词向量的训练
    def embedding_train(self):
        print("词向量训练 start")
        starttime = datetime.datetime.now()
        # 获取文件夹中所有文件
        sent = word2vec.PathLineSentences(self.save_path)
        # 具体参数在self.parms设置
        model = word2vec.Word2Vec(sentences=tqdm(sent), **self.parms)
        endtime = datetime.datetime.now()
        print ('秒：', (endtime - starttime).seconds)
        # 保存模型---载入word2vec.Word2Vec.load("\\name.model")
        model.save(self.embedding_path + '\\' + "word_embedding1.model")
        model.wv.save_word2vec_format(self.embedding_path + '\\' + "word_embedding1.txt", binary=0)
        # 载入bin、txt文件: gensim.models.KeyedVectors.load_word2vec_format('/ .txt/bin', binary=False) 
        model.wv.save_word2vec_format(self.embedding_path + '\\' + "word_embedding1.bin", binary=0)
        print("##词向量训练已完成## the word_embedding and model exists in " + self.embedding_path)
        return(model)
    
    
    # 词向量库的增量训练
    def update_model(self,text,model_path):
        """
        text:是已经预处理的语料
        model_path:已训练的模型
        """
        model = word2vec.Word2Vec.load(model_path)#保存的模型
        model.build_vocab(text, update = True)#更新词汇表
        model.train(text, total_examples = model.corpus_count, epochs = model.iter)
        ###model.save---保存
        model.save(self.embedding_path+ '\\' + "word_embedding1.model")
        model.wv.save_word2vec_format(self.embedding_path + '\\' + "word_embedding1.model", binary=0)
        model.wv.save_word2vec_format(self.embedding_path + '\\' + "word_embedding1.model", binary=0)
        print("the model/txt/bin updated in " + str(self.embedding_path))    
        
