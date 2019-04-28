
# 读取已有的word文档
# 参考1：https://www.jianshu.com/p/94ac13f6633e
# 参考2: https://python-docx.readthedocs.io/en/latest/api/text.html#paragraph-objects
# 参考3: https://www.jianshu.com/p/867fe84d8440
import sys
import re

from docx import Document

def main():
    
    print ('译者名字' + ' 翻译项目' + ' 文档总字数' + ' 汉字总字数' + ' 预留参数位' + ' 时间' + '\n')
    f = open("1.txt")
    lines = f.readlines()
    for line in lines:
        document_path = line.strip()
        # document_path = str(sys.argv[1])
        # 创建文档对象
        document = Document(document_path)
        # 将路径按照斜杠拆分，得到文件名
        document_name = line.strip().split('/')
        document_name_count = len(document_name)
        # 将文件名按照横杠拆分，以便列表输出
        document_item = document_name[document_name_count-1].split('-')

        # 读取文档中所有的段落列表
        ps = document.paragraphs
        
        # 每个段落都有text 属性
        # 这里返回一个List值
        ps_detail = [(x.text) for x in ps]
        
        # 总共有多少段落
        ps_detail_count = len(ps_detail)

        # 用正则来选出汉字
        hanzi_regex = re.compile(r'[\u4E00-\u9FA5]')

        word_count = 0
        hanzi_count = 0

        for i in range(ps_detail_count):
            word_count += len(ps_detail[i])
            hanzi_count += len(hanzi_regex.findall(ps_detail[i]))
        
        document_item_time = document_item[3].split('.')
        print (document_item[0] + ' ' + document_item[1] + ' ' + str(word_count) + ' ' + str(hanzi_count) + ' ' + document_item[2] + ' ' + document_item_time[0])

if __name__ == '__main__':
    main()