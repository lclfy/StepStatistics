# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
# https://www.zhihu.com/question/303305223

import easyocr
import ssl
import os
import glob
ssl._create_default_https_context = ssl._create_unverified_context
project_path = os.path.join(
        os.path.dirname(__file__)
    )
WSI_MASK_PATH = project_path#存放图片的文件夹路径
paths = glob.glob(os.path.join(WSI_MASK_PATH, '*.jpeg'))
reader = easyocr.Reader(['ch_sim','en'])
texts = ''
for j in range(len(paths)):
    imageText = reader.readtext(paths[j])
    for i in range(len(imageText)):
        texts = texts + imageText[i][1] + '\n'
f = open('log.txt','w')
print(texts, file=f)
print(texts)




# See PyCharm help at https://www.jetbrains.com/help/pycharm/
