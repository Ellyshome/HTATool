#=====用药统计======
def getlin(file):
    xyf={}
    flag = False
    with open(file, 'r', encoding='utf-8') as f:
        for line in f.readlines():
            if len(line)<2: continue # 行太短则无效
            if line[:3]=='西药费' or line[:4]=='中成药费':flag = True
            elif line[0] != '\t':flag = False
            idx = line.index('\t')  # 找到第一个空格的位置
            if flag == True:
                tag = line[idx+1:].split('\t')
                if tag[0] in xyf: xyf[tag[0]]=xyf[tag[0]] + float(tag[1])  #若有则增加
                else: xyf[tag[0]] = float(tag[1])  #若无则新建
    return(xyf)
import sys
while 1:
    filename = sys.argv[1] or input('拖入文件或写入文件名>>>')
    sys.argv[1] = ''
    lista = getlin(filename)
    for idx,li in enumerate(lista,start=1):print(f'{idx}\t{lista[li]} \t: {li}')