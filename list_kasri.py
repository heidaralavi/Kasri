import pandas as pd
import numpy as np
import re

def find_dl(x):
    regexp = re.compile(r'([ ]|^|[,]|[.]|[-]|[:])[D][L,l]')
    if regexp.search(str(x)):
        #print('ok')
        return 'yes'
    else:
        return 'no'



anbar_41 = pd.read_excel('41.xlsx', dtype='string').apply(lambda x: x.str.strip()).fillna('')
anbar_42 = pd.read_excel('42.xlsx', dtype='string').apply(lambda x: x.str.strip()).fillna('')
anbar_43 = pd.read_excel('43.xlsx', dtype='string').apply(lambda x: x.str.strip()).fillna('')
anbar_44 = pd.read_excel('44.xlsx', dtype='string').apply(lambda x: x.str.strip()).fillna('')




anbar_karfarma = pd.concat([anbar_41, anbar_42, anbar_43, anbar_44], axis=0).reset_index()

print(anbar_karfarma.size)
#print (anbar_karfarma.iloc[0:33427,6:7].apply(find_dl,axis=1))
anbar_karfarma['yes-no'] = anbar_karfarma.apply(find_dl,axis=1)
    
#anbar_karfarma.to_excel('1.xlsx')