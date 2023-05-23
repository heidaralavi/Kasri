import pandas as pd

def make_float(x):
    return float(x)


anbar_41 = pd.read_excel('41.xlsx',usecols=[0, 9], names=['Code_kala', 'Mandeh_karfarma'], dtype='string').apply(lambda x: x.str.strip()).fillna('0')
anbar_42 = pd.read_excel('42.xlsx',usecols=[0, 9], names=['Code_kala', 'Mandeh_karfarma'], dtype='string').apply(lambda x: x.str.strip()).fillna('0')
anbar_43 = pd.read_excel('43.xlsx',usecols=[0, 9], names=['Code_kala', 'Mandeh_karfarma'], dtype='string').apply(lambda x: x.str.strip()).fillna('0')
anbar_44 = pd.read_excel('44.xlsx',usecols=[0, 9], names=['Code_kala', 'Mandeh_karfarma'], dtype='string').apply(lambda x: x.str.strip()).fillna('0')
pamidco = pd.read_excel('پامیدکو.xlsx',usecols=[1, 4], names=['Code_kala', 'Mandeh_pamidco'], dtype='string').apply(lambda x: x.str.strip()).fillna('0')
parseh = pd.read_excel('گزارش قطعات به ازای دستورکارها.xlsx',usecols=[2,3,18,24], names=['AppTag','dastoorkar','Code_kala', 'tedad'], dtype='string').apply(lambda x: x.str.strip()).fillna('0')

anbar_karfarma = pd.concat([anbar_41, anbar_42, anbar_43, anbar_44], axis=0)

anbarha = anbar_karfarma.merge(pamidco,'left',on=['Code_kala']).fillna('0')
anbarha['Mandeh_karfarma']=anbarha['Mandeh_karfarma'].apply(make_float)
anbarha['Mandeh_pamidco']=anbarha['Mandeh_pamidco'].apply(make_float)
anbarha['mojudi']=anbarha['Mandeh_karfarma']+anbarha['Mandeh_pamidco']
anbarha['masraf_shode']=0.0

for item in parseh['Code_kala']:
    t = parseh.loc[parseh['Code_kala']==item]
    #anbarha.loc[anbarha['Code_kala']==item , 'masraf_shode']=454545
    print(item , t['tedad'].values)
    



#final_report = parseh.merge(anbarha,'left',on=['Code_kala']).fillna('0')
#anbarha.to_excel('1.xlsx',index = False)