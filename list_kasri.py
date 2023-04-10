import pandas as pd
import re

def find_dl(x):
    regexp = re.compile(r'[Dd][Ll][ ]*\d+')
    if regexp.search(str(x)):
        #print('ok')
        return 'ساختنی'
    else:
        return 'خریدنی'

def join_column(x):
    return ','.join(x)

def sum_column(x):
    return x.astype(float).sum()

def kasri_cal(x):
    x=x.astype(float)
    
    return x['موجودی_کارفرما']+x['موجودی پامیدکو']-x['تجمیع_مورد_نیاز']

anbar_41 = pd.read_excel('41.xlsx', dtype='string').apply(lambda x: x.str.strip()).fillna('')
anbar_41.columns = anbar_41.columns.str.strip()
anbar_42 = pd.read_excel('42.xlsx', dtype='string').apply(lambda x: x.str.strip()).fillna('')
anbar_42.columns = anbar_42.columns.str.strip()
anbar_43 = pd.read_excel('43.xlsx', dtype='string').apply(lambda x: x.str.strip()).fillna('')
anbar_43.columns = anbar_43.columns.str.strip()
anbar_44 = pd.read_excel('44.xlsx', dtype='string').apply(lambda x: x.str.strip()).fillna('')
anbar_44.columns = anbar_44.columns.str.strip()
anbar_pamidco = pd.read_excel('پامیدکو.xlsx', dtype='string').apply(lambda x: x.str.strip()).fillna('')
anbar_pamidco.columns = anbar_pamidco.columns.str.strip()

# preparing anbar pamidco
anbar_pamidco.drop(anbar_pamidco.iloc[:,[0,2,3,5]],inplace=True,axis=1)
anbar_pamidco.rename(columns={'مانده مقدار كالا در انبار':'موجودی پامیدکو'},inplace=True)
#anbar_pamidco.to_excel('1.xlsx')

# Preparing parse
parse = pd.read_excel('گزارش قطعات به ازای دستورکارها.xlsx', dtype='string').apply(lambda x: x.str.strip()).fillna('')
parse =parse[parse['کد قطعه'] != ""]
filter_map= parse['کد قطعه'].str.contains('44000000\d+')
parse_nashenakhte =parse[filter_map]
parse_nashenakhte = parse_nashenakhte[['کد قطعه','شماره دستور کار','AppTag','تعداد','توضیحات']]
parse_nashenakhte.rename(columns={'کد قطعه':'كد  كالا','شماره دستور کار':'تجمیع_درخواست','AppTag':'APP_Tag','تعداد':'تجمیع_مورد_نیاز','توضیحات':'شرح کالا'},inplace=True)
parse = parse[~filter_map]
parse = parse.groupby('کد قطعه').agg(تجمیع_درخواست=pd.NamedAgg(column = 'شماره دستور کار', aggfunc = join_column),APP_Tag=pd.NamedAgg(column = 'AppTag', aggfunc = join_column),تجمیع_مورد_نیاز=pd.NamedAgg(column = 'تعداد', aggfunc = sum_column))

# preparing anbar karfarma
anbar_karfarma = pd.concat([anbar_41, anbar_42, anbar_43, anbar_44], axis=0)
anbar_karfarma['شرح کالا']=anbar_karfarma.iloc[:,[2,6,5]].apply(lambda x : '{} {} {}'.format(x[0],x[1],x[2]), axis=1)
anbar_karfarma['ساختنی-خریدنی'] = anbar_karfarma['شرح کالا'].apply(find_dl)
anbar_karfarma.drop(anbar_karfarma.iloc[:,[1,2,3,4,5,6,7,8]],inplace=True,axis=1)
anbar_karfarma.rename(columns={"مانده": "موجودی_کارفرما"},inplace=True)

# preparing kharid
kharid = pd.read_excel('خرید.xlsx', dtype='string').apply(lambda x: x.str.strip())
kharid.columns = kharid.columns.str.strip()
filter_map= kharid['تاريخ'].str.contains('(1399|14)\d+')
kharid = kharid[filter_map] 
filter_map= kharid['رسيد شده'] == '0'
kharid = kharid[filter_map]
kharid = kharid.groupby('كد قلم').agg(تجمیع_درخواست_خرید=pd.NamedAgg(column = 'درخواست', aggfunc = join_column),تجمیع_مقدار_تاییدشده=pd.NamedAgg(column = 'تاييد شده', aggfunc = sum_column))

# merge dataframes
final_report = parse.merge(anbar_karfarma,'left',left_on = ['کد قطعه'],right_on=['كد  كالا'])

final_report = final_report.iloc[:,[3,0,1,2,5,6,4]]

final_report = final_report.merge(anbar_pamidco,'left',left_on = ['كد  كالا'],right_on=['كد كالا']).fillna('')

final_report.drop(final_report.iloc[:,[7]],axis = 1,inplace = True)

final_report.iloc[:,[6,7]].replace('','0',inplace=True)

final_report['کسری'] =final_report.iloc[:,[3,6,7]].apply(kasri_cal,axis=1)
final_report=final_report[final_report['کسری']<0]
final_report = final_report.merge(kharid,'left',left_on = ['كد  كالا'],right_on=['كد قلم'])
final_report = pd.concat([final_report,parse_nashenakhte],axis=0,)


final_report.to_excel('1.xlsx',index = False)

