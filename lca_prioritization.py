import pandas as pd
import numpy as np
import numpy_financial as npf
import math
from openpyxl import load_workbook
import shutil
import time

t1=time.time()

input= ('lca_prioritization_input.xlsx')
output=('lca_prioritization_output.xlsx')
b=pd.read_excel(input,sheet_name='input_a')
c=pd.read_excel(input,sheet_name='input_b')
b.set_index('strategy',inplace=True)
c.set_index('code',inplace=True)

t2=time.time()
################################################################################################
#B6 - Environmental & Economic
################################################################################################
b['el']=(b['ed_h']*b['el_h'])*3.6
b['ec_h']=((b['ed_h']*3.6+b['el']-b['egh_h']*3.6)/b['eef_h'])-(b['ege_h']*3.6)
b['ec_w']=((b['ed_w']*3.6        -b['egh_w']*3.6)/b['eef_w'])-(b['ege_w']*3.6)
b['ec_t']=b['ec_h']+b['ec_w']
for n in b.index:                                                          ####ENVIRONMENTAL####
    b.loc[n,'gw_b6_h']=b.loc[n,'ec_h']*c.loc[b.loc[n,'process_h'],'gw_b6']  #GPW  heating
    b.loc[n,'gw_b6_w']=b.loc[n,'ec_w']*c.loc[b.loc[n,'process_w'],'gw_b6']  #GPW  dhw
    b.loc[n,'gw_b6']= b.loc[n,'gw_b6_h']+ b.loc[n,'gw_b6_w']                #GPW  total B6
    b.loc[n,'pe_b6_h']=b.loc[n,'ec_h']*c.loc[b.loc[n,'process_h'],'pe_b6']  #NRPE heating
    b.loc[n,'pe_b6_w']=b.loc[n,'ec_w']*c.loc[b.loc[n,'process_w'],'pe_b6']  #NRPE dhw
    b.loc[n,'pe_b6']= b.loc[n,'pe_b6_h']+ b.loc[n,'pe_b6_w']                #NRPE total B6

    fc_h_base=fc_h_sum=b.loc[n,'ec_h']*c.loc[b.loc[n,'process_h'],'fc_b6'] ######ECONOMIC#######
    fc_w_base=fc_w_sum=b.loc[n,'ec_w']*c.loc[b.loc[n,'process_w'],'fc_b6'] 
    for y in range(1,int(b.loc[n,'rslb'])):
        fc_h_sum=fc_h_sum+(fc_h_base*(1+(c.loc[b.loc[n,'process_h'],'fc_in_b6']))**y)
        fc_w_sum=fc_w_sum+(fc_w_base*(1+(c.loc[b.loc[n,'process_w'],'fc_in_b6']))**y)
    b.loc[n,'fc_b6_h']=fc_h_sum/b.loc[n,'rslb']                                 #Full Cost heating
    b.loc[n,'fc_b6_w']=fc_w_sum/b.loc[n,'rslb']                                  #Full Cost dhw
    b.loc[n,'fc_b6']=b.loc[n,'fc_b6_h']+b.loc[n,'fc_b6_w']                  #Full Cost total B6

################################################################################################
#A1-3 - A5 - B2 - B4_A1-3 - B4_A5 - Environmental & Economic
################################################################################################
#STAGE A1-3 ENV & EC
for n in b.index:
    for i in ('gw_a13','pe_a13','fc_a13'):
        b.loc[n,i]=((
        c.loc[b.loc[n,'mat_1'],i]*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*((b.loc[n,'th']))+       
        c.loc[b.loc[n,'mat_2'],i]*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']+
        c.loc[b.loc[n,'mat_3'],i]*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med'])/(b.loc[n,'s']*b.loc[n,'rslb'])) 
#STAGE A5 EC
for n in b.index:
    b.loc[n,'fc_a5']=((
    c.loc[b.loc[n,'mat_1'],'fc_a5']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']+       
    c.loc[b.loc[n,'mat_2'],'fc_a5']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']+
    c.loc[b.loc[n,'mat_3'],'fc_a5']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med'])/(b.loc[n,'s']*b.loc[n,'rslb'])) 

#STAGE B2 EC
for n in b.index:
    fc_b2_base=fc_b2_sum=(
    c.loc[b.loc[n,'mat_1'],'fc_b2']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']+
    c.loc[b.loc[n,'mat_2'],'fc_b2']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_1_med']+
    c.loc[b.loc[n,'mat_3'],'fc_b2']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_1_med'])
    for y in range(1,int(b.loc[n,'rslb'])):
        fc_b2_sum=fc_b2_sum+(fc_b2_base*(1+b.loc[n,'inf'])**y)
    b.loc[n,'fc_b2']=fc_b2_sum/(b.loc[n,'s']*b.loc[n,'rslb'])

#STAGE B4_A1-3 ENV & EC
for n in c.index:
    c.loc[n,'rp']=(math.ceil(b.loc['base','rslb']/c.loc[n,'rslm']))-1
for n in b.index:
    b.loc[n,'gw_b4_a13']=((
    c.loc[b.loc[n,'mat_1'],'gw_a13']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*c.loc[b.loc[n,'mat_1'],'rp']*((b.loc[n,'th']))+
    c.loc[b.loc[n,'mat_2'],'gw_a13']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']*c.loc[b.loc[n,'mat_2'],'rp']+
    c.loc[b.loc[n,'mat_3'],'gw_a13']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med']*c.loc[b.loc[n,'mat_3'],'rp'])/(b.loc[n,'s']*b.loc[n,'rslb']))
    b.loc[n,'pe_b4_a13']=((
    c.loc[b.loc[n,'mat_1'],'pe_a13']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*c.loc[b.loc[n,'mat_1'],'rp']*((b.loc[n,'th']))+
    c.loc[b.loc[n,'mat_2'],'pe_a13']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']*c.loc[b.loc[n,'mat_2'],'rp']+
    c.loc[b.loc[n,'mat_3'],'pe_a13']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med']*c.loc[b.loc[n,'mat_3'],'rp'])/(b.loc[n,'s']*b.loc[n,'rslb']))
    b.loc[n,'fc_b4_a13']=((
    c.loc[b.loc[n,'mat_1'],'fc_a13']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*c.loc[b.loc[n,'mat_1'],'rp']*((b.loc[n,'th']))*((1+b.loc[n,'inf'])**c.loc[b.loc[n,'mat_1'],'rslm'])+
    c.loc[b.loc[n,'mat_2'],'fc_a13']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']*c.loc[b.loc[n,'mat_2'],'rp']*((1+b.loc[n,'inf'])**c.loc[b.loc[n,'mat_2'],'rslm'])+
    c.loc[b.loc[n,'mat_3'],'fc_a13']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med']*c.loc[b.loc[n,'mat_3'],'rp']*((1+b.loc[n,'inf'])**c.loc[b.loc[n,'mat_3'],'rslm']))
    /(b.loc[n,'s']*b.loc[n,'rslb']))

#STAGE B4_A5 EC
for n in b.index:
    b.loc[n,'fc_b4_a5']=((
    c.loc[b.loc[n,'mat_1'],'fc_a5']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*c.loc[b.loc[n,'mat_1'],'rp']*((1+b.loc[n,'inf'])**c.loc[b.loc[n,'mat_1'],'rslm'])+
    c.loc[b.loc[n,'mat_2'],'fc_a5']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']*c.loc[b.loc[n,'mat_2'],'rp']*((1+b.loc[n,'inf'])**c.loc[b.loc[n,'mat_1'],'rslm'])+
    c.loc[b.loc[n,'mat_3'],'fc_a5']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med']*c.loc[b.loc[n,'mat_3'],'rp']*((1+b.loc[n,'inf'])**c.loc[b.loc[n,'mat_1'],'rslm']))
    /(b.loc[n,'s']*b.loc[n,'rslb']))

################################################################################################
#PRIORIZATION

#enviornmental
b['gw_t']=b['gw_a13']+b['gw_b4_a13']+b['gw_b6']
b['pe_t']=b['pe_a13']+b['pe_b4_a13']+b['pe_b6']
b['ner']=(b.loc['base','pe_b6']-b['pe_b6'])/(b['pe_t']-b['pe_b6'])
b['emp']=((b['gw_a13']+b['gw_b4_a13'])*b['rslb'])/(
    (b.loc['base','gw_b6']-b['gw_b6'])-(b['gw_a13']+b['gw_b4_a13']))

#economic
b['fc_t']=b['fc_b6']+b['fc_a13']+b['fc_a5']+b['fc_b2']+b['fc_b4_a13']+b['fc_b4_a5']
for n in b.index:
    irr_inv=(b.loc[n,'fc_a13']+b.loc[n,'fc_a5']+b.loc[n,'fc_b4_a13']+b.loc[n,'fc_b4_a5'])*b.loc[n,'rslb']
    irr_ann=[None]*int(b.loc[n,'rslb'])
    van=[None]*200
    van[0]=0
    irr_ann[0]=(-1)*irr_inv
    for x in range(1,200):
        irr_save=((b.loc['base','ec_h']-b.loc[n,'ec_h'])*(c.loc[b.loc[n,'process_h'],'fc_b6']*(1+c.loc[b.loc[n,'process_h'],'fc_in_b6'])**x)+
                  (b.loc['base','ec_w']-b.loc[n,'ec_w'])*(c.loc[b.loc[n,'process_w'],'fc_b6']*(1+c.loc[b.loc[n,'process_w'],'fc_in_b6'])**x))
        irr_b2=((  c.loc[b.loc[n,'mat_1'],'fc_b2']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']+
                   c.loc[b.loc[n,'mat_2'],'fc_b2']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']+
                   c.loc[b.loc[n,'mat_3'],'fc_b2']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med'])/b.loc[n,'s'])*((1+b.loc[n,'inf'])**x)
        van[x]=(van[x-1])+ irr_save - irr_b2
    for x in range(1,int(b.loc[n,'rslb'])):
        irr_save=((b.loc['base','ec_h']-b.loc[n,'ec_h'])*(c.loc[b.loc[n,'process_h'],'fc_b6']*(1+c.loc[b.loc[n,'process_h'],'fc_in_b6'])**x)+
                  (b.loc['base','ec_w']-b.loc[n,'ec_w'])*(c.loc[b.loc[n,'process_w'],'fc_b6']*(1+c.loc[b.loc[n,'process_w'],'fc_in_b6'])**x))
        irr_b2=((  c.loc[b.loc[n,'mat_1'],'fc_b2']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']+
                   c.loc[b.loc[n,'mat_2'],'fc_b2']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']+
                   c.loc[b.loc[n,'mat_3'],'fc_b2']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med'])/(b.loc[n,'s']*b.loc[n,'rslb']))*((1+b.loc[n,'inf'])**x)
        irr_ann[x]=irr_save-irr_b2
    b.loc[n,'irr']=npf.irr(irr_ann)
#    z=[ele for ele in van if ele > 0]
#    lcp=van.index(min(z))

b['gw_b6_red']=b.loc['base','gw_b6']-b['gw_b6']                             #GWP Operational reduction
b['gw_int']=b['gw_a13']+b['gw_b4_a13']                                      #GWP Internal 
b['pe_b6_red']=b.loc['base','pe_b6']-b['pe_b6']                             #NRPE Operational reduction
b['pe_int']=b['pe_a13']+b['pe_b4_a13']                                      #NRPE Internal 
b['fc_b6_red']=b.loc['base','fc_b6']-b['fc_b6']                             #FC  Operational reduction
b['fc_int']=b['fc_a13']+b['fc_a5']+b['fc_b2']+b['fc_b4_a13']+b['fc_b4_a5']  #FC  Internal

################################################################################################
#RESULTS

#reindex
lca=b.filter(['ec_h','ec_w','ec_t',
'gw_b6_h','gw_b6_w','gw_b6','gw_a13','gw_b4_a13','gw_t',
'pe_b6_h','pe_b6_w','pe_b6','pe_a13','pe_b4_a13','pe_t',
'fc_b6_h','fc_b6_w','fc_b6','fc_a13','fc_a5','fc_b2','fc_b4_a13','fc_b4_a5','fc_t',
'ner','irr'])

#formatting
for n in ('ec_h','ec_w','ec_t','fc_b6_h','fc_b6_w','fc_b6','fc_a13','fc_a5','fc_b2','fc_b4_a13','fc_b4_a5','fc_a5','fc_t','ner'):
    lca[n]=round(lca[n],2) # round to 2 decimals
for n in ('gw_b6_h','gw_b6_w','gw_b6','pe_b6_h','pe_b6_w','pe_b6','gw_a13','pe_a13','gw_b4_a13','pe_b4_a13','gw_t','pe_t'):
    lca[n] = lca[n].map('{:.2E}'.format) # cientific format with 2 decimals
lca['irr']= lca['irr'].map('{:,.2%}'.format) # percentage with 2 decimals

t3=time.time()

### (Export Excel) ###
shutil.copy(input,output)
book = load_workbook(output)
writer = pd.ExcelWriter(output,engine='openpyxl')
writer.book=book
lca.to_excel(writer, sheet_name='output')
writer.save()
print('Results have been saved correcty')

t4=time.time()

################################################################################################
#EXECUTION TIME
print('EXECUTION TIMES:')
print('Input data import execution time',(t2-t1)*1000,'ms')
print('Env-Ec assessment execution time: ',(t3-t2)*1000,'ms')
print('Output data export execution time: ',(t4-t3)*1000,'ms')
