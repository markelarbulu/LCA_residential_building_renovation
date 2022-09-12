import pandas as pd
import numpy as np
import numpy_financial as npf
import math
from openpyxl import load_workbook
from sklearn.linear_model import LinearRegression
import sympy as sp
import matplotlib.pyplot as plt 
import shutil
import time

t1=time.time()

input=  ('lca_optimization_input.xlsx')
output= ('lca_optimization_output.xlsx')
b=pd.read_excel(input,sheet_name='input_a')
c=pd.read_excel(input,sheet_name='input_b')
b.set_index('strategy',inplace=True)
c.set_index('code',inplace=True)

t2=time.time()
################################################################################################
#B6 - Environmental & Economic
################################################################################################
b['el']=(b['ed_h']*b['el_h'])*3.6
b['ec_h']=((b['ed_h']*3.6+b['el'])/b['eef_h'])

for n in b.index:                                                          ####ENVIRONMENTAL####
    b.loc[n,'gw_b6_h']=b.loc[n,'ec_h']*c.loc[b.loc[n,'process_h'],'gw_b6']  #GPW  heating
    b.loc[n,'pe_b6_h']=b.loc[n,'ec_h']*c.loc[b.loc[n,'process_h'],'pe_b6']  #NRPE heating

    fc_h_base=fc_h_sum=b.loc[n,'ec_h']*c.loc[b.loc[n,'process_h'],'fc_b6'] ######ECONOMIC#######
    for y in range(1,int(b.loc[n,'rslb'])):
        fc_h_sum=fc_h_sum+(fc_h_base*(1+(c.loc[b.loc[n,'process_h'],'fc_in_b6']))**y)
    b.loc[n,'fc_b6_h']=fc_h_sum/b.loc[n,'rslb']                                 #Full Cost heating
    
################################################################################################
#A1-3 - A5 - B2 - B4_A1-3 - B4_A5 - Environmental & Economic
################################################################################################
#STAGE A1-3 ENV & EC
for n in b.index:
    for i in ('gw_a13','pe_a13','fc_a13'):
        b.loc[n,i]=((
        c.loc[b.loc[n,'mat_1'],i]*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*((b.loc[n,'th'])/100)+       
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
    for y in range(1,b.loc[n,'rslb']):
        fc_b2_sum=fc_b2_sum+(fc_b2_base*(1+(b.loc[n,'inf']))**y)
    b.loc[n,'fc_b2']=fc_b2_sum/(b.loc[n,'s']*b.loc[n,'rslb'])

#STAGE B4_A1-3 ENV & EC
for n in c.index:
    c.loc[n,'rp']=(math.ceil(b.loc['base','rslb']/c.loc[n,'rslm']))-1
for n in b.index:
    b.loc[n,'gw_b4_a13']=((
    c.loc[b.loc[n,'mat_1'],'gw_a13']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*c.loc[b.loc[n,'mat_1'],'rp']*((b.loc[n,'th'])/100)+
    c.loc[b.loc[n,'mat_2'],'gw_a13']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']*c.loc[b.loc[n,'mat_2'],'rp']+
    c.loc[b.loc[n,'mat_3'],'gw_a13']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med']*c.loc[b.loc[n,'mat_3'],'rp'])/(b.loc[n,'s']*b.loc[n,'rslb']))
    b.loc[n,'pe_b4_a13']=((
    c.loc[b.loc[n,'mat_1'],'pe_a13']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*c.loc[b.loc[n,'mat_1'],'rp']*((b.loc[n,'th'])/100)+
    c.loc[b.loc[n,'mat_2'],'pe_a13']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']*c.loc[b.loc[n,'mat_2'],'rp']+
    c.loc[b.loc[n,'mat_3'],'pe_a13']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med']*c.loc[b.loc[n,'mat_3'],'rp'])/(b.loc[n,'s']*b.loc[n,'rslb']))
    b.loc[n,'fc_b4_a13']=((
    c.loc[b.loc[n,'mat_1'],'fc_a13']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*c.loc[b.loc[n,'mat_1'],'rp']*((b.loc[n,'th'])/100)*((1+b.loc[n,'inf'])**c.loc[b.loc[n,'mat_1'],'rslm'])+
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
#TOTALS

b['gw_t']=b['gw_a13']+b['gw_b4_a13']+b['gw_b6_h']
b['pe_t']=b['pe_a13']+b['pe_b4_a13']+b['pe_b6_h']
b['fc_t']=b['fc_b6_h']+b['fc_a13']+b['fc_a5']+b['fc_b2']+b['fc_b4_a13']+b['fc_b4_a5']

b['gw_int']=b['gw_a13']+b['gw_b4_a13']                                      
b['fc_int']=b['fc_a13']+b['fc_a5']+b['fc_b2']+b['fc_b4_a13']+b['fc_b4_a5'] 
b['pe_int']=b['pe_a13']+b['pe_b4_a13'] 

b['pe_red']=b.loc['base','pe_b6_h']-b['pe_b6_h']
b['fc_red']=b.loc['base','fc_b6_h']-b['fc_b6_h']

################################################################################################
#RESULTS
#reindex
lca_opt=b.filter(['th','ed_h','pe_b6_h','pe_a13','pe_b4_a13',
'fc_b6_h','fc_a13','fc_a5','fc_b2','fc_b4_a13','fc_b4_a5','fc_red',
'fc_int','pe_red','pe_int'])

lca_opt=lca_opt.drop(['base'])

################################################################################################
################################################################################################
#ENVIORMENTAL OPTIMIZATION
print('ENVIRONMENTAL ASSESSMENT:')


#NRPE operational reduction (logaritmic)
xaxes1 = lca_opt['th'].values.reshape(-1,1)
xaxesnew = np.log(xaxes1) #new x axes being the logarit of the input, y = m*log(x_input)+c
yaxes1 = lca_opt['pe_red'].values.reshape(-1,1)
linear_regressor = LinearRegression()
linear_regressor.fit(xaxesnew,yaxes1)
yaxes1_pred = linear_regressor.predict(xaxesnew)
an = linear_regressor.coef_[0][0]
bn = linear_regressor.intercept_[0]
label1n = r'f(x) = %0.4f*ln(x) % +0.4f'%(an,bn)
#print('Operational reduction NRPE ',label1n)

#NRPE embodied (linear)
xaxes2 = lca_opt['th'].values.reshape(-1,1)
yaxes2 = lca_opt['pe_int'].values.reshape(-1,1)
linear_regressor2 = LinearRegression()
linear_regressor2.fit(xaxes2,yaxes2)
linear_regressor2.fit(xaxes2,yaxes2)
yaxes2_pred = linear_regressor2.predict(xaxes2)
mn = linear_regressor2.coef_[0,0]
cn = linear_regressor2.intercept_[0]
label2n= r'g(x) = %0.4f*x % +0.4f'%(mn,cn)
#print('Embodied NRPE ',label2n)

xn=np.setdiff1d(np.linspace(np.amin(lca_opt['th']),np.amax(lca_opt['th']),100),[0]) #to remove the zero
def f1n(xn):
    return an*np.log(xn)+bn
y1n=f1n(xn)
def f2n(xn):
    return mn*xn + cn
y2n=f2n(xn)
def f3n(xn):
    return (f1n(xn)) / (f2n(xn))
y3n=f3n(xn)

#PREDICTION
max_index=np.argmax(y3n)
opt_ner=int(xn[max_index])
print('Optimal NER (heating) ',round(np.amax(y3n),2))
print('Optimal TH ENVIRONMENTAL ',opt_ner)


################################################################################################
################################################################################################
#ENERGY DEMAND

xaxes1 = lca_opt['th'].values.reshape(-1,1)
xaxesnew = np.log(xaxes1) #new x axes being the logarit of the input, y = m*log(x_input)+c
yaxes1 = lca_opt['ed_h'].values.reshape(-1,1)
linear_regressor = LinearRegression()
linear_regressor.fit(xaxesnew,yaxes1)
yaxes1_pred = linear_regressor.predict(xaxesnew)
ae = linear_regressor.coef_[0][0]
be = linear_regressor.intercept_[0]
xe=np.setdiff1d(np.linspace(20,np.amax(lca_opt['th']),100),[0]) #to remove the zero
def f(xe):
    return ae*np.log(xe)+be

t3=time.time()


################################################################################################
################################################################################################
#ECONOMIC OPTIMIZATION
print('ECONOMIC ASSESSMENT:')


df=pd.DataFrame({'th':[0]})
for n in range(np.amin(lca_opt['th']),np.amax(lca_opt['th'])+1):
    df.loc[len(df.index)]=[n]
df.drop(index=df.index[0],axis=0,inplace=True)
for n in df.index:   
    df.loc[n,'ed']=f(df.loc[n,'th'])
    
    df.loc[n,'el']=(df.loc[n,'ed']*b.iloc[0,6])*3.6
    df.loc[n,'ec']=((df.loc[n,'ed']*3.6+df.loc[n,'el'])/b.iloc[0,5])

    fc_h_base=fc_h_sum=df.loc[n,'ec']*c.loc[b.iloc[0,7],'fc_b6'] ######ECONOMIC#######
    for y in range(1,b.iloc[0,1]):
        fc_h_sum=fc_h_sum+(fc_h_base*(1+(c.loc[b.iloc[0,7],'fc_in_b6']))**y)
    df.loc[n,'fc_b6_h']=fc_h_sum/b.iloc[0,1]                                 #Full Cost heating
    
    df.loc[n,'fc_a13']=((
        c.loc[b.loc['d0','mat_1'],'fc_a13']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']*((df.loc[n,'th'])/100)+       
        c.loc[b.loc['d0','mat_2'],'fc_a13']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_2_med']+
        c.loc[b.loc['d0','mat_3'],'fc_a13']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_3_med'])/(b.iloc[0,0]*b.iloc[0,1])) 
    
    df.loc[n,'fc_a5']=((
        c.loc[b.loc['d0','mat_1'],'fc_a5']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']+       
        c.loc[b.loc['d0','mat_2'],'fc_a5']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_2_med']+
        c.loc[b.loc['d0','mat_3'],'fc_a5']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_3_med'])/(b.iloc[0,0]*b.iloc[0,1])) 

    df.loc[n,'fc_b4_a13']=((
        c.loc[b.loc['d0','mat_1'],'fc_a13']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']*c.loc[b.loc['d0','mat_1'],'rp']*((df.loc[n,'th'])/100)*((1+b.iloc[0,2])**c.loc[b.loc['d0','mat_1'],'rslm'])+
        c.loc[b.loc['d0','mat_2'],'fc_a13']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_2_med']*c.loc[b.loc['d0','mat_2'],'rp']*((1+b.iloc[0,2])**c.loc[b.loc['d0','mat_2'],'rslm'])+
        c.loc[b.loc['d0','mat_3'],'fc_a13']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_3_med']*c.loc[b.loc['d0','mat_3'],'rp']*((1+b.iloc[0,2])**c.loc[b.loc['d0','mat_3'],'rslm']))
        /(b.iloc[0,0]*b.iloc[0,1]))   
    df.loc[n,'fc_b4_a5']=((
        c.loc[b.loc['d0','mat_1'],'fc_a5']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']*c.loc[b.loc['d0','mat_1'],'rp']*((1+b.iloc[0,2])**c.loc[b.loc['d0','mat_1'],'rslm'])+
        c.loc[b.loc['d0','mat_2'],'fc_a5']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_2_med']*c.loc[b.loc['d0','mat_2'],'rp']*((1+b.iloc[0,2])**c.loc[b.loc['d0','mat_1'],'rslm'])+
        c.loc[b.loc['d0','mat_3'],'fc_a5']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_3_med']*c.loc[b.loc['d0','mat_3'],'rp']*((1+b.iloc[0,2])**c.loc[b.loc['d0','mat_1'],'rslm']))
        /(b.iloc[0,0]*b.iloc[0,1]))
    
    fc_b2_base=fc_b2_sum=(
        c.loc[b.loc['d0','mat_1'],'fc_b2']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']+
        c.loc[b.loc['d0','mat_2'],'fc_b2']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_1_med']+
        c.loc[b.loc['d0','mat_3'],'fc_b2']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_1_med'])
    for y in range(1,int(b.iloc[0,1])):
        fc_b2_sum=fc_b2_sum+(fc_b2_base*(1+b.iloc[0,2])**y)
    df.loc[n,'fc_b2']=fc_b2_sum/(b.iloc[0,0]*b.iloc[0,1])
    
    
    df['fc_t']=df['fc_b6_h']+df['fc_a13']+df['fc_a5']+df['fc_b2']+df['fc_b4_a13']+df['fc_b4_a5']
for n in df.index:
    irr_inv=(df.loc[n,'fc_a13']+df.loc[n,'fc_a5']+df.loc[n,'fc_b4_a13']+df.loc[n,'fc_b4_a5'])*b.iloc[0,1]
    irr_ann=[None]*int(b.iloc[0,1])
    van=[None]*200
    van[0]=0
    irr_ann[0]=(-1)*irr_inv
    for x in range(1,200):
        irr_save=(b.loc['base','ec_h']-df.loc[n,'ec'])*(c.loc[b.loc['d0','process_h'],'fc_b6']*(1+c.loc[b.loc['d0','process_h'],'fc_in_b6'])**x)
        irr_b2=((  c.loc[b.loc['d0','mat_1'],'fc_b2']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']+
                   c.loc[b.loc['d0','mat_2'],'fc_b2']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_1_med']+
                   c.loc[b.loc['d0','mat_3'],'fc_b2']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_1_med'])/b.iloc[0,0])*((1+b.iloc[0,2])**x)
        van[x]=(van[x-1])+ irr_save - irr_b2
        #print(van[x])
    for x in range(1,int(b.iloc[0,1])):
        irr_save=(b.loc['base','ec_h']-df.loc[n,'ec'])*(c.loc[b.loc['d0','process_h'],'fc_b6']*(1+(c.loc[b.loc['d0','process_h'],'fc_in_b6']))**x)
        irr_b2=((c.loc[b.loc['d0','mat_1'],'fc_b2']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']+
                 c.loc[b.loc['d0','mat_2'],'fc_b2']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_1_med']+
                 c.loc[b.loc['d0','mat_3'],'fc_b2']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_1_med'])/(b.iloc[0,0]*b.iloc[0,1]))*((1+b.iloc[0,2])**x)
        irr_ann[x]=irr_save-irr_b2
        df.loc[n,'irr']=npf.irr(irr_ann)

#PREDICTION
column = df["irr"]
max_value = round(column.max(),4)
max_index = column.idxmax()
opt_irr=df.loc[max_index,'th']
print('Optimal IRR (heating) ',max_value)
print('Optimal TH ECONOMIC ',opt_irr)


################################################################################################
################################################################################################
#ENERGY DEMAND
print('ENERGETIC DEMANDS:')

ed_env=round(f(opt_ner),2)
ed_ec=round(f(opt_irr),2)
print('Energy demand (heating) of the environmental optimal): ',ed_env)
print('Energy demand (heating) of the economic optimal): ',ed_ec)

t4=time.time()

################################################################################################
################################################################################################
################################################################################################
#Export to Excel

opt=pd.DataFrame({'scope':['opt_env','opt_ec'],'th_[mm]':[int(opt_ner),int(opt_irr)],'ed_opt':[ed_env,ed_ec]})
opt.set_index(['scope'])

shutil.copy(input,output)
book = load_workbook(output)
writer = pd.ExcelWriter(output,engine='openpyxl')
writer.book=book
opt.to_excel(writer, sheet_name='output')
writer.save()
print('Results have been saved correcty')

t5=time.time()

################################################################################################
#EXECUTION TIME
print('EXECUTION TIMES:')
print('Input data import execution time',(t2-t1)*1000,'ms')
print('Env assessment execution time: ',(t3-t2)*1000,'ms')
print('Ec assessment execution time: ',(t4-t3)*1000,'ms')
print('Output data export execution time: ',(t5-t4)*1000,'ms')
