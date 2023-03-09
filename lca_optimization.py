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
    #env
    b.loc[n,'gw_b4_a13']=((
    c.loc[b.loc[n,'mat_1'],'gw_a13']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*c.loc[b.loc[n,'mat_1'],'rp']*((b.loc[n,'th']))+
    c.loc[b.loc[n,'mat_2'],'gw_a13']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']*c.loc[b.loc[n,'mat_2'],'rp']+
    c.loc[b.loc[n,'mat_3'],'gw_a13']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med']*c.loc[b.loc[n,'mat_3'],'rp'])/(b.loc[n,'s']*b.loc[n,'rslb']))
    b.loc[n,'pe_b4_a13']=((
    c.loc[b.loc[n,'mat_1'],'pe_a13']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*c.loc[b.loc[n,'mat_1'],'rp']*((b.loc[n,'th']))+
    c.loc[b.loc[n,'mat_2'],'pe_a13']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']*c.loc[b.loc[n,'mat_2'],'rp']+
    c.loc[b.loc[n,'mat_3'],'pe_a13']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med']*c.loc[b.loc[n,'mat_3'],'rp'])/(b.loc[n,'s']*b.loc[n,'rslb']))
    

    #ec (taking into account the inflation for each repacement year)
    if int(c.loc[b.loc[n,'mat_1'],'rp']) != 0:
        for y in range(1,int(1+c.loc[b.loc[n,'mat_1'],'rp'])):
            fc_b4_a13_1=c.loc[b.loc[n,'mat_1'],'fc_a13']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*c.loc[b.loc[n,'mat_1'],'rp']*((b.loc[n,'th']))*((1+b.loc[n,'inf'])**((c.loc[b.loc[n,'mat_1'],'rslm'])*y))
    else: fc_b4_a13_1=0
    if c.loc[b.loc[n,'mat_2'],'rp'] != 0:
        for y in range(1,int(1+c.loc[b.loc[n,'mat_2'],'rp'])):
            fc_b4_a13_2=c.loc[b.loc[n,'mat_2'],'fc_a13']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']*c.loc[b.loc[n,'mat_2'],'rp']*((1+b.loc[n,'inf'])**((c.loc[b.loc[n,'mat_2'],'rslm'])*y))    
    else: fc_b4_a13_2=0        
    if c.loc[b.loc[n,'mat_3'],'rp'] != 0:
        for y in range(1,int(1+c.loc[b.loc[n,'mat_1'],'rp'])):
            fc_b4_a13_3=c.loc[b.loc[n,'mat_3'],'fc_a13']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med']*c.loc[b.loc[n,'mat_3'],'rp']*((1+b.loc[n,'inf'])**((c.loc[b.loc[n,'mat_3'],'rslm'])*y))
    else: fc_b4_a13_3=0        
    b.loc[n,'fc_b4_a13']=((fc_b4_a13_1+fc_b4_a13_2+fc_b4_a13_3)/((b.loc[n,'s']*b.loc[n,'rslb'])))

#STAGE B4_A5 EC

for n in b.index:        
    if c.loc[b.loc[n,'mat_1'],'rp'] != 0:
            for y in range(1,int(1+c.loc[b.loc[n,'mat_1'],'rp'])):
                fc_b4_a5_1=c.loc[b.loc[n,'mat_1'],'fc_a5']*c.loc[b.loc[n,'mat_1'],'conv']*b.loc[n,'mat_1_med']*((1+b.loc[n,'inf'])**((c.loc[b.loc[n,'mat_1'],'rslm'])*y))
    else: fc_b4_a5_1=0        
    if c.loc[b.loc[n,'mat_2'],'rp'] != 0:
            for y in range(1,int(1+c.loc[b.loc[n,'mat_2'],'rp'])):
                fc_b4_a5_2=c.loc[b.loc[n,'mat_2'],'fc_a5']*c.loc[b.loc[n,'mat_2'],'conv']*b.loc[n,'mat_2_med']*((1+b.loc[n,'inf'])**((c.loc[b.loc[n,'mat_2'],'rslm'])*y))    
    else: fc_b4_a5_2=0         
    if c.loc[b.loc[n,'mat_3'],'rp'] != 0:
            for y in range(1,int(1+c.loc[b.loc[n,'mat_1'],'rp'])):
                fc_b4_a5_3=c.loc[b.loc[n,'mat_3'],'fc_a5']*c.loc[b.loc[n,'mat_3'],'conv']*b.loc[n,'mat_3_med']*((1+b.loc[n,'inf'])**((c.loc[b.loc[n,'mat_3'],'rslm'])*y))
    else: fc_b4_a5_3=0 
    b.loc[n,'fc_b4_a5']=((fc_b4_a5_1+fc_b4_a5_2+fc_b4_a5_3)/((b.loc[n,'s']*b.loc[n,'rslb'])))
    
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
    #ENERGY DEMAND, ENERGY LOSSES, ENERGY CONSUMPTION
    df.loc[n,'ed']=f(df.loc[n,'th'])
    df.loc[n,'el']=(df.loc[n,'ed']*b.loc['d0','el'])*3.6
    df.loc[n,'ec']=((df.loc[n,'ed']*3.6+df.loc[n,'el'])/b.loc['d0','eef_h'])

    #FC b6
    fc_h_base=fc_h_sum=df.loc[n,'ec']*c.loc[b.loc['d0','process_h'],'fc_b6']
    for y in range(1,int(b.loc['d0','rslb'])):
        fc_h_sum=fc_h_sum+(fc_h_base*(1+(c.loc[b.loc['d0','process_h'],'fc_in_b6']))**y)
    df.loc[n,'fc_b6_h']=fc_h_sum/b.loc['d0','rslb']                                
    #FC a1-3
    df.loc[n,'fc_a13']=((
        c.loc[b.loc['d0','mat_1'],'fc_a13']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']*((df.loc[n,'th'])/100)+       
        c.loc[b.loc['d0','mat_2'],'fc_a13']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_2_med']+
        c.loc[b.loc['d0','mat_3'],'fc_a13']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_3_med'])/(b.loc['d0','s']*b.loc['d0','rslb']))
    #FC a5    
    df.loc[n,'fc_a5']=((
        c.loc[b.loc['d0','mat_1'],'fc_a5']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']+       
        c.loc[b.loc['d0','mat_2'],'fc_a5']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_2_med']+
        c.loc[b.loc['d0','mat_3'],'fc_a5']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_3_med'])/(b.loc['d0','s']*b.loc['d0','rslb'])) 
    #FC b4-A1-3
    for y in c.index:
        c.loc[y,'rp']=(math.ceil(b.loc['d0','rslb']/c.loc[y,'rslm']))-1
    if int(c.loc[b.loc['d0','mat_1'],'rp']) != 0:
        for y in range(1,int(1+c.loc[b.loc['d0','mat_1'],'rp'])):
            fc_b4_a13_1=c.loc[b.loc['d0','mat_1'],'fc_a13']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']*c.loc[b.loc['d0','mat_1'],'rp']*((df.loc[n,'th'])/100)*((1+b.loc['d0','inf'])**((c.loc[b.loc['d0','mat_1'],'rslm'])*y))
    else: fc_b4_a13_1=0
    if c.loc[b.loc['d0','mat_2'],'rp'] != 0:
        for y in range(1,int(1+c.loc[b.loc['d0','mat_2'],'rp'])):
            fc_b4_a13_2=c.loc[b.loc['d0','mat_2'],'fc_a13']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_2_med']*c.loc[b.loc['d0','mat_2'],'rp']*((1+b.loc['d0','inf'])**((c.loc[b.loc['d0','mat_2'],'rslm'])*y))    
    else: fc_b4_a13_2=0        
    if c.loc[b.loc['d0','mat_3'],'rp'] != 0:
        for y in range(1,int(1+c.loc[b.loc['d0','mat_1'],'rp'])):
            fc_b4_a13_3=c.loc[b.loc['d0','mat_3'],'fc_a13']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_3_med']*c.loc[b.loc['d0','mat_3'],'rp']*((1+b.loc['d0','inf'])**((c.loc[b.loc['d0','mat_3'],'rslm'])*y))
    else: fc_b4_a13_3=0        
    df.loc[n,'fc_b4_a13']=((fc_b4_a13_1+fc_b4_a13_2+fc_b4_a13_3)/((b.loc['d0','s']*b.loc['d0','rslb'])))
    #FC b4-A5
    if c.loc[b.loc['d0','mat_1'],'rp'] != 0:
        for y in range(1,int(1+c.loc[b.loc['d0','mat_1'],'rp'])):
            fc_b4_a5_1=c.loc[b.loc['d0','mat_1'],'fc_a5']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']*((1+b.loc['d0','inf'])**((c.loc[b.loc['d0','mat_1'],'rslm'])*y))
    else: fc_b4_a5_1=0        
    if c.loc[b.loc['d0','mat_2'],'rp'] != 0:
            for y in range(1,int(1+c.loc[b.loc['d0','mat_2'],'rp'])):
                fc_b4_a5_2=c.loc[b.loc['d0','mat_2'],'fc_a5']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_2_med']*((1+b.loc['d0','inf'])**((c.loc[b.loc['d0','mat_2'],'rslm'])*y))    
    else: fc_b4_a5_2=0         
    if c.loc[b.loc['d0','mat_3'],'rp'] != 0:
            for y in range(1,int(1+c.loc[b.loc['d0','mat_1'],'rp'])):
                fc_b4_a5_3=c.loc[b.loc['d0','mat_3'],'fc_a5']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_3_med']*((1+b.loc['d0','inf'])**((c.loc[b.loc['d0','mat_3'],'rslm'])*y))
    else: fc_b4_a5_3=0 
    df.loc[n,'fc_b4_a5']=((fc_b4_a5_1+fc_b4_a5_2+fc_b4_a5_3)/((b.loc['d0','s']*b.loc['d0','rslb'])))
    #FC b2
    fc_b2_base=fc_b2_sum=(
        c.loc[b.loc['d0','mat_1'],'fc_b2']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']+
        c.loc[b.loc['d0','mat_2'],'fc_b2']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_1_med']+
        c.loc[b.loc['d0','mat_3'],'fc_b2']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_1_med'])
    for y in range(1,int(b.loc['d0','rslb'])):
        fc_b2_sum=fc_b2_sum+(fc_b2_base*(1+b.loc['d0','inf'])**y)
    df.loc[n,'fc_b2']=fc_b2_sum/(b.loc['d0','s']*b.loc['d0','rslb'])
    #FC totals
    df['fc_t']=df['fc_b6_h']+df['fc_a13']+df['fc_a5']+df['fc_b2']+df['fc_b4_a13']+df['fc_b4_a5']

    #prioritization IRR

    irr_inv=(df.loc[n,'fc_a13']+df.loc[n,'fc_a5']+df.loc[n,'fc_b4_a13']+df.loc[n,'fc_b4_a5'])*b.loc['d0','rslb']
    irr_ann=[None]*int(b.loc['d0','rslb'])
    van=[None]*200
    van_abs=[None]*200
    irr_save=[None]*200
    irr_b2=[None]*200
    for x in range(0,200):
        epi_h_1=c.loc[b.loc['base','process_h'],'fc_in_b6']
        epi_h_2=c.loc[b.loc['d0','process_h'],'fc_in_b6']
        fc_b6h_1=c.loc[b.loc['base','process_h'],'fc_b6']
        fc_b6h_2=c.loc[b.loc['d0','process_h'],'fc_b6']
        irr_save_h=(((b.loc['base','ec_h']*fc_b6h_1*(1+epi_h_1)**x))-
                    (df.loc[n,'ec']*     fc_b6h_2*(1+epi_h_2)**x))          
        irr_b2[x]=((c.loc[b.loc['d0','mat_1'],'fc_b2']*c.loc[b.loc['d0','mat_1'],'conv']*b.loc['d0','mat_1_med']+
                    c.loc[b.loc['d0','mat_2'],'fc_b2']*c.loc[b.loc['d0','mat_2'],'conv']*b.loc['d0','mat_2_med']+
                    c.loc[b.loc['d0','mat_3'],'fc_b2']*c.loc[b.loc['d0','mat_3'],'conv']*b.loc['d0','mat_3_med'])
                    /b.loc['d0','s'])*((1+b.loc['d0','inf'])**x)
        irr_save[x]= irr_save_h
        if x==0:
            van[x]=(-1)*irr_inv
            irr_ann[x]=(-1)*irr_inv
        else:
            van[x]=(van[x-1])+ irr_save[x] - irr_b2[x]
        van_abs[x]=int(abs(van[x]))
    for x in range(1,int(b.loc['d0','rslb'])):
        irr_ann[x]=irr_save[x]-irr_b2[x]
    df.loc[n,'irr']=npf.irr(irr_ann)
    try:
        van_array=np.array(van_abs)
        lcp=van_array.argmin()
        df.loc[n,'lcpb']=lcp
    except: df.loc[n,'lcpb']= 0
    if lcp==200: df.loc[n,'lcpb']= "NaN"

#PREDICTION
column = df["irr"]
max_value = round(column.max(),4)
max_index = column.idxmax()
opt_irr=df.loc[max_index,'th']
opt_lcpb=df.loc[max_index,'lcpb']
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

opt=pd.DataFrame({'scope':['opt_env','opt_ec'],'th_[mm]':[int(opt_ner),int(opt_irr)],'ed_opt':[ed_env,ed_ec],'NER_opt':[round(np.amax(y3n),2),'-'],'IRR_opt':['-',max_value],'LCPB_opt':['-',opt_lcpb]})
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
