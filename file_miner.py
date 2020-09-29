# -------------------------------------------------------------- HOURS REGISTERED ------------------------------------------

def hoursregister(filepath):
    #Add appropriate paddings to each if condition. 
    #Merge the conditions. 
    #Use Ffill and Bfill to fill in emp names and numbers. 
    #Use groupbys to figure some logic out

    temp=[]
    newtemp=[]

    flag = 0

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        while line:
            linecount = linecount + 1
            line = fp.readline()

            if re.findall('(PERIOD END DATE (0|1)\d{1})-((0|1|2)\d{1})-((19|20)\d{2})',line) and flag==0:
                date = (re.search('(PERIOD END DATE) ((0|1)\d{1})-((0|1|2)\d{1})-((19|20)\d{2})',line)[0].split()[-1])
                flag = 1

            if len(line.strip()) and line[:len('HOURS REGISTER')]!='HOURS REGISTER' and line[:len('RUN DATE')]!='RUN DATE' and line[:len('EMPLOYEE NAME')]!='EMPLOYEE NAME':
                #print(line)
                if line[:3].strip().isnumeric():
                    #print(line)
                    #print(line[:2].strip(),line[2:8].strip(),line[8:14].strip(),line[14:26].strip(),line[26:31].strip(), line[31:45].strip(),line[45:56].strip(),line[56:].strip())                
                    temp.append([line[:2].strip()]+[line[2:8].strip()]+[line[8:14].strip()]+[line[14:26].strip()]+[None]+
                                [line[26:31].strip()]+[line[31:35]]+[line[35:45].strip()]+[line[45:56].strip()]+[line[56:].strip()])
                if line[:3].strip().isalpha():
                    #print(line)
                    #print(line[:33],line[33:35],line[36:44],line[44:56],line[56:])
                    temp.append([None]+[None]+[None]+[None]+[line[:33].strip()]+[None]+[line[33:35].strip()]+[line[36:44].strip()]+
                                [line[44:56].strip()]+[line[56:].strip()])  

                if ~line[:3].strip().isalpha() and ~line[:3].strip().isnumeric() and (line[33:35].strip().isnumeric() or line[36:39] == 'TOT' or line[36:44] == 'VACAVAIL') and not re.search('° 0',line[:33]):
                    #print(line)
                    temp.append([None]+[None]+[None]+[None]+[line[:33].strip()]+[None]+[line[33:36].strip()]+[line[36:44].strip()]+
                                [line[44:56].strip()]+[line[56:].strip()]) 


    df = pd.DataFrame(temp,columns=['L3','L4','L5','EMP NUMBER','EMP NAME','TYPE','HC','HOUR DESCRIPTION','Current Hours','YTD HOURS'])

    df = df[df['EMP NUMBER']!='L3 TOTALS']
    df = df[df['EMP NUMBER']!='L4 TOTALS']
    df = df[df['EMP NUMBER']!='L2 TOTALS']
    df = df[df['EMP NUMBER']!='TOTAL CMQ']
    df = df[df['EMP NUMBER']!='TOTALS CMQ']
    df = df[df['EMP NUMBER']!='TOTALS  CMQ']

    df['HC'][df['HC']=='REGU']=''
    df['HC'][df['HC']=='GU']=''

    df['HC'][df['HC']=='']=0
    df['HC'][df['HC']=='    ']=0
    #df['EMP NUMBER'][df['EMP NUMBER']=='L3 TOTALS']=None

    df['YTD HOURS'][df['YTD HOURS']=='']=0

    df['HOUR DESCRIPTION'][df['HOUR DESCRIPTION']=='LAR']='REGULAR'

    df['EMP NAME'][df['EMP NAME']==''] = None
    df[['L3','L4','L5','EMP NUMBER']] = df[['L3','L4','L5','EMP NUMBER']].ffill()
    df['EMP NAME'] = df['EMP NAME'].bfill(limit=1)
    df['EMP NAME'] = df['EMP NAME'].ffill()

    df['HC'] = df['HC'].astype(float)
    df['YTD HOURS'] = pd.to_numeric(df['YTD HOURS'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)

    df['MGMT CNTR'] = df['L3']+df['L4']
    df['EMP NUMBER'][df['EMP NUMBER'].str.len()>5] = df['EMP NUMBER'][df['EMP NUMBER'].str.len()>5].str[4:]
    df['Current Hours'] = pd.to_numeric(df['Current Hours'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)

    #Summary Output
    #a = df[df['HOUR DESCRIPTION']!=''].groupby(['EMP NUMBER','EMP NAME','HOUR DESCRIPTION'])[['Current Hours','HC']].agg({'sum'})

    df['TYPE'] = df['TYPE'].ffill()
    df['Period End Date'] = pd.to_datetime(date)
    
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )

        # Excel output
    with pd.ExcelWriter('{}//Hours Registered - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
        #a.to_excel(writer, sheet_name='Summary')
    display(df)

# ------------------------------------------ RRD REG OT HOURS----------------------------------------------

def rrdreg_othrswages(filepath):
    #filepath = 'Files/RRD REG_OT HRS WAGES.TXT'
    filepath = 'Files/RRD REG_OT HRS WAGES.TXT'

    temp=[]
    newtemp=[]

    flag=0

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])

        while line:
            line = fp.readline()

            #print(re.findall('(OVERTIME HOURS AND WAGES FOR (0|1)\d{1})/((0|1|2)\d{1})/((19|20)\d{2})',line.strip()))

            if(len(line.strip())):
                newline = line.strip()
                if(newline[:5].strip() in ['BLE','BMWE','BRC','BRS','IAM','IBEW','SMWIA','TCU','TWU','UTU B','UNION']):
                    #print(newline)
                    temp.append([newline[:5]]+newline[5:].split())
    
    
    df = pd.DataFrame(temp)
    union_list = df[0].loc[(df[0]=='UNION')].index

    for i in range(len(union_list) -1):
        if i == 0:
            temp =  df.loc[union_list[i]+1:union_list[i+1]-1:]
        else:
            temp = pd.merge(temp,df.loc[union_list[i]+1:union_list[i+1]-1:],on=0)
        
    headers= ['UNION','REG HRS MTD','OT HRS MTD','TOTAL HRS','REGGRMTD','OTGRSMTD','OTGSMTD','TOTALEARN','NSPEARN','VAC EARN','HOL EARN','ANN EARN','BRTHDY EARN','PER EARN','MIL EARN','SICK EARN','FAM ILL EARN','FAMILY DTH EARN','WRKM COMP EARN','JURY DUTY EARN','COURT ATTD EARN','FAM ILL EARN','UNION BUS EARN','VISION EARN','ACCT REPT EARN','UNIFORM EARN','MEAL EARN','INJ ON DUTY EARN','NSP HRS','VAC HRS','BRTHDY HRS','ANNV HRS','HOL HRS','PERS HRS','SICK PAY HRS','ON DUTY HRS','JURY DTY HRS','MIL HRS','COURT ATTD HRS','FMLY DTH HRS','SICK PAY HRS','ODI HRS','UNION BUS HRS']
        
    df = temp.rename(columns=dict(zip(list(temp),headers)))

    for col in list(df)[1:-1]:
        exec("df[\'{}\'] = df[\'{}\'].apply(pd.to_numeric,errors='coerce').fillna(0).astype(float)".format(col,col))

    date = str(date.date())
    df['Report Date'] = date

    display(df)
    
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    
    # Excel output
    with pd.ExcelWriter('{}//RRD REG_OT HOURS WAGES - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)

# ---------------------------------------------------- MPC WORK COMP REPORT -------------------------------------------------

def mpcworkcompreport(filepath):
    #filepath = 'Files/MPC WORK COMP REPORT.TXT'

    filepath = 'Files/MPC WORK COMP REPORT.TXT'

    temp=[]
    newtemp=[]

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])


        while line:
            linecount = linecount + 1
            line = fp.readline()

            if(len(line)):
                if re.search('[0-9][0-9][0-9][0-9]',line[:4]):
                    #print(line)
                    temp.append(line[:48].split()+[line[48:53]]+[line[53:74]]+[line[74:86]]+
                                [line[86:92]]+[line[92:114]]+[line[114:119]]+[line[119:128]]+[line[128:131]])
                
    df = pd.DataFrame(temp,columns=['COMP','COST CTR','ACCT','CTR','FC','DR CR','AMOUNT','JE#',
                               'JOURNAL DESCRIPTION','EFF DATE','EMPL NUM','EMPLOYEE NAME','DAYS IOD','DESCRIPT','3RD PTY ADM'])

    #df['HOURS WORKED'] = df['HOURS WORKED'].astype(float) 
    df['EFF DATE'] = pd.to_datetime(df['EFF DATE'])
    df['AMOUNT'] = pd.to_numeric(df['AMOUNT'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)
    df['JE#'] = np.where(df['JE#']=='     ',np.NaN,df['JE#'])
    df['JE#'] = df['JE#'].ffill()
    df['Report Date'] = date

    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    date = str(date.date())
    
    display(df)
    #Excel output
    with pd.ExcelWriter('{}//MPC WORK COMP REPORT - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
    
#--------------------------------------------- MPC TREAWORK REPORT-----------------------------------    
def mpctreaworkreport(filepath):
    #filepath = 'Files/MPC TREA WORK REPORT.TXT'
    temp=[]

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])

        while line and linecount<=lc:
            linecount = linecount + 1
            line = fp.readline()
            if(len(line.strip())):
                if re.search('([0-9][0-9][0-9][0-9])',line[:5]):
                    temp.append(line[:45].split()+[line[45:49]]+[line[50:71]]+
                                [line[72:81]]+[line[83:92]]+[line[93:98]]+[line[99:114]])

    df = pd.DataFrame(temp,columns=['COMP','COST CTR','ACCT','FC','DR CR','AMOUNT','JE#','JOURNAL DESCRIPTION','EFF DATE','TRANS DATE','DESC CODE','DESCRIPTION'])

    df['EFF DATE'] = pd.to_datetime(df['EFF DATE'])
    df['TRANS DATE'] = pd.to_datetime(df['TRANS DATE'])
    df['AMOUNT'] = pd.to_numeric(df['AMOUNT'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)    
    df['Report Date'] = date
    df['JE#'] = np.where(df['JE#']=='    ',np.NaN,df['JE#'])
    df['JE#'] = df['JE#'].ffill()
    
    display(df)
    
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    date = str(date.date())

    #Excel output
    with pd.ExcelWriter('{}//MPC TREAWORK REPORT - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
    
    
#------------------------------- MPC AP CLAIMS REPORT----------------------------------------------------------------

def mpcapclaimsreport(filepath):
    #filepath = 'Files/MPC AP CLAIMS REPORT.TXT'

    temp=[]
    newtemp=[]

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])

        while line and linecount<=lc:
            linecount = linecount + 1
            line = fp.readline()
            if(len(line.strip())):
                if re.search('([0-9][0-9][0-9][0-9])',line[:4]):
                    temp.append(line[:49].split()+[line[49:53]]+[line[54:64]]+
                                [line[65:77]]+[line[78:90]]+[line[91:93]]+[line[93:109]]+[line[110:120]]+[line[121:131]])

    df = pd.DataFrame(temp,columns=['COMP','COST CTR','ACCT','FC','CTR','DR CR','AMOUNT','JE#','EFF DATE',
                               'INVOICE','VENDOR','INV LINE','CHECK NUMBER','CHECK DATE','INV DATE'])
    df['Report Date'] = date
    df['JE#'] = np.where(df['JE#']=='   ',np.NaN,df['JE#'])
    df['JE#'] = df['JE#'].ffill()
    
    display(df)
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    date = str(date.date())

    # Excel Report
    with pd.ExcelWriter('{}//MPC AP CLAIMS REPORT - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
    
# ---------------------------------------------------- MPC MAT AND SUPP-----------------------------------------------------------------

def mpcmatandsupp(filepath):
    #filepath = 'Files/MPC MAT AND SUPP.TXT'

    temp=[]
    newtemp=[]

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])
        while line:
            linecount = linecount + 1
            line = fp.readline()
            if(len(line.strip())):
                if re.search('([0-9][0-9][0-9][0-9])',line[:4]):
                    temp.append(line[:49].split() + [line[50:55]] + [line[55:69]] + [line[70:87]] + [line[87:92]] + [line[92:95]] +
                          [line[95:103]] + [line[103:112]] + [line[112:115]] + [line[115:126]])
                    #print(temp)
    df = pd.DataFrame(temp,columns=['COMP', 'MGMT CTR', 'ACCT', 'FUNC', 'COST CTR', 'DR CR', 'AMOUNT', 'JE#', 'JOURNAL DESCRIPTION', 'DOC ID', 'LINE', 'CL', 'LOT', 'WORK ORD', 'TR', 'ITEM DESC'])

    df['AMOUNT'] = pd.to_numeric(df['AMOUNT'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)
    df['Report Date'] = date
    df['JE#'] = np.where(df['JE#']=='     ',np.NaN,df['JE#'])
    df['JE#'] = df['JE#'].ffill()
    
    
    display(df)
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    date = str(date.date())
    
    with pd.ExcelWriter('{}//MPC MAT AND SUPPLY REPORT - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
    
# ------------------------------------------------------------------ MPC GL JOURNALS------------------------------------------------------

def mpcgljournals(filepath):
    #filepath = 'Files/MPC GL JOURNALS.TXT'

    temp=[]
    newtemp=[]

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])
        while line:
            linecount = linecount + 1
            line = fp.readline()
            if(len(line.strip())):
                if re.search('([0-9][0-9][0-9][0-9])',line[:17]):
                    temp.append(line[:45].strip().split()+[line[46:51].strip()]+[line[52:73].strip()]+[line[74:97].strip()]+
                                [line[97:107].strip()]+[line[108:119].strip()])

    df = pd.DataFrame(temp,columns=['COMP','MGMT CNTR','ACCT','COST CNTR','FC','DR CR','JE#','AMOUNT','JOURNAL','ENTRY DATE','EFFECTIVE DATE'])

    df['EFFECTIVE DATE'] = pd.to_datetime(df['EFFECTIVE DATE'])
    df['ENTRY DATE'] = pd.to_datetime(df['ENTRY DATE'])
    df['AMOUNT'] = pd.to_numeric(df['AMOUNT'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)
    df['Report Date'] = date    
    df['JE#'] = np.where(df['JE#']=='',np.NaN,df['JE#'])
    df['JE#'] = df['JE#'].ffill()
    
    
    display(df)
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    date = str(date.date())

    with pd.ExcelWriter('{}//MPC GL JOURNALS - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
    
#------------------------------------------------------------- MPC REVENUE------------------------------------------------------

def mpcrevenue(filepath):
    #filepath = 'Files/MPC REVENUE.TXT'

    temp=[]
    newtemp=[]

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])
        while line:
            linecount = linecount + 1
            line = fp.readline()
            if(len(line.strip())):
                if re.search('([0-9][0-9][0-9][0-9])',line[:17]):
                    temp.append(line[:45].strip().split()+[line[46:51].strip()]+[line[52:73].strip()]+
                                [line[74:97].strip()]+[line[97:107].strip()]+[line[108:119].strip()])


    df = pd.DataFrame(temp,columns=['COMP','MGMT CNTR','ACCT','COST CNTR','FC','DR CR','JE#','AMOUNT','JOURNAL','ENTRY DATE','EFFECTIVE DATE'])

    df['EFFECTIVE DATE'] = pd.to_datetime(df['EFFECTIVE DATE'])
    df['ENTRY DATE'] = pd.to_datetime(df['ENTRY DATE'])
    df['AMOUNT'] = pd.to_numeric(df['AMOUNT'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)
    df['Report Date'] = date
    
    df['JE#'] = np.where(df['JE#']=='',np.NaN,df['JE#'])
    df['JE#'] = df['JE#'].ffill()

    display(df)
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    date = str(date.date())

    # Excel Output
    with pd.ExcelWriter('{}//MPC REVENUE REPORT - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
        
# ------------------------------------------------------------------ MPC PROV LIAB----------------------------------------------------------

def mpcprovliab(filepath):
    #filepath = 'Files/MPC PROV LIAB.TXT'
    temp=[]
    newtemp=[]

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])

        while line:
            linecount = linecount + 1
            line = fp.readline()
            if(len(line)):
                if re.search('[0-9][0-9][0-9][0-9]',line[:8].strip()): 
                    temp.append(line[:52].split()+[line[53:57].strip()]+[line[58:70].strip()]+[line[70:75].strip()]+[line[75:85].strip()]+[line[86:97].strip()]+[line[97:127].strip()])

    df = pd.DataFrame(temp,columns=['COMP','MGMT CTR','ACCT','FC','COST CTR','DR CR','AMOUNT','JE#','DESCRIPTION','PAYMENT','VENDOR','PO NUM','ITEM DESCRIPTION'])
    df['AMOUNT'] = pd.to_numeric(df['AMOUNT'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)
    df['Report Date'] = date    
    date = str(date.date())

    df['JE#'] = np.where(df['JE#']=='',np.NaN,df['JE#'])
    df['JE#'] = df['JE#'].ffill()
    display(df)

    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )

    with pd.ExcelWriter('{}//MPC PROV LIAB REPORT - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)


# --------------------------------------------------------MPC PAYROLL ----------------------------------------------------------------------

def mpcpayroll(filepath):
    #filepath = 'Files/MPC PAYROLL.TXT'

    temp=[]

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])
        while line:
            linecount = linecount + 1
            line = fp.readline()
            if(len(line)):
                if re.search('[0-9][0-9][0-9][0-9]',line[:6]):
                    temp.append(line[:48].split()+[line[48:53]]+[line[53:74]]+[line[74:86]]+[line[86:91]]+[line[92:99]]+[line[100:105]]+
                                [line[105:110]]+[line[111:114]]+[line[116:121]]+[line[123:130]])

    df = pd.DataFrame(temp,columns=['COMP','MGMT CNTR','ACCT','FUNC','DR CR','AMOUNT',
                                  'JE#','JOURNAL DESCRIPTION','EFF DATE','EMPL NUM','WORK ORDER',
                                  'CHG POS','REG POS','DAY CODE','HOURS WORKED','PAY RATE'])
    df['HOURS WORKED'] = df['HOURS WORKED'].astype(float) 
    df['EFF DATE'] = pd.to_datetime(df['EFF DATE'])
    df['AMOUNT'] = pd.to_numeric(df['AMOUNT'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)
    df['Report Date'] = date

    display(df)
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    date = str(date.date())
    
    df['JE#'] = np.where(df['JE#']=='     ',np.NaN,df['JE#'])
    df['JE#'] = df['JE#'].ffill()
        
    with pd.ExcelWriter('{}//MPC PAYROLL REPORT - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)


# ----------------------------------------------------- MPC ACCT RECV REPORT -------------------------------------------------------

def mpcacctrecvreport(filepath):
    #filepath = 'Files/MPC ACCT RECV REPORT.TXT'
    temp=[]

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])
        while line:
            linecount = linecount + 1
            line = fp.readline()
            #print(line.strip())
            if(len(line)):
                if (line[:5].strip() != '') and not(re.search('PAGE ',line)):
                    temp.append([line[0:5].strip(),line[5:11].strip(),line[11:16].strip(),line[17:20].strip(),line[20:26].strip(),
                                 line[25:29].strip(),line[29:49].strip(),line[50:56].strip(),line[56:68].strip(),line[68:71].strip(),
                                 line[72:77].strip(),line[77:87].strip(),line[87:103].strip(),line[103:116].strip(),line[116:].strip()])

    df= pd.DataFrame(temp,columns=['COMP','MGMT CHECK CTR', 'ACCT', 'FUNC' ,'COST CTR', 'DR CR', 'AMOUNT', 'JE#', 'EFF DATE', 'BTCH', 'TRAIN ID', 'CUST NUM', 'CUST NAME', 'REF NO','DESCRIPT'])
    #df = df[(df['COMP'].str.isnumeric()) & ((df['COMP'].str.len())==4) & (~(df['COMP'].str.contains('/'))) & ((df['INV DATE'].str.contains('/')))]
    df = df[df['COMP'].str.match('^[0-9]*$')]
    df['EFF DATE'] = pd.to_datetime(df['EFF DATE'])
    df['AMOUNT'] = pd.to_numeric(df['AMOUNT'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)
    df['Report Date'] = date
    df['JE#'] = np.where(df['JE#']=='',np.NaN,df['JE#'])
    df['JE#'] = df['JE#'].ffill()
    
    display(df)
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    date = str(date.date())

    with pd.ExcelWriter('{}//MPC ACCT RECV REPORT - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)


# ---------------------------------------------------------- MGT DTL DRWN F_STICK----------------------------------------------------------

def mgtdtldrwnf_stck(filepath):
    #filepath = 'Files/MGT DTL DRWN F_STCK.TXT'

    temp=[]
    flag=0
    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()

        while line:
            linecount = linecount + 1
            line = fp.readline()
            #print(line.strip())
            if(len(line)):
                if re.findall('(ENDING) ((0|1)\d{1})/((0|1|2)\d{1})/((19|20)\d{2})',line) and flag==0:
                    date = (re.search('(ENDING) ((0|1)\d{1})/((0|1|2)\d{1})/((19|20)\d{2})',line)[0].split()[-1])
                    flag = 1
                if (line[:4].isnumeric()):
                    temp.append([line[0:4].strip(),line[5:11].strip(),line[11:16].strip(),line[16:20].strip(),line[20:33].strip(),
                                line[34:45].strip(),line[45:49].strip(),line[49:59].strip(),line[59:85].strip(),line[85:93].strip(),line[93:112].strip(),
                                line[112:131].strip(),line[131:120].strip()])

    df= pd.DataFrame(temp[1:],columns=['COMP', 'ACCT', 'FN','COST CNTR','REQ#', 'REQ DATE', 'CLASS', 'LOT', 'ITEM DESCRIPTION', 'QUANTITY', 'AMOUNT', 'JE/NO', 'WORD ORD#'])
    df['REQ DATE'] = pd.to_datetime(df['REQ DATE'])
    date = date.replace('/','-')
    df['Report Date'] = pd.to_datetime(date)
    
    display(df)
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )

    with pd.ExcelWriter('{}//MGT DTL DRWN F_STCK - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
            df.to_excel(writer, sheet_name='Raw Data',index=None)
    
# --------------------------------------------------------------- MPC ACCT PAYABLE --------------------------------------------------------

def mpcacctpayable(filepath):
    #filepath = 'Files/MPC ACCT PAYABLE.TXT'
    temp=[]

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])
        while line:
            linecount = linecount + 1
            line = fp.readline()
            #print(line.strip())
            if(len(line)):
                if (line[:5].strip() != '') and not(re.search('PAGE ',line)):
                    temp.append([line[0:5].strip(),line[5:11].strip(),line[11:16].strip(),line[17:20].strip(),line[20:26].strip(),
                                 line[25:29].strip(),line[29:49].strip(),line[50:56].strip(),line[57:65].strip(),line[66:83].strip(),
                                 line[83:91].strip(),line[94:105].strip(),line[105:120].strip(),line[121:].strip()])

    df= pd.DataFrame(temp,columns=['COMP','MGMT CHECK CTR', 'ACCT', 'FUNC' ,'COST CTR', 'DR CR', 'AMOUNT', 'JE#', 'VENDOR NAME', 'INVOICE', 'PO NUM', 'INV DATE', 'CHECK NUMBER', 'CHECK DATE'])
    #df = df[(df['COMP'].str.isnumeric()) & ((df['COMP'].str.len())==4) & (~(df['COMP'].str.contains('/'))) & ((df['INV DATE'].str.contains('/')))]
    df = df[df['COMP'].str.match('^[0-9]*$')]
    df['CHECK DATE'] = pd.to_datetime(df['CHECK DATE'],infer_datetime_format=True)
    df['INV DATE'] = pd.to_datetime(df['INV DATE'],infer_datetime_format=True)
    df['AMOUNT'] = pd.to_numeric(df['AMOUNT'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)
    df['Report Date'] = date

    display(df)
    date = str(date.date())

    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    
    with pd.ExcelWriter('{}//MPC ACCT PAYABLE - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
    
# ------------------------------------------------------- GMP 11 EXT MGCNT COMP-------------------------------------------------------------


def gmp11extmgcnt_comp(filepath):
    #filepath = 'Files/GMP11 EXT MGCNT_COMP.TXT'

    temp = []
    lc = 1000

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])

        while line:
            #linecount = linecount+1
            line = fp.readline()
            if line[:39].strip().isnumeric():
                temp.append(line.strip().split())
            #temp.append([line[:5].strip()]+[line[8:62].strip()]+[line[62:74].strip()]+[line[74:].strip()])

    df = pd.DataFrame(temp,columns=['MGMT CNTR','COMP'])
    df['Report Date'] = date

    display(df)
    date = str(date.date())
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )


    with pd.ExcelWriter('{}//GMP11 EXT MGCNT_COMP REPORT - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)


# ------------------------------------------------------- GMP 11 EXT COMP MGCNT-------------------------------------------------------------

def gmp11extcomp_mgcnt(filepath):
    #filepath = 'Files/GMP11 EXT COMP_MGCNT.TXT'

    temp = []
    lc = 1000

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])

        while line:
            #linecount = linecount+1
            line = fp.readline()
            if line[:39].strip().isnumeric():
                temp.append(line.strip().split())
            #temp.append([line[:5].strip()]+[line[8:62].strip()]+[line[62:74].strip()]+[line[74:].strip()])

    df = pd.DataFrame(temp,columns=['MGMT CNTR','COMP'])
    df['Report Date'] = date
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    date = str(date.date())

    display(df)
    
    with pd.ExcelWriter('{}//GMP11 EXT COMP_MGCNT REPORT - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
    
    
# -----------------------------------------------------GL MENU------------------------------------------------------------------------


def glmenu01(filepath):
    #filepath = 'Files/GL MENU01.TXT'

    temp=[]
    newtemp = []

    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[-1])

        while line:
            linecount = linecount + 1
            line = fp.readline()
            newline = line.strip()
            if(len(newline)):
                #print(newline)
                if re.search('^[0-9]*$',newline[0:6]):
                    temp.append([newline[:13].strip(),newline[13:43].strip(),newline[43:].strip()])
                    #temp.append(newline)
                elif newline[:14]=='TOTAL ACCOUNTS':
                    temp.append([newline[:14].strip(),newline[15:26].strip(),newline[26:39].strip(),newline[40:].strip()])

    df = pd.DataFrame(temp,columns=['ACCOUNT NO.','ACCOUNT DESCRIPTION','COST CENTERS',''])
    df['Report Date'] = date

    display(df)
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    date = str(date.date())

    
    with pd.ExcelWriter('{}//GL MENU01- {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
# ------------------------------------------------- FINAL HIER ROLLUP--------------------------------------------------------------

def finalhierrollup(filepath):
    #filepath = 'Files/FINAL HIER ROLLUP.TXT'

    temp=[]
    lc=1000
    with open(filepath) as fp:
        linecount = 0    
        line = fp.readline()
        date = pd.to_datetime(line.split()[0])
        while line:
            line = fp.readline()
            newline = line.strip()
            #print(newline)
            if len(newline): 
                if newline[:2]=='..':

                    temp.append([newline[:15].strip()]+[newline[16:58].strip()]+[newline[59:64].strip()]+
                                [newline[65:68].strip()]+[newline[69:75].strip()]+[newline[76:82].strip()]+
                                [newline[83:89].strip()]+[newline[90:].strip()])
                    linecount = linecount + 1

    df = pd.DataFrame(temp,columns=['HIERARCHY','DESCRIPTION','CTR','1100','2100','2200','2300','3100'])
    df['HIERARCHY'] = df['HIERARCHY'].str.replace('.','0')
    df['HIERARCHY'] = df['HIERARCHY'].astype('int64')
    df['Report Date'] = date

    display(df)
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1])
    
    date = str(date.date())


    with pd.ExcelWriter('{}//FINAL HIER ROLLUP - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)

#---------------------------------------------------- CTD REF OT HRS-----------------------------------------------------------------------

def ctdreg_othrswages(filepath):
    
    #filepath = 'Files/CTD REG_OT HRS WAGES.txt'
    headers = []
    temp=[]

    with open(filepath) as fp:
        linecount = 0

        line = fp.readline()
        date = pd.to_datetime(line.split()[0])

        while line:
            line = fp.readline()
            newline = line.strip().split()
            if len(newline):
                #print(newline)
                if (newline[0]=='LVL2'):
                    #header = newline
                    linecount = linecount+1
                    temp.append(newline)
                    #print(newline)
                elif (re.search('^[0-9]*$',newline[0])):
                    #print(newline)
                    linecount = linecount+1
                    temp.append(newline)

    df = pd.DataFrame(temp[1:])
    df['Report Date'] = date

    display(df)

    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    date = str(date.date())

    with pd.ExcelWriter('{}//CTD REG_OT HRS WAGES REPORT - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
    
#------------------------------------------------- BATCH PROOF --------------------------------------------------------------


def batchproof(filepath):
    #filepath = 'Files/BATCH PROOF.txt'
    temp = list()
    #filepath='Files/BATCH PROOF.TXT'
    flag=0

    with open(filepath) as fp:
        linecount = 0
        line = fp.readline()
        while line:
            line = fp.readline()
            newline = line.strip()
            if len(newline):

                if re.findall('(DATE: )((0|1)\d{1})/((0|1|2)\d{1})/((19|20)\d{2})',line) and flag==0:
                    date = (re.search('(DATE: )((0|1)\d{1})/((0|1|2)\d{1})/((19|20)\d{2})',line)[0].split()[-1])
                    flag = 1
                if newline[:2] == 'PT':
                    temp.append(newline[:71].split()+[newline[79:89]]+[newline[98:107]]+[newline[108:116]]+[newline[117:]])
                    linecount = linecount+1

                if newline[:2] == '**':
                    newline = newline.split()
                    if len(newline)==15:
                        temp.append([None]+[None]+[None]+[None]+[None]+[None]+['Subtotal for Source']+[newline[6]]+[None]+[newline[10]]+[newline[14]]+[None]+[None])

    df = pd.DataFrame(temp, columns=['TC','ITEM','1','DC','COMP','ACCOUNT','CENTER','SOURCE','DATE','DR_AMOUNT','CR_AMOUNT','OPER ID','COMMENTS'])

    df['DATE'] = pd.to_datetime(df['DATE'],format='%m%d%Y')
    df['DR_AMOUNT'] = pd.to_numeric(df['DR_AMOUNT'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)
    df['CR_AMOUNT'] = pd.to_numeric(df['CR_AMOUNT'].replace(',','',regex=True), errors='coerce').fillna(0).astype(float)
    df['Report Date'] = pd.to_datetime(date)

    display(df)   

    date = date.replace('/','-')

    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )

    with pd.ExcelWriter('{}//BATCH PROOF - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)

#------------------------------------------------- CHART OF ACCOUNTS--------------------------------------------------------------

def chartofaccounts(filepath):
    #filepath = 'Files/CHART OF ACCOUNTS (3).txt'
    filepath = 'Files/CHART OF ACCOUNTS (3).txt'
    with open(filepath) as fp:
    # Read one line at a time
        temp = {}
        key = 0
        temp.setdefault(key, [])
        flag=0
        line = fp.readline()
        count = 0
        while line:
            line = fp.readline()
            newline = line.strip()
            count = count+1

            if len(newline):
                if re.findall('(ISSUED )((0|1)\d{1})/((0|1|2)\d{1})/(\d+)',line) and flag==0:
                    date = re.search('(ISSUED )((0|1)\d{1})/((0|1|2)\d{1})/(\d+)',line)[0].split()[-1]
                    flag = 1

                if(re.match('[0-9]|T|B', newline[:1])):
                    if(newline[:1]=='B'):
                        key = key+1
                        temp[key] = [newline[:15].strip()]
                        temp[key].append(newline[16:20])
                        temp[key].append(newline[21:41])
                        temp[key] = temp[key] + newline[41:58].split()
                        temp[key].append(newline[58:77])
                        temp[key] = temp[key] + newline[78:].split()

                    #count = count+1
                    if(newline[:3]=='TOT'):
                        temp[key].append(newline[:18])
                        temp[key].append(newline[19:23])
                        temp[key] = temp[key] + newline[23:].split()

                    if(re.match('[0-9]',newline[:1])):
                        temp[key].append(newline.split())

    df = pd.DataFrame.from_dict(temp,orient='index')
    df = df.drop(df.index[0])

    df['Active centers'] = df.loc(axis=1)[7:].fillna('').apply(lambda x: ''.join(str(x)),axis=1)
    df = df[[0,1,2,3,4,5,6,'Active centers']]
    df = df.rename(columns={0:'CT',1:'ACCT NO',2:'ACCOUNT DESCRIPTION',3:'CL',4:'GR',5:'NS',6:'NF'})

    date = date.replace('/','-')

    display(df)
    outputfile = '//'.join(str(s) for s in filepath.split('/')[:-1] )
    
    with pd.ExcelWriter('{}//CHART OF ACCOUNTS - {}.xlsx'.format(outputfile,date)) as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Raw Data',index=None)
