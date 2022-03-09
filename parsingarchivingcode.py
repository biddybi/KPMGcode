import os
import pandas as pd
rootdir = os.getcwd()

cl0=pd.DataFrame()
cl1=pd.DataFrame()
a=[]
b=[]
        
def is_contains_chinese(strs):
    for _char in strs:
        if '\u4e00' <= _char <= '\u9fa5':
            return True
    return False

                
def read_file_rule1(path):
    global cl0, cl1,a,b
    require_cols = [2, 3]
    df = pd.read_excel(path, usecols = require_cols)
    corp= str((path.split("\\")[-1]).split("_")[1])
   
    
    for i in df.index:
        rows=df.iloc[[i]]
        if type(df['Unnamed: 2'][i])==str and is_contains_chinese(df['Unnamed: 2'][i])==False:
            if df['Unnamed: 2'][i][0]=='A'and df['Unnamed: 2'][i][1]=='_':
                if len(df['Unnamed: 2'][i])==20 and df['Unnamed: 2'][i] !='A_91440101231245152Y':
                    cl0=cl0.append(rows,ignore_index=False)
                    a.append(corp)
                else:
                    if df['Unnamed: 2'][i] !='A_91440101231245152Y':
                        cl1=cl1.append(rows,ignore_index=False)
                        b.append(corp)
            elif df['Unnamed: 2'][i][0]=='E'and df['Unnamed: 2'][i][1]=='_':
                if len(df['Unnamed: 2'][i])==13 and df['Unnamed: 2'][i][10]=='0'and df['Unnamed: 2'][i][11]=='0'and df['Unnamed: 2'][i][12]=='0':
                    cl0=cl0.append(rows,ignore_index=False)
                    a.append(corp)
                else:
                    if df['Unnamed: 2'][i] !='E_88888888001':
                        cl1=cl1.append(rows,ignore_index=False)
                        b.append(corp)
            elif df['Unnamed: 2'][i][0]=='F'and df['Unnamed: 2'][i][1]=='_':
                if len(df['Unnamed: 2'][i])==18:
                    cl0=cl0.append(rows,ignore_index=False)
                    a.append(corp)
                else:
                    if df['Unnamed: 2'][i] !='F_111111111110001':	
                        cl1=cl1.append(rows,ignore_index=False)
                        b.append(corp)
            else:
                cl1=cl1.append(rows,ignore_index=False)
                b.append(corp)
    
    

if __name__ == "__main__":
    for tmpDir in filter(os.path.isdir, os.listdir(rootdir)):  #遍历与py脚本相同的directory, 只遍历folder
        print('Reading folder: ' + tmpDir)   #打印folder
        maxDate = -1
        for file in os.listdir(tmpDir):
            if "客商_匹配表" in file:
                maxDate = max(int(file.split('\\')[-1].split('_')[0]), maxDate)
        for file in os.listdir(tmpDir):
            if str(maxDate) in file:
                if "客商_匹配表" in file:
                    full_path = os.path.join(tmpDir,file)
                    print('Newest File path is: ' + full_path)
                    read_file_rule1(full_path)
    cl0['corp']=a
    cl1['corp']=b
    cl0.to_excel('rightkeshang.xlsx')
    cl1.to_excel('wrongkeshang.xlsx')
    
    
