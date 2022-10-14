from pickle import TRUE
import pandas as pd
import numpy as np
from numpy import False_
from datetime import datetime,timedelta
import openpyxl


location=r'C:\Users\aditya.gupta\Downloads\Daily DPR_PKGC FINAL 09-Oct-22 1.xlsx'
location2=r'C:\Users\aditya.gupta\Downloads\PKG C AT Tracker 09-10-2022.xlsx'
df=pd.read_excel(location,sheet_name='Day Wise Progress',skiprows=1)

df1=pd.read_excel(location2,sheet_name='Master Sheet',skiprows=2)

print("Exectuted succesfully 1")


####################################################################################################################

 ##dpr start
df2=pd.DataFrame()
df2['GP LGD Code']=df[r'GP LGD Code']
df2['District']=df[r'District']
df2['Zone Name']=df[r'Zone Name']
df2['Mandal Name']=df[r'Mandal']

df2['GP Name']=df[r'Target GP']
df2['Date of Activity']=df[r' Date of Activity']
#df2['T&D Scope']=df1[r'Green Field (New T&D)-Kms']
df2['T&D']=df[r'T&D']
#df2["brown"]=df1[r'Brown Field (MB)-Kms']
#df2['missing']=df1[r'DRT Missing Section Length (Kms)']
#df2['DRT Scope'] = df1[r'Brown Field (MB)-Kms']+df1[r'DRT Missing Section Length (Kms)']
df2['DRT']=df[r'DRT']

df2['DIT']=df[r'DIT']
#df2['Blowing Scope']=df1[r'Blowing Scope- Kms']
df2['Blowing']=df[r'Blowing']

#df2.drop_duplicates(subset="GP LGD Code",keep=False,inplace=True)
print("Exectuted succesfully 2")
df2.dropna(inplace = True)  
print(df2.info())

##DPR ENDED.

#########################################################################################################################

##### AT TRACKER

df3=pd.DataFrame()
print("Exectuted succesfully 3")
print("========",df1.columns)
df3['GP LGD Code']=df1[r'Target GP LGD Code']
df3['Zone Name']=df1[r'Zone']
df3['District']=df1[r'District']
df3['Mandal Name']=df1[r'Mandal']
df3['Gp Lit up']=df1[r'Litup date']
print("Exectuted succesfully 4")

df3['Common Offer Date']=df1[r'Integrated offered Date']
df3['Common First visit Date']=df1[r'Integrated 1st Visit Date']
#df3['Common Second visit Date']=df1[r'Common Second visit  Date']
df3['Common PPs cleared Date']=df1[r'Integrated PP cleared']
df3['Common ATC 4 Date']=df1[r'4-ATC released']
df3['BOQ Date']=df1[r'BOQ']
df3['Document Sub Date']=df1[r'T fiber document submitted']
df3['Proforma Invoice Date']=df1[r'Pro forma Invoice Raised']
df3['Tax Invoice Date']=df1[r'Tax Invoice']
print("Exectuted succesfully 5")
df3.dropna(inplace = True)  

print(df3.info())
#df3.drop_duplicates(subset="GP LGD Code",keep=False,inplace=True)


#with pd.ExcelWriter(r'C:\Users\sri.krishna\Documents\Dashboard_folder\AT-Tracker22.xlsx') as writer:

 #df3.to_excel(writer,sheet_name="AT Tracker",index=True)

 ##############################################################################################################################



df4 = pd.merge(df2, df3, on='GP LGD Code', how="outer")

print("Exectuted succesfully 6")
print(df4.info())
df4.to_excel('final-merged-C.xlsx')

#with pd.ExcelWriter(r'C:\Users\sri.krishna\Documents\Dashboard_folder\final_merged-C.xlsx') as writer:
   
    # use to_excel function and specify the sheet_name and index
    # to store the dataframe in specified sheet

    #df3.to_excel(writer,sheet_name="AT Tracker",index=False)
 #   df4.to_excel(writer, sheet_name="Final", index=False)


  #  print("Exectuted succesfully")