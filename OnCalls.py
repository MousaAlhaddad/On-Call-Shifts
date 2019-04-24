import pandas as pd
import numpy as np

file = 'OnCalls.xlsx'
xl = pd.ExcelFile(file)
xlData = xl.parse(0)
np.random.seed(42)

Names=[]
Days=[]

nP = 8
nD = 30
n = nD + nP - (nD % nP)
m = n / nP

for x in range(len(xlData)):
    for y in range(1,n+1):
        Excluded = False 
        for z in range(1,4):
            if y == xlData.iloc[x,z]:
                Excluded = True
        if Excluded == False:
                Names.append(xlData.Name[x])
                Days.append(y)

list_cols=[Days,Names]
list_labels=["Day","Name"]
zipped= list(zip(list_labels,list_cols))
data=dict(zipped)
df = pd.DataFrame(data).sort_values('Day').reset_index(drop=True)

schedule = pd.DataFrame()
for x in range(1,n+1):
    DayCounts = df.Day.value_counts()
    minDays = DayCounts[DayCounts == min(DayCounts)].index.tolist()
    d= int(np.random.random()*len(minDays))
    d = minDays[d]
    sdf = df[df.Day == d].reset_index(drop=True)
    selectedDay = df[df.Day == d].index.tolist()
    r= int(np.random.random()*len(sdf))
    schedule = schedule.append(sdf.iloc[r])
    df = df[(df.Day != d)]
    df = df[(df.Name != sdf.iloc[r].Name) | (df.Day != d-1)]
    df = df[(df.Name != sdf.iloc[r].Name) | (df.Day != d+1)]
    df = df[(df.Name != sdf.iloc[r].Name) | (df.Day != d-2)]
    df = df[(df.Name != sdf.iloc[r].Name) | (df.Day != d+2)]
    p = (d - ( d % nP)) / nP + 1
    if d % nP == 0:
        p = p - 1
    for y in range(int(nP*(p-1)+1),int(nP*(p)+1)):
        df = df[(df.Name != sdf.iloc[r].Name) | (df.Day != y)]    
    NameCounts = schedule.Name.value_counts()
    for y in range(len(NameCounts)):
        if NameCounts.iloc[y] == m:
            df = df[(df.Name != NameCounts.index[y])]
    
schedule = schedule.sort_values('Day').reset_index(drop=True)   

schedule = schedule.iloc[:nD]

print(schedule)
print ()
print(schedule.Name.value_counts())
print ()
for x in schedule.Name.unique():
    print (x)
    print (schedule.Day[schedule.Name == x].unique())

schedule = schedule.set_index('Day')
writer = pd.ExcelWriter('On-Call Schedule.xlsx')
schedule.to_excel(writer,'Sheet1')
schedule.Name.value_counts().to_excel(writer,'Sheet2')
writer.save()
