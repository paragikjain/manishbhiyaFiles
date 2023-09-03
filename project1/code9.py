import pandas as pd
import json
import math
import re
import sys

def findWorkCode(inputstr):
    pattern = r'\(([^()]*\d+\/[^()]*)\)'
    match = re.search(pattern, inputstr)
    if match:
        value_in_parentheses = match.group(1)
        return value_in_parentheses
    else:
        return "NOT FOUND"



def iterate_excel_rows(file_path, sheet_number,WorkNameDict):
    df = pd.read_excel(file_path, sheet_name=sheet_number)  # Assuming sheet index 0 corresponds to Sheet 1
    currWork = ""
    if(sheet_number >= 3):
        sheet_number = sheet_number-3
    if(sheet_number >= 3):
        sheet_number = sheet_number-3
    mode = 1
    for index, row in df.iterrows():

        if pd.isna(row[0]):
            mode = 2
            continue
        if mode == 1:
            if(type(row[0]) is not int):
                if row[0] not in WorkNameDict:
                    WorkNameDict[row[0]] = {}
                currWork = row[0]
            else:
                    if row[0] not in WorkNameDict[currWork]:
                        WorkNameDict[currWork][row[0]] = [0,0,0,0]
                    WorkNameDict[currWork][row[0]][0] = WorkNameDict[currWork][row[0]][0]+1
                    WorkNameDict[currWork][row[0]][sheet_number+1] = WorkNameDict[currWork][row[0]][sheet_number+1]+ row[1]
        else:
            currrow = "Work Name: " + row[2]+"("+row[3]+") "
            if currrow not in WorkNameDict:
                WorkNameDict[currrow] = {}
            if row[0] not in WorkNameDict[currrow]:
                WorkNameDict[currrow][row[0]] = [0,0,0,0]
            WorkNameDict[currrow][row[0]][0] = WorkNameDict[currrow][row[0]][0]+1
            WorkNameDict[currrow][row[0]][sheet_number+1] = WorkNameDict[currrow][row[0]][sheet_number+1]+ row[1]



def iterate_excel_rows_way2(file_path, sheet_number,WorkNameDict):
    df = pd.read_excel(file_path, sheet_name=sheet_number)  # Assuming sheet index 0 corresponds to Sheet 1
    currWork = ""
    for index, row in df.iterrows():
        if pd.notna(row[1]):
            currWork = str(row[1])
        if currWork not in WorkNameDict:
            WorkNameDict[currWork] = {"sheet3" : [0,0,0,row[2]]}
        else:
            WorkNameDict[currWork]["sheet3"][3] = WorkNameDict[currWork]["sheet3"][3] + row[2]

if len(sys.argv) > 1:
    argument = sys.argv[1]
else:
    print("Please make sure you have provided file path")
# Provide the path to your Excel file
excel_file_path = argument
WorkNameDict = {}
WorkNameDict2 = {}
WorkNameDict3 = {}
iterate_excel_rows(excel_file_path,0,WorkNameDict)
iterate_excel_rows(excel_file_path,1,WorkNameDict)
iterate_excel_rows_way2(excel_file_path,2,WorkNameDict)

iterate_excel_rows(excel_file_path,3,WorkNameDict2)
iterate_excel_rows(excel_file_path,4,WorkNameDict2)
iterate_excel_rows_way2(excel_file_path,5,WorkNameDict2)

iterate_excel_rows(excel_file_path,6,WorkNameDict3)
iterate_excel_rows(excel_file_path,7,WorkNameDict3)
iterate_excel_rows_way2(excel_file_path,8,WorkNameDict3)

# print(json.dumps(WorkNameDict, indent=4))

finalData = []
column = ["Muster Roll No", "WorkCode","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19"]

#iterate the dictionary and fill the results to pandas dataframe
for key, value in WorkNameDict.items():
    count = 0
    sum0 = 0
    sum1 = 0
    sum2 = 0

    count2 = 0
    sum3 = 0
    sum4 = 0
    sum5 = 0

    count3 = 0
    sum6 = 0
    sum7 = 0
    sum8 = 0

    for key1, list in value.items():
        if key1 != "sheet3":
            count =  count+1
        sum0 = sum0 + list[1]
        sum1 = sum1 + list[2]
        sum2 = sum2 + list[3]

    #Revalidate entries which have workcode in their name
    matchresult = findWorkCode(str(key))
    if(matchresult != "NOT FOUND"):
        finalData.append({'Muster Roll No': key,'WorkCode': matchresult,'5': count, '6': sum0, '7': sum1, '8': sum2, '9': sum1+sum2,'10': count2, '11': sum3, '12': sum4, '13': sum5, '14': sum4+sum5, '15': count3, '16': sum6, '17': sum7, '18': sum8, '19': sum7+sum8})


#iterate the dictionary and fill the results to pandas dataframe
for key, value in WorkNameDict2.items():
    count = 0
    sum0 = 0
    sum1 = 0
    sum2 = 0

    count2 = 0
    sum3 = 0
    sum4 = 0
    sum5 = 0

    count3 = 0
    sum6 = 0
    sum7 = 0
    sum8 = 0

    for key1, list in value.items():
        if key1 != "sheet3":
            count2 =  count2+1
        sum3 = sum3 + list[1]
        sum4 = sum4 + list[2]
        sum5 = sum5 + list[3]

    #Revalidate entries which have workcode in their name
    matchresult = findWorkCode(str(key))
    if(matchresult != "NOT FOUND"):
        finalData.append({'Muster Roll No': key,'WorkCode': matchresult,'5': count, '6': sum0, '7': sum1, '8': sum2, '9': sum1+sum2,'10': count2, '11': sum3, '12': sum4, '13': sum5, '14': sum4+sum5, '15': count3, '16': sum6, '17': sum7, '18': sum8, '19': sum7+sum8})


#iterate the dictionary and fill the results to pandas dataframe
for key, value in WorkNameDict3.items():
    count = 0
    sum0 = 0
    sum1 = 0
    sum2 = 0

    count2 = 0
    sum3 = 0
    sum4 = 0
    sum5 = 0

    count3 = 0
    sum6 = 0
    sum7 = 0
    sum8 = 0

    for key1, list in value.items():
        if key1 != "sheet3":
            count3 =  count3+1
        sum6 = sum6 + list[1]
        sum7 = sum7 + list[2]
        sum8 = sum8 + list[3]

    #Revalidate entries which have workcode in their name
    matchresult = findWorkCode(str(key))
    if(matchresult != "NOT FOUND"):
        finalData.append({'Muster Roll No': key,'WorkCode': matchresult,'5': count, '6': sum0, '7': sum1, '8': sum2, '9': sum1+sum2,'10': count2, '11': sum3, '12': sum4, '13': sum5, '14': sum4+sum5, '15': count3, '16': sum6, '17': sum7, '18': sum8, '19': sum7+sum8})



#aggragte function in case two workcodes are in different row , below is the column wise aggreagtion scheme
aggregation_functions = {
    'Muster Roll No': 'first',
    '5': 'sum',
    '6': 'sum',
    '7': 'sum',
    '8': 'sum',
    '9': 'sum',
    '10': 'sum',
    '11': 'sum',
    '12': 'sum',
    '13': 'sum',
    '14': 'sum',
    '15': 'sum',
    '16': 'sum',
    '17': 'sum',
    '18': 'sum',
    '19': 'sum'
}


df = pd.DataFrame(finalData, columns=column)
df_merged = df.groupby('WorkCode').agg(aggregation_functions).reset_index()
df_merged.to_excel("final_"+excel_file_path, index=False)



