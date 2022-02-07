# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import os
import openpyxl
import requests



def isDuplicatedCode(list, code) :
    isDuplicated = False
    for j in range(0, len(list)):
        obj = list[j]

        if (obj["business_code"] == code):
            isDuplicated = True
            break
    return isDuplicated


def getICTCodeFromFile():
    businessCodeList = []
    df = pd.read_excel("./ict_code.xlsx", usecols=[4,5,6,7,8,9,10,11,12,13,14], skiprows=[0,1,2,3], skipfooter=4)
    for i in range(0, len(df["2019년 귀속(개편후) 업종코드"])):
        if not isDuplicatedCode(businessCodeList, df["2019년 귀속(개편후) 업종코드"][i]) :

            businessCodeInfo = {
                "business_code": df["2019년 귀속(개편후) 업종코드"][i],
                "large_group_code" : df["대분류(Code)"][i],
                "large_group_name" : df["대분류"][i],
                "medium_group_code" : df["중분류(Code)"][i],
                "medium_group_name" : df["중분류"][i],
                "small_group_code" : df["세분류(Code)"][i],
                "small_group_name" : df["세분류"][i],
                "x_small_group_name" : df["세세분류"][i],
            }

            businessCodeList.append(businessCodeInfo)

    return businessCodeList

if __name__ == '__main__':
    businessCodes = getICTCodeFromFile()
    print(businessCodes)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
