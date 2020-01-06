import docx2txt
from summa import keywords
import os
import re
import urllib.request
import time
from bs4 import BeautifulSoup
print ("------------------------------------------------------------")
print ("Auto Keywords V0.5 By kinoko")
print ("------------------------------------------------------------")
def Read_Dictionary(File_Path):
    Temp = open(File_Path,'r')
    Temp_Dictionary = Temp.readlines()
    Dic = []
    for each in Temp_Dictionary:
        TempElement = each[0:-1]
        Dic.append(TempElement)
    return Dic

def Get_Numbers(Input_Html):
    TempT = str(Input_Html)
    Strong1 = TempT.find("<strong>")
    Strong2 = TempT.find("</strong>")
    RawNumber = re.sub("<strong>","",TempT[Strong1-1:Strong2])
    FinalNumber =0
    if RawNumber.find(",")>-1:
        FinalNumber = int(re.sub(",","",RawNumber))
    else:
        FinalNumber = int(RawNumber)
    return FinalNumber

def Search_Hits(queryStr):
    if queryStr.find(" ")>-1:
        TempStr = queryStr.split()
        TempStr2 = ""
        for each in TempStr:
            TempStr2 = TempStr2 + each + "+"
        InStrFinal = str(TempStr2[0:-1])
    else:
        InStrFinal = queryStr   
    url = 'https://link.springer.com/search?query=' + InStrFinal
    Retry = 1
    html_code = "None"
    Hits = 0
    while Retry ==1:
        print ("Pending Communication to Springer")
        j = 1
        while j <11:
                time.sleep(1)
                print("Connect Springer in %d second(s)" %(11-j))
                j+= 1
        try:
            headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11'}
            opener = urllib.request.build_opener()
            opener.addheaders = [headers]
            print ("Connecting Springer, Keyword is", InStrFinal)
            request = urllib.request.Request(url)
            request.add_header("user-agent","Mozilla/0.5")
            response_obj = urllib.request.urlopen(request)
            html_code = response_obj.read().decode('utf-8')
            soup = BeautifulSoup(html_code,'lxml')
            CountsDetected1 = soup.findAll(name='h1', attrs={"id":"number-of-search-results-and-search-terms"})
            CountsDetected = Get_Numbers(CountsDetected1)
            print("Successful, %d Hits" %(CountsDetected))
            Retry =0
            Hits = CountsDetected
            break
        except Exception as e:
            print("Failed")
            Retry =1
            i = 1
            while i <11:
                time.sleep(1)
                print("Retry in %d second(s)" %(11-i))
                i+= 1
            continue
    return Hits


CustomizedDictionary= Read_Dictionary("E:\Customized Dictionary\\test.txt")

PathPrefix = "E:\Test Manuscript Pool\\"
TempFileList = os.listdir("E:\Test Manuscript Pool")
TempFileList2 = []
for each in TempFileList:
    TruePath = PathPrefix+ each
    TempFileList2.append(TruePath)



for eachpath in TempFileList2:
    print("----------------------")
    print(eachpath)
    print("Keywords:")
    print("----------------------")
    text = docx2txt.process(eachpath)
    FilePath = eachpath
    Temp1 = keywords.keywords(text)
    Temp2 = Temp1.split("\n")
    Temp3 = Temp2[0:49]
    for each in Temp3:
        Length = len(each)
        if Length == 1:
            Temp3.remove(each)
        if each in CustomizedDictionary:
            Temp3.remove(each)
    print(Temp3)
    print("----------------------")
    print("Searching Springer for Possible Frequences")
    Temp4=[]
    for each in Temp3:
        TempResult = Search_Hits(each)
        Temp4.append(TempResult)
    DictForThisCase ={}
    i = 0
    for each in Temp3:
        DictForThisCase[each]=Temp4[i]
        i+=1
    DictAscending1 = sorted(DictForThisCase.items(),key=lambda x:x[1])
    DictAscending=[]
    for each in DictAscending1:
        Reform = each[0]
        DictAscending.append(Reform)
    print("----------------------")
    print("Current Result for",FilePath)
    print(DictAscending)
    print("Exporting Data To Txt")
    FileName1 = FilePath.split("\\")
    FileName2 = FileName1[-1]
    FileName3 = FileName2.split()
    FileName = FileName3[0]
    LogPath=  "E:\Test Manuscript Pool\\"+FileName+"-Result.txt"
    Log= open(LogPath,'w')
    print(FilePath,file=Log)
    print("---------------------------------------",file=Log)
    print("TextRank",file=Log)
    print(Temp3,file=Log)
    print("---------------------------------------",file=Log)
    print("Adjusted by Springer",file=Log)
    print(DictAscending,file=Log)
    print("---------------------------------------",file=Log)
    Log.close()
    print("Exported!")
    