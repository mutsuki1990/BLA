import docx2txt
from summa import keywords
import os
import re
import urllib.request
import time
from bs4 import BeautifulSoup
import requests
import datetime
print ("------------------------------------------------------------")
print ("Auto Keywords V0.7 By kinoko")
print ("------------------------------------------------------------")
def Read_Dictionary(File_Path):
    Temp = open(File_Path,'r')
    Temp_Dictionary = Temp.readlines()
    Dic = []
    for each in Temp_Dictionary:
        TempElement = each[0:-1]
        Dic.append(TempElement)
    return Dic

def Notification(Title,Content):
    api = "https://sc.ftqq.com/SCU76107Tca7efaf0f8233d327b6feb4e9582ab295e128e5cc0886.send"
    NotificationTitle = Title
    NotificationContent = Content
    data = {
            "text":NotificationTitle,
            "desp":NotificationContent
            }
    requests.post(api, params=data, timeout=10)
    print("Notification Sent to Wechat")
    Msg = "S"
    return Msg

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
    RetryTime = 0
    while Retry ==1:
        print ("Pending Communication to Springer")
        j = 1
        while j <11:
                time.sleep(1)
                print("Connect Springer in %d second(s)" %(11-j))
                j+= 1
        try:
            print ("Connecting Springer, Keyword is", InStrFinal)
            TimeBefore = datetime.datetime.now()
            request = urllib.request.Request(url)
            request.add_header("user-agent","Mozilla/0.5")
            response_obj = urllib.request.urlopen(request,timeout = 60)
            ConnectionStatus = response_obj.getcode()
            print(ConnectionStatus)
            html_code = response_obj.read().decode('utf-8')
            soup = BeautifulSoup(html_code,'lxml')
            CountsDetected1 = soup.findAll(name='h1', attrs={"id":"number-of-search-results-and-search-terms"})
            CountsDetected = Get_Numbers(CountsDetected1)
            Timeafter = datetime.datetime.now()
            TotalTimeConsumed = Timeafter - TimeBefore
            print("Successful, %d Hits, Time consumed %d second(s)" %(CountsDetected,TotalTimeConsumed.seconds))
            Retry =0
            Hits = CountsDetected
            break
        except Exception as e:
            print("Failed")
            Retry =1
            i = 1
            j = 1
            RetryTime +=1
            print("Retried Time", RetryTime)
            if RetryTime <4:
                while i <11:
                    time.sleep(1)
                    print("Retry in %d second(s)" %(11-i))
                    i+= 1
                continue
            else:
                print("Program turn to sleep mode, reactivate in 300s")
                X = Notification("Time out too many", "Turn into SleepMode")
                while j <11:
                    time.sleep(30)
                    print("Retry in %d second(s)" %(30*(11-j)))
                    j+= 1
                Y = Notification("Restarting", "Restarting")
                print("Program Restarting")
                RetryTime = 0
                continue
            
    return Hits


CustomizedDictionary= Read_Dictionary("E:\Customized Dictionary\\test.txt")

PathPrefix = "E:\Test Manuscript Pool\\"
TempFileList = os.listdir("E:\Test Manuscript Pool")
TempFileList2 = []
for each in TempFileList:
    TruePath = PathPrefix+ each
    TempFileList2.append(TruePath)

WholeWorkList = str(len(TempFileList2)) + "case(s) in total"
Not1 = Notification("AutoKeywords 0.5 started",WholeWorkList)

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

Z= Notification("All Done", "All Done")
    