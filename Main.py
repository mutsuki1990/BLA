import pandas as pd
import numpy as np
import re
import datetime
import time
from dateutil import rrule
from tkinter import *
import tkinter.filedialog

print("--------------------------------")
print("Working Tracking Plugin V0.61, Programmed by Kinoko")
print("--------------------------------")


def Remove_Null(Input_Array):
    Temp_Array = pd.DataFrame(Input_Array)
    Temp_Array.dropna(inplace=True)
    Temp_Array2 =np.array(Temp_Array)
    return Temp_Array2

def Primitive_Process(Input_Excel_Data):
    PrimitiveNoNull = Input_Excel_Data['Trigger'].notnull()
    LastRow = 0
    for Cells in PrimitiveNoNull:
        LastRow = LastRow +1
        if Cells == False:
            break
    TransferToNpRawData= np.array(TempRestoration)
    RawDataWithoutFirstRow = TransferToNpRawData[0:LastRow]
    TriggerIndex = RawDataWithoutFirstRow[:,1:2]
    TriggerWithoutNull = Remove_Null(TriggerIndex)
    CaseIndexHor=np.unique(TriggerWithoutNull)
    CaseIndexVer=CaseIndexHor.reshape(CaseIndexHor.shape[0],1)
    TriggerIndexHor = TriggerIndex.reshape(1,TriggerIndex.shape[0])
    CaseInitiated =["Case Name","Contacted Persons","Response from Them","Receiving Date","First Revision Date","First Revision Due Date","Second Revision Date","Second Revision Due Date"]
    CaseSorted = np.array(CaseInitiated)
    for element in CaseIndexHor:
        for cell in TriggerIndexHor:
            LabelIndex=list(np.where(element==cell)[0])
            FirstOneHit= LabelIndex[0]
            LastOneHit = LabelIndex[-1]
            TempSlicedTotal = RawDataWithoutFirstRow[FirstOneHit:LastOneHit]
            SlicedAddressA = Remove_Null(TempSlicedTotal[:,4:5])
            SlicedReviewingA = Remove_Null(TempSlicedTotal[:,5:6])
            SlicedReceivedDate = Remove_Null(TempSlicedTotal[:,8:9])
            Sliced1stRevisionStarting = Remove_Null(TempSlicedTotal[:,9:10])
            Sliced1stRevisionEnding = Remove_Null(TempSlicedTotal[:,10:11])
            Sliced2ndRevisionStarting = Remove_Null(TempSlicedTotal[:,11:12])
            Sliced2ndRevisionEnding = Remove_Null(TempSlicedTotal[:,12:13])
            SlicedAddressB =SlicedAddressA.reshape(1,SlicedAddressA.shape[0])
            SlicedAddress = SlicedAddressB[0]
            SlicedReviewingB = SlicedReviewingA.reshape(1,SlicedReviewingA.shape[0])
            SlicedReviewing = SlicedReviewingB[0]
            ThisLine = [element,SlicedAddress,SlicedReviewing,SlicedReceivedDate,Sliced1stRevisionStarting,Sliced1stRevisionEnding,Sliced2ndRevisionStarting,Sliced2ndRevisionEnding]
            CaseSorted = np.vstack([CaseSorted,ThisLine])
     
    print("All data loaded, %d case(s) detected" %(CaseIndexHor.shape[0]))
    print("------------------------------------------------------------")
    AllData = [CaseIndexHor,CaseIndexVer,CaseSorted]
    return AllData

def Get_Agreed_Time(Input_String):
    DateTemp = Input_String.split("-")
    DateAgreeFirstT1 = DateTemp[0].split()
    DateAgreeFirstT2 = DateAgreeFirstT1[-1].strip()
    DateAgreeFirst = Grab_Single_Date(DateAgreeFirstT2)
    DateAgreeLast = Grab_Single_Date(DateTemp[1].strip())
    AgreedDate = [DateAgreeFirst,DateAgreeLast]
    return AgreedDate

def Get_Remind_Time(Input_String):
    Temp1 = Input_String.upper()
    DateTemp = Temp1.split("RM")
    if len(DateTemp) == 4:
        FirstReminder = DateTemp[1].strip()
        SecondReminder = DateTemp[2].strip()
        ThirdReminder = DateTemp[3].strip()
        ReminderTime = [FirstReminder,SecondReminder,ThirdReminder,3]
    elif len(DateTemp) == 3:
        FirstReminder = DateTemp[1].strip()
        SecondReminder = DateTemp[2].strip()
        ReminderTime = [FirstReminder,SecondReminder,"Nothing",2]
    elif len(DateTemp) == 2:
        FirstReminder = DateTemp[1].strip()
        ReminderTime = [FirstReminder,"Nothing","Nothing",1]
    return ReminderTime

def Grab_Single_Date(Input_String):
    if Input_String.find(".")>-1:
        TempDate = Input_String.split(".")
        MonthAbout = TempDate[0]
        DayAbout = TempDate[1]
        Month = re.sub("[^0-9]", "", MonthAbout)
        Day =  re.sub("[^0-9]", "", DayAbout)
        Date = Month + "." +Day
    return Date

def Set_Date(Input_String):
    TempDate = Input_String.split(".")
    Month = int(TempDate[0])
    Day = int(TempDate[1])
    CurrentDate = datetime.date.today()
    CurrentYear = CurrentDate.year
    Choice1 = datetime.datetime(CurrentYear,Month,Day)
    Choice2 = datetime.datetime(CurrentYear-1,Month,Day)
    DateDiff = rrule.rrule(rrule.DAILY,dtstart=Choice1,until=CurrentDate).count()
    if DateDiff<0:
        #x
        DateSet = Choice2
    else:
        DateSet = Choice1  
    return DateSet
 
def Get_FirstWaveOperationTime_When_Reminded(Input_String):
    Temp1 = Input_String.upper()
    DateTemp = Temp1.split("RM")
    if DateTemp[0].find("REED")>-1:
        OperationTime = ["ARM",Get_Agreed_Time(DateTemp[0])]
    elif DateTemp[0].find("NVITE")>-1:
        if DateTemp[0].find("TART")>-1:
            SecondReviewStartTimeTemp = DateTemp[0].split("TART")
            TempChoice1 = SecondReviewStartTimeTemp[-1].strip()
            TempChoice2 = re.sub("2ND","",SecondReviewStartTimeTemp[0].strip())
            SecondReviewStart = Grab_Single_Date(TempChoice1)
            InviteDate = Grab_Single_Date(TempChoice2)
            SecondReview = [InviteDate,SecondReviewStart]
            OperationTime = ["2ndS",SecondReview]
        else:
            DateTempSecondInvited = DateTemp[0].split()
            SecondInvitedTime = DateTempSecondInvited[2]
            OperationTime = ["2nd",SecondInvitedTime]                                    
    return OperationTime

def Get_InvitationOperationTime_No_Reminder(Input_String):
    Temp1 = Input_String.upper()
    if Temp1.find("TART")>-1:
        Temp2 = Temp1.split("TART")
        InviteTemp = re.sub("2ND","",Temp2[0])
        StartTemp = Temp2[1]
        InviteDate = Grab_Single_Date(InviteTemp)
        StartDate = Grab_Single_Date(StartTemp)
        OperationTime = [InviteDate,StartDate]
    else:
        Temp2 = re.sub("2ND","",Temp1)
        InviteDate = Grab_Single_Date(Temp2)
        OperationTime = InviteDate
    return OperationTime

def Count_String(Input_String,Detect_String):
    WithoutPeriod = len(re.sub(Detect_String,"",Input_String))
    WithPeriod = len(Input_String)
    PeriodCount = (WithPeriod - WithoutPeriod)/len(Detect_String)
    return PeriodCount

def Strip_Comments_Time(Input_String):
    TempA =Input_String.upper()
    if TempA.find("NVI")>-1:
        TempB = TempA.split()
        OpinionTime = Grab_Single_Date(TempB[-1].strip())
    else:
        TempB = TempA.split()
        OpinionTime = Grab_Single_Date(TempB[-2].strip())
    return OpinionTime 



def Sort_Status(Input_String):
    ReviewerResponse =["22 Invalid", "Nothing","Nothing","N",22]
    ThisReviewerResponse = Input_String.upper()
    if ThisReviewerResponse.find("'")>-1:
        ThisReviewerResponse = re.sub("'","",ThisReviewerResponse)

    if ThisReviewerResponse.find("D")>-1:
        #x
        if ThisReviewerResponse.find("D")== 0:
            #x
            if ThisReviewerResponse.find("AC")<0:
                #x
                ReviewerResponse = ["1 No Acknowledgement","Nothing","Nothing","N",1]
            else:
                ReviewerResponse = ["2 Declined and Acknowledged","Nothing","Nothing","A",2]
        elif ThisReviewerResponse.find("GREED")>-1:
            if ThisReviewerResponse.find("RM")<0:
                #x
                ReviewerResponse = ["3 Agreed No Reminder",Get_Agreed_Time(ThisReviewerResponse),"Nothing","N",3]
            else:
                ReviewerResponse = ["4 Agreed and Reminded",Get_FirstWaveOperationTime_When_Reminded(ThisReviewerResponse),Get_Remind_Time(ThisReviewerResponse),"Y",4]
        elif ThisReviewerResponse.find("NVITE")>-1:
            #x
            if Count_String(ThisReviewerResponse,"2ND") == 2.0:
                #x
                if ThisReviewerResponse.find("CCE")>-1:
                    #x
                    ReviewerResponse = ["5 2nd Accepted", Strip_Comments_Time(ThisReviewerResponse), "Nothing","N",5]
                elif ThisReviewerResponse.find("EJE")>-1:
                    #x
                    ReviewerResponse = ["6 2nd Rejected", Strip_Comments_Time(ThisReviewerResponse), "Nothing","N",6]
                elif ThisReviewerResponse.find("AJO")>-1:
                        ReviewerResponse = ["7 2nd Major", Strip_Comments_Time(ThisReviewerResponse), "Nothing","N",7]
                elif ThisReviewerResponse.find("INO")>-1:
                        ReviewerResponse = ["8 2nd Minor", Strip_Comments_Time(ThisReviewerResponse), "Nothing","N",8]                       
            else: 
                #x
                if ThisReviewerResponse.find("TART")>-1:
                    #x
                    if ThisReviewerResponse.find("RM")>-1:
                        #x
                        ReviewerResponse = ["9 2nd Review Started and Reminded",Get_FirstWaveOperationTime_When_Reminded(ThisReviewerResponse),Get_Remind_Time(ThisReviewerResponse),"Y",9]
                    else:
                        ReviewerResponse = ["10 2nd Review and Started No Reminder Yet",Get_InvitationOperationTime_No_Reminder(ThisReviewerResponse),"Nothing","N",10]
                else:
                    #x
                    if ThisReviewerResponse.find("RM")>-1:
                        #x
                        ReviewerResponse = ["11 2nd Invited no Response but Reminded", Get_FirstWaveOperationTime_When_Reminded(ThisReviewerResponse), Get_Remind_Time(ThisReviewerResponse),"Y",11]
                    else:
                        #x
                        ReviewerResponse = ["12 2nd Invited no Response no Reminder", Get_InvitationOperationTime_No_Reminder(ThisReviewerResponse),"Nothing","N",12 ]
        elif ThisReviewerResponse.find("DELI")>-1:
            #x
            ReviewerResponse = ["13 Undelivered","Nothing","Nothing","N",13]
        elif ThisReviewerResponse.find("ANCE")>-1:
            #x
            ReviewerResponse = ["14 Cancelled","Nothing","Nothing","N",14]
        elif ThisReviewerResponse.find("ECE")>-1:
            #x
            if ThisReviewerResponse.find("CCE")>-1:
                #x
                ReviewerResponse = ["15 1st Accepted", Strip_Comments_Time(ThisReviewerResponse), "Nothing","N",15]
            elif ThisReviewerResponse.find("EJE")>-1:
                ReviewerResponse = ["16 1st Rejected", Strip_Comments_Time(ThisReviewerResponse), "Nothing","N",16]
            elif ThisReviewerResponse.find("AJO")>-1:
                ReviewerResponse = ["17 1st Major", Strip_Comments_Time(ThisReviewerResponse), "Nothing","N",17]
            elif ThisReviewerResponse.find("INO")>-1:
                ReviewerResponse = ["18 1st Minor", Strip_Comments_Time(ThisReviewerResponse), "Nothing","N",18] 
    else:
        #x
        if Count_String(ThisReviewerResponse,"/") == 0.0:
            #x
            if not ThisReviewerResponse.isalnum():
                if ThisReviewerResponse.find(".")>-1:
                    ReviewerResponse =["19 First Invitation", Grab_Single_Date(ThisReviewerResponse),"Nothing","N",19]
                else:
                    ReviewerResponse =["22 Invalid", "Nothing","Nothing","N",22]
            else:
                ReviewerResponse =["22 Invalid", "Nothing","Nothing","N",22]   
        elif Count_String(ThisReviewerResponse,"/") == 1.0:
            FirstInvitationTemp1 = ThisReviewerResponse.split("/")
            FirstInvitationDate = Grab_Single_Date(FirstInvitationTemp1[0])
            FirstReminderDate = Grab_Single_Date(FirstInvitationTemp1[1])
            ReviewerResponse =["20 Invitation with 1 Reminder", FirstInvitationDate,FirstReminderDate,"Y",20]
        elif Count_String(ThisReviewerResponse,"/") == 2.0:
            InvitationTemp2 = ThisReviewerResponse.split("/")
            FirstInvitationDate = Grab_Single_Date(InvitationTemp2[0])
            FirstReminderDate = Grab_Single_Date(InvitationTemp2[1])
            SecondReminderDate = Grab_Single_Date(InvitationTemp2[2])
            RemindDates = [FirstReminderDate,SecondReminderDate]
            ReviewerResponse =["21 Invitation with 2 Reminder", FirstInvitationDate,RemindDates,"Y",21]           
    return ReviewerResponse

def Sort_Response(Input_Array):
    TempStatus1 = Input_Array[4]
    if TempStatus1 == 22 or TempStatus1 == 2 or TempStatus1 == 13 or TempStatus1 == 14 or TempStatus1 == 21:
        SortedResponse = [1,"Response Terminated","Nothing","Nothing","Nothing"]
    if TempStatus1 == 1:
        SortedResponse = [2,"Need Acknowledgement","Nothing","Nothing","Nothing"]
    if TempStatus1 == 3:
        AgreedTimeLast = Set_Date(Input_Array[1][1])
        CurrentDate = datetime.date.today()
        DateDiff1 = rrule.rrule(rrule.DAILY,dtstart=AgreedTimeLast,until=CurrentDate).count()
        DateDiff2 = rrule.rrule(rrule.DAILY,dtstart=CurrentDate,until=AgreedTimeLast).count()
        if DateDiff1>0:
            SortedResponse = [3,"Agreed Need Reminder",AgreedTimeLast,CurrentDate,DateDiff1]
        elif DateDiff2<4:
            SortedResponse = [3,"Agreed Need Reminder",AgreedTimeLast,CurrentDate,-DateDiff2]
        else:
            SortedResponse = [8,"Don't Worry","Nothing","Nothing","Nothing"]
    if TempStatus1 == 4:
        ReminderData = Input_Array[2]
        ReminderCounts = ReminderData[3]
        if ReminderCounts <2:
            LastReminder = Set_Date(ReminderData[0])
            CurrentDate = datetime.date.today()
            DateDiff1 = rrule.rrule(rrule.DAILY,dtstart=LastReminder,until=CurrentDate).count()
            DateDiff2 = rrule.rrule(rrule.DAILY,dtstart=CurrentDate,until=LastReminder).count()
            if DateDiff1>2:
                #x
                SortedResponse = [3,"Agreed Need Reminder Again",LastReminder,CurrentDate,DateDiff1]
            else:
                SortedResponse = [8,"Don't Worry","Nothing","Nothing","Nothing"]
        else:
            SortedResponse = [0,"On Your Count","Nothing","Nothing","Nothing"]
    if TempStatus1 == 9:
        ReminderData = Input_Array[2]
        ReminderCounts = ReminderData[3]
        if ReminderCounts <2:
            LastReminder = Set_Date(ReminderData[0])
            CurrentDate = datetime.date.today()
            DateDiff1 = rrule.rrule(rrule.DAILY,dtstart=AgreedTimeLast,until=CurrentDate).count()
            DateDiff2 = rrule.rrule(rrule.DAILY,dtstart=CurrentDate,until=AgreedTimeLast).count()
            if DateDiff1>2:
                #x
                SortedResponse = [4,"2nd Review Need Reminder Again",LastReminder,CurrentDate,DateDiff1]
            elif DateDiff2<4:
                SortedResponse = [4,"2nd Review Need Reminder Again",LastReminder,CurrentDate,-DateDiff2]
            else:
                SortedResponse = [8,"Don't Worry","Nothing","Nothing","Nothing"]
        else:
            SortedResponse = [0,"On Your Count","Nothing","Nothing","Nothing"]
    if TempStatus1 == 10:
        OperationData = Input_Array[1]
        StartDate = Set_Date(OperationData[1])
        CurrentDate = datetime.date.today()
        DateDiff1 = rrule.rrule(rrule.DAILY,dtstart=StartDate,until=CurrentDate).count()
        DateDiff2 = rrule.rrule(rrule.DAILY,dtstart=CurrentDate,until=StartDate).count()
        if DateDiff1>0:
            SortedResponse = [4,"2nd Review Need Reminder",StartDate,CurrentDate,DateDiff1]
        elif DateDiff2<4:
            SortedResponse = [4,"2nd Review Need Reminder",StartDate,CurrentDate,-DateDiff2]
        else:
            SortedResponse = [8,"Don't Worry","Nothing","Nothing","Nothing"]
    if TempStatus1 == 11:
        OperationData = Input_Array[2]
        ReminderCounts = OperationData[3]
        if ReminderCounts <2:
            LastReminderDate = Set_Date(OperationData[0])
            CurrentDate = datetime.date.today()
            DateDiff1 = rrule.rrule(rrule.DAILY,dtstart=LastReminderDate,until=CurrentDate).count()
            DateDiff2 = rrule.rrule(rrule.DAILY,dtstart=CurrentDate,until=LastReminderDate).count()
            if DateDiff1>2:
                #x
                SortedResponse = [4,"2nd Review Need Reminder Again",LastReminderDate,CurrentDate,DateDiff1]
            elif DateDiff2<4:
                SortedResponse = [4,"2nd Review Need Reminder Again",LastReminderDate,CurrentDate,-DateDiff2]
            else:
                SortedResponse = [8,"Don't Worry","Nothing","Nothing","Nothing"]
        else:
            SortedResponse = [0,"On Your Count","Nothing","Nothing","Nothing"]
    if TempStatus1 == 12:
        SecondInvitationDate = Set_Date(Input_Array[1])
        CurrentDate = datetime.date.today()
        DateDiff1 = rrule.rrule(rrule.DAILY,dtstart=SecondInvitationDate,until=CurrentDate).count()
        DateDiff2 = rrule.rrule(rrule.DAILY,dtstart=CurrentDate,until=SecondInvitationDate).count()
        if DateDiff1>0:
            #x
            SortedResponse = [4,"2nd Review Need Reminder",LastReminder,CurrentDate,DateDiff1]
        elif DateDiff2<4:
            SortedResponse = [4,"2nd Review Need Reminder",LastReminder,CurrentDate,-DateDiff2]
        else:
            SortedResponse = [8,"Don't Worry","Nothing","Nothing","Nothing"]
    if TempStatus1 == 15:
        ReportReceivedDate = Set_Date(Input_Array[1])
        SortedResponse = [5,"1st Accept",ReportReceivedDate,"Nothing","Nothing"]
    if TempStatus1 == 16:
        ReportReceivedDate = Set_Date(Input_Array[1])
        SortedResponse = [5,"1st Reject",ReportReceivedDate,"Nothing","Nothing"]
    if TempStatus1 == 17:
        ReportReceivedDate = Set_Date(Input_Array[1])
        SortedResponse = [5,"1st Major",ReportReceivedDate,"Nothing","Nothing"]
    if TempStatus1 == 18:
        ReportReceivedDate = Set_Date(Input_Array[1])
        SortedResponse = [5,"1st Minor",ReportReceivedDate,"Nothing","Nothing"]
    if TempStatus1 == 19:
        FirstInvitationDate = Set_Date(Input_Array[1])
        CurrentDate = datetime.date.today()
        DateDiff1 = rrule.rrule(rrule.DAILY,dtstart=FirstInvitationDate,until=CurrentDate).count()
        DateDiff2 = rrule.rrule(rrule.DAILY,dtstart=CurrentDate,until=FirstInvitationDate).count()
        if DateDiff1>1:
            #x
            SortedResponse = [6,"1st Invite Need Reminder",FirstInvitationDate,CurrentDate,DateDiff1]
        else:
            SortedResponse = [9,"1st Good","Nothing","Nothing","Nothing"]
    if TempStatus1 == 20:
        ReminderDate = Set_Date(Input_Array[2])
        CurrentDate = datetime.date.today()
        DateDiff1 = rrule.rrule(rrule.DAILY,dtstart=ReminderDate,until=CurrentDate).count()
        DateDiff2 = rrule.rrule(rrule.DAILY,dtstart=CurrentDate,until=ReminderDate).count()
        if DateDiff1>2:
            #x
            SortedResponse = [6,"1st Invite Need Reminder Again",ReminderDate,CurrentDate,DateDiff1]
        else:
            SortedResponse = [9,"1st Good","Nothing","Nothing","Nothing"]
    if TempStatus1 == 5:
        ReportReceivedDate = Set_Date(Input_Array[1])
        SortedResponse =[7,"2nd Accept",ReportReceivedDate,"Nothing","Nothing"]
    if TempStatus1 == 6:
        ReportReceivedDate = Set_Date(Input_Array[1])
        SortedResponse =[7,"2nd Reject",ReportReceivedDate,"Nothing","Nothing"]
    if TempStatus1 == 7:
        ReportReceivedDate = Set_Date(Input_Array[1])
        SortedResponse =[7,"2nd Major",ReportReceivedDate,"Nothing","Nothing"]
    if TempStatus1 == 8:
        ReportReceivedDate = Set_Date(Input_Array[1])
        SortedResponse =[7,"2nd Minor",ReportReceivedDate,"Nothing","Nothing"]
    return SortedResponse        

def Read_Manuscript_Date(Input_Array):
    ReceivedDateTemp = Input_Array[3]
    FirstRevisionStartDateTemp = Input_Array[4]
    FirstRevisionDueDateTemp = Input_Array[5]
    SecondRevisionStartDateTemp = Input_Array[6]
    SecondRevisionDueDateTemp = Input_Array[7]
    if len(ReceivedDateTemp)>0:
        ReceivedDateTemp2 =int(ReceivedDateTemp[0][0].astype(datetime.datetime)/1000000000-28800)
        ReceivedDate= datetime.datetime.fromtimestamp(ReceivedDateTemp2)
        if len(FirstRevisionStartDateTemp)>0:
            FirstRevisionStartDateTemp2 =int(FirstRevisionStartDateTemp[0][0].astype(datetime.datetime)/1000000000-28800)
            FirstRevisionStartDate= datetime.datetime.fromtimestamp(FirstRevisionStartDateTemp2)
            if len(FirstRevisionDueDateTemp)>0:
                FirstRevisionDueDateTemp2 =int(FirstRevisionDueDateTemp[0][0].astype(datetime.datetime)/1000000000-28800)
                FirstRevisionDueDate= datetime.datetime.fromtimestamp(FirstRevisionDueDateTemp2)
                if len(SecondRevisionStartDateTemp)>0:
                    SecondRevisionStartDateTemp2 =int(SecondRevisionStartDateTemp[0][0].astype(datetime.datetime)/1000000000-28800)
                    SecondRevisionStartDate= datetime.datetime.fromtimestamp(SecondRevisionStartDateTemp2)
                    if len(SecondRevisionDueDateTemp)>0:
                        SecondRevisionDueDateTemp2 =int(SecondRevisionDueDateTemp[0][0].astype(datetime.datetime)/1000000000-28800)
                        SecondRevisionDueDate= datetime.datetime.fromtimestamp(SecondRevisionDueDateTemp2)
                        ManuscriptDate=[5,"Quite Full",ReceivedDate,FirstRevisionStartDate,FirstRevisionDueDate,SecondRevisionStartDate,SecondRevisionDueDate]
                    else:
                        ManuscriptDate=[4,"No Second Revision Due Date",ReceivedDate,FirstRevisionStartDate,FirstRevisionDueDate,SecondRevisionStartDate,"Nothing"]
                else:
                    ManuscriptDate=[3,"No Second Revision Date",ReceivedDate,FirstRevisionStartDate,FirstRevisionDueDate,"Nothing","Nothing"]
            else:
                ManuscriptDate=[2,"No First Revision Due Date",ReceivedDate,FirstRevisionStartDate,"Nothing","Nothing","Nothing"]
        else:
            ManuscriptDate=[1,"No First Revision Date",ReceivedDate,"Nothing","Nothing","Nothing","Nothing"]
            
    else:
        ManuscriptDate=[0,"No Valid Date","Nothing","Nothing","Nothing","Nothing","Nothing"]
    return ManuscriptDate

def Sort_Case(Input_Array):
    ThisCaseResponses = Input_Array[2]
    ThisCaseAddress = Input_Array[1]
    OtherDate = Read_Manuscript_Date(Input_Array)
    ThisCaseResponsesPrimitiveTreated1 = list(range(len(ThisCaseAddress)))
    ThisCaseResponsesPrimitiveTreated2 = list(range(len(ThisCaseAddress)))
    j=0
    CaseNumber = Input_Array[0].strip()
    FirstWaveReportCollected = 0
    FirstWaveReportReject = 0
    SecondWaveReportCollected = 0
    SecondWaveReportReject = 0
    OngoingPositive = 0
    FirstWaveReport = ["FirstWave"]
    SecondWaveReport = ["SecondWave"]
    for response in ThisCaseResponses:
        ThisCaseResponseSorted1 = Sort_Status(response)
        ThisCaseResponseSorted2 = Sort_Response(ThisCaseResponseSorted1)
        if ThisCaseResponseSorted2[0] == 2:
            ThisCaseResponseSorted = 2
        elif ThisCaseResponseSorted2[0] == 3:
            ThisCaseResponseSorted = 3
        elif ThisCaseResponseSorted2[0] == 4:
            ThisCaseResponseSorted = 4
        elif ThisCaseResponseSorted2[0] == 6:
            ThisCaseResponseSorted = 6
        elif ThisCaseResponseSorted2[0] == 5:
            Status = ThisCaseResponseSorted2[1].split()
            ThisCaseResponseSorted = 1
            Decision = Status[1].upper()
            FirstWaveReport.append(ThisCaseResponseSorted2)
            if Decision.find("EJE")>-1:
                FirstWaveReportReject = FirstWaveReportReject + 1
                FirstWaveReportCollected = FirstWaveReportCollected + 1
            else:
                FirstWaveReportCollected = FirstWaveReportCollected + 1
        elif ThisCaseResponseSorted2[0] == 7:
            Status = ThisCaseResponseSorted2[1].split()
            ThisCaseResponseSorted = 1
            Decision = Status[1].upper()
            SecondWaveReport.append(ThisCaseResponseSorted2)
            if Decision.find("EJE")>-1:
                SecondWaveReportReject = SecondWaveReportReject + 1
                SecondWaveReportCollected = SecondWaveReportCollected + 1
            else:
                SecondWaveReportCollected = SecondWaveReportCollected + 1
        elif ThisCaseResponseSorted2[0] == 8:
            OngoingPositive = OngoingPositive +1
            ThisCaseResponseSorted = 1
        else:
            ThisCaseResponseSorted = 1
        ThisCaseResponsesPrimitiveTreated1[j]= ThisCaseResponseSorted
        ThisCaseResponsesPrimitiveTreated2[j]= ThisCaseResponseSorted2
        j+=1
    AddressNeedsAction1 = ["AddressNeedsAction"]
    ResponseNeedsAction1 = ["ResponseNeedsAction"]
    k = 0
    FirstReminderPendingCount = 0
    AgreedReminderPendingCount = 0
    SecondReminderPendingCount = 0
    AcknowledgePendingCount = 0
    for sorted in ThisCaseResponsesPrimitiveTreated1:
        if ThisCaseResponsesPrimitiveTreated1[k]==2 or ThisCaseResponsesPrimitiveTreated1[k]==3 or ThisCaseResponsesPrimitiveTreated1[k]==4 or ThisCaseResponsesPrimitiveTreated1[k]==6:
            NeedsActionThisTime = ThisCaseResponsesPrimitiveTreated2[k]
            AddressPayAttentionThisTime = ThisCaseAddress[k]
            AddressNeedsAction1.append(AddressPayAttentionThisTime)
            ResponseNeedsAction1.append(NeedsActionThisTime)
        if ThisCaseResponsesPrimitiveTreated1[k]==3:
            AgreedReminderPendingCount = AgreedReminderPendingCount +1
        if ThisCaseResponsesPrimitiveTreated1[k]==4:
            SecondReminderPendingCount = SecondReminderPendingCount +1
        if ThisCaseResponsesPrimitiveTreated1[k]==6:
            FirstReminderPendingCount = FirstReminderPendingCount +1
        if ThisCaseResponsesPrimitiveTreated1[k]==2:
            AcknowledgePendingCount = AcknowledgePendingCount + 1
        k+=1
    RNAN1 = np.array(ResponseNeedsAction1)
    RNAN2 = RNAN1.reshape(1,len(ResponseNeedsAction1))[0]
    RNAL1 = list(RNAN2)
    ResponseNeedsAction = RNAL1[1:len(RNAL1)]
    ANAN1 = np.array(AddressNeedsAction1)
    ANAN2 = ANAN1.reshape(1,len(AddressNeedsAction1))[0]
    ANAL1 = list(ANAN2)
    AddressNeedsAction = ANAL1[1:len(ANAL1)]
    NeedAttentionCount = len(AddressNeedsAction)
    FWRN1 = np.array(FirstWaveReport)
    FWRN2 = FWRN1.reshape(1,len(FirstWaveReport))[0]
    FWRL1 = list(FWRN2)
    FWRL2 = list(FWRN1[1:len(FWRL1)])
    SWRN1 = np.array(SecondWaveReport)
    SWRN2 = SWRN1.reshape(1,len(SecondWaveReport))[0]
    SWRL1 = list(SWRN2)
    SWRL2 = list(SWRN1[1:len(SWRL1)])

    if len(FWRL2)>0:
        if len(SWRL2)>0:
            OriginalMax = SWRL2[0][2]
            m=0
            NowMax = OriginalMax
            while m<len(SWRL2):
               TestChoice = SWRL2[m][2]
               DateDiff =rrule.rrule(rrule.DAILY,dtstart=TestChoice,until=NowMax).count()
               if DateDiff<=0:
                   NowMax = SWRL2[m][2]
               m+=1
            LatestDate = [2,NowMax]
        else:
            OriginalMax = FWRL2[0][2]
            m=0
            NowMax = OriginalMax
            while m<len(FWRL2):
               TestChoice = FWRL2[m][2]
               DateDiff =rrule.rrule(rrule.DAILY,dtstart=TestChoice,until=NowMax).count()
               if DateDiff<=0:
                   NowMax = FWRL2[m][2]
               m+=1
            LatestDate = [1,NowMax]
    else:
        if len(SWRL2)>0:
            OriginalMax = SWRL2[0][2]
            m=0
            NowMax = OriginalMax
            while m<len(SWRL2):
               TestChoice = SWRL2[m][2]
               DateDiff =rrule.rrule(rrule.DAILY,dtstart=TestChoice,until=NowMax).count()
               if DateDiff<=0:
                   NowMax = SWRL2[m][2]
               m+=1
            LatestDate = [3,NowMax]
        else:
            LatestDate= [0,"Nothing"]
    RevisionAlert=[-1,"No Alert","Nothing","Nothing"]
    if OtherDate[0]==5 and len(FWRL2)==0 and len(SWRL2) ==0:
        DueDate =OtherDate[4]
        CurrentDate = datetime.date.today()
        DateDiff1 = rrule.rrule(rrule.DAILY,dtstart=DueDate,until=CurrentDate).count()
        DateDiff2 = rrule.rrule(rrule.DAILY,dtstart=CurrentDate,until=DueDate).count()
        if DateDiff2 <4 or DateDiff1>0:
            RevisionAlert = [2,"Second Revision Alert",OtherDate[4],LatestDate[1]]
    if OtherDate[0]==3 and len(SWRL2)==0 and SecondReminderPendingCount ==0:
        DueDate =OtherDate[4]
        CurrentDate = datetime.date.today()
        DateDiff1 = rrule.rrule(rrule.DAILY,dtstart=DueDate,until=CurrentDate).count()
        DateDiff2 = rrule.rrule(rrule.DAILY,dtstart=CurrentDate,until=DueDate).count()
        if DateDiff2 <4 or DateDiff1>0:
            RevisionAlert = [1,"First Revision Alert",OtherDate[2],LatestDate[1]]
    if OtherDate[0]==0:
        RevisionAlert = [0,"Please Check Receive Date","Nothing","Nothing"]
    if LatestDate[0]==1 and OtherDate[0]==1:
        DueDate= LatestDate[1]
        CurrentDate = datetime.date.today()
        DateDiff1 = rrule.rrule(rrule.DAILY,dtstart=DueDate,until=CurrentDate).count()
        if DateDiff1>10:
            if FirstWaveReportReject>0:
                RevisionAlert = [3,"Guess Rejected?",DueDate,"Nothing"]
            else:
                RevisionAlert = [4,"Please Check Current Case Status","Nothing","Nothing"]            
    TotalPositive = OngoingPositive + AgreedReminderPendingCount + SecondReminderPendingCount + FirstWaveReportCollected + SecondWaveReportCollected
    
    if TotalPositive>2:
        k = 0
        Index =[]
        FirstReminderPendingCount = 0

        
        while k < len(ResponseNeedsAction):
            if ResponseNeedsAction[k][0] == 6:
                Index.append(k)
            k+=1
            
        ThoseNeedToDeleteInResponse=[]
        ThoseNeedToDeleteInAddress=[]
        
        for each in Index:
            ThoseNeedToDeleteInResponse.append(ResponseNeedsAction[each])
            ThoseNeedToDeleteInAddress.append(AddressNeedsAction[each])

        for each in ThoseNeedToDeleteInResponse:
            ResponseNeedsAction.remove(each)
            
        for each in ThoseNeedToDeleteInAddress:
            ThoseNeedToDeleteInAddress.remove(each)

    Result = [CaseNumber,FirstReminderPendingCount,AcknowledgePendingCount,AgreedReminderPendingCount,SecondReminderPendingCount,FirstWaveReportCollected,FirstWaveReportReject,SecondWaveReportCollected,SecondWaveReportReject,FWRL2,SWRL2,ResponseNeedsAction,AddressNeedsAction,OtherDate,OtherDate[0],LatestDate[0],LatestDate[1],RevisionAlert]
    return Result

i=1
CaseSortingFinalized1 = ["Elements","Address Need To Action","Respective Response","Receiving Date","First Revision Date","First Revision Due Date","Second Revision Date","Second Revision Due Date","Pending Acknowledgement","Pending Reminder after Agreed",] 
CaseSortingFinalized = np.array(CaseSortingFinalized1)

global IfContinueToProcess 
global ValidFileDetected 

IfContinueToProcess = 0
ValidFileDetected = 0

root = Tk()

def File_Select():
    Status = [0,"Nothing Selected"]
    filename = tkinter.filedialog.askopenfilename(filetypes=[('XLS', '*.xls'), ('XLSX', 'xlsx')])
    if filename !='':
        Status = [1,filename]
    else:
        Status = [0,"Nothing Selected"]
    root.destroy()
    if Status[0]==1:
        global IfContinueToProcess
        IfContinueToProcess =1
        print("Retrieving data from",Status[1])
    else:
        print("Nothing selected")
    return Status   
  
FilePath= File_Select()
root.mainloop()

if IfContinueToProcess == 1:
   TempRestoration=pd.read_excel(FilePath[1],sheet_name='Sheet1')
   print(TempRestoration.columns[1].upper())
   if TempRestoration.columns[1].upper() != "TRIGGER":
       print("INVALID DATAFILE!")
   else:
       CaseSorted1 =Primitive_Process(TempRestoration)
       CaseSorted = CaseSorted1[2]
       CaseIndexHor = CaseSorted1[0]
       while i<=CaseIndexHor.shape[0]:
           ThisCase = CaseSorted[i]
           print("Processing",ThisCase[0].strip())
           Result=Sort_Case(ThisCase)
           ReviAlert= Result[-1]
           if ReviAlert[0] == -1:
               print("%s done, %d first pending reminder(s), %d pending acknowledgement(s), %d agreed needs reminder, %d second review need reminder, %d first wave reports collected, %d first wave rejected, %d second wave reports collected, %d second wave rejected" %(Result[0],Result[1],Result[2],Result[3],Result[4],Result[5],Result[6],Result[7],Result[8]))        
               print("+++++++")
               j = 0
               while j < len(Result[11]):
                   print("%s %s Last reaction date: %s Day passed %s" %(Result[12][j],Result[11][j][1],Result[11][j][2],str(Result[11][j][4])))
                   j+=1
               print("------------------------------------------------------------")    
           else:
               print("%s done, %d first pending reminder(s), %d pending acknowledgement(s), %d agreed needs reminder, %d second review need reminder, %d first wave reports collected, %d first wave rejected, %d second wave reports collected, %d second wave rejected, %s" %(Result[0],Result[1],Result[2],Result[3],Result[4],Result[5],Result[6],Result[7],Result[8],ReviAlert[1]))
               print("+++++++")
               j = 0
               while j < len(Result[11]):
                   print("%s %s Last reaction date: %s Day passed %s" %(Result[12][j],Result[11][j][1],Result[11][j][2],str(Result[11][j][4])))
                   j+=1
               print("------------------------------------------------------------")
           i+=1





        
    