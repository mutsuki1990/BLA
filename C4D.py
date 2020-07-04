import c4d
from c4d import gui
import csv
import cmath
import re
# Welcome to the world of Python


# Script state in the menu or the command palette
# Return True or c4d.CMD_ENABLED to enable, False or 0 to disable
# Alternatively return c4d.CMD_ENABLED|c4d.CMD_VALUE to enable and check/mark
#def state():
#    return True

# Main function

def OpenSource(Path):
    print ("Reading Data")
    CSV1 = open(Path)
    CSVFile1 = csv.reader(CSV1)
    VData = []

    for Row in CSVFile1:

        Senti = Row[1]
        Confi = Row[2]
        LikeCount = Row[0]
        if Senti == "0":
            Mid = 2*cmath.sqrt(cmath.sqrt(float(Row[0])+1))
            H = int(Mid.real)+int(float(Row[4])*10)
            Hei = H
            TR = (1-float(Row[4]))*0.7
            CP = Row[4]
        elif Senti == "1":
            Mid1 = cmath.sqrt(float(Row[0])+1)
            Mid2 = cmath.sqrt(int(Mid1.real)+int(float(Row[3])*10))
            H = int(Mid2.real)
            TR = -1
            CP = Row[3]
        elif Senti == "2":
            Mid1 = cmath.sqrt(float(Row[0])+1)
            Mid2 = cmath.sqrt(cmath.sqrt(int(Mid1.real)+int(float(Row[3])*10)))
            H = int(Mid2.real)
            TR = -1
            CP = Row[3]
        PointPara = [Senti,LikeCount,H,TR,CP]
        VData.append(PointPara)
    print ("Reading Done")
    return VData

def GenerateMatrix():

    print ("Generating Matrix")
    AxixY = 0
    AXA = []

    for i in range(1,81):

        AxixX = 0
        AXs = []

        for j in range(1,126):

            Cordinate = [AxixX,AxixY]
            AXA.append(Cordinate)
            AxixX = AxixX + 1

        AxixY = AxixY + 1

    print ("Generation Done")
    return AXA


def CreateOb(Source,Matrix):

    Index = 0
    for EachData in Source:

        print ("Generating Object at %d, %d" %(Matrix[Index][0],Matrix[Index][1]))

        if EachData[0] == "0":

            obj = c4d.BaseObject(c4d.Ocone)
            obj.SetRelPos(c4d.Vector(Matrix[Index][0],Matrix[Index][1],0))
            doc.InsertObject(obj)
            Cone = doc.GetFirstObject()
            RotVector =(0.78,0,0)
            Cone[c4d.PRIM_CONE_BRAD] = 0.7
            Cone[c4d.PRIM_CONE_TRAD] = EachData[3]
            Cone[c4d.PRIM_CONE_SEG] = 4
            H = EachData[2]
            Cone[c4d.PRIM_CONE_HEIGHT]=H
            obj.SetRelPos(c4d.Vector(Matrix[Index][0],H/2,Matrix[Index][1]))
            obj.SetAbsRot(RotVector)
            matlist = doc.GetMaterials()
            if float(EachData[4])> 0.4 and float(EachData[4])<=0.85:
                mat1 = doc.SearchMaterial("Anti1")
                textag1 = c4d.TextureTag()
                textag1.SetMaterial(mat1)
                textag1[c4d.TEXTURETAG_PROJECTION]=c4d.TEXTURETAG_PROJECTION_UVW
                Cone.InsertTag(textag1)
            elif float(EachData[4])> 0.85 and float(EachData[4])<=0.96:
                mat2 = doc.SearchMaterial("Anti2")
                textag2 = c4d.TextureTag()
                textag2.SetMaterial(mat2)
                textag2[c4d.TEXTURETAG_PROJECTION]=c4d.TEXTURETAG_PROJECTION_UVW
                Cone.InsertTag(textag2)
            else:
                mat3 = doc.SearchMaterial("Anti3")
                textag3 = c4d.TextureTag()
                textag3.SetMaterial(mat3)
                textag3[c4d.TEXTURETAG_PROJECTION]=c4d.TEXTURETAG_PROJECTION_UVW
                Cone.InsertTag(textag3)

            c4d.EventAdd()
            Index = Index + 1

        elif EachData[0] == "1":

            obj = c4d.BaseObject(c4d.Ocylinder)
            obj.SetRelPos(c4d.Vector(Matrix[Index][0],Matrix[Index][1],0))
            doc.InsertObject(obj)
            Cyd = doc.GetFirstObject()
            Cyd[c4d.PRIM_CYLINDER_RADIUS] = 0.7
            Cyd[c4d.PRIM_CYLINDER_SEG] = 4
            RotVector =(0.78,0,0)
            H = EachData[2]
            Cyd[c4d.PRIM_CYLINDER_HEIGHT]=H
            obj.SetRelPos(c4d.Vector(Matrix[Index][0],-H/2,Matrix[Index][1]))
            obj.SetAbsRot(RotVector)
            mat = doc.SearchMaterial("Neutral")
            textag = c4d.TextureTag()
            textag.SetMaterial(mat)
            textag[c4d.TEXTURETAG_PROJECTION]=c4d.TEXTURETAG_PROJECTION_UVW
            Cyd.InsertTag(textag)
            c4d.EventAdd()
            Index = Index + 1

        else:

            obj = c4d.BaseObject(c4d.Ocylinder)
            obj.SetRelPos(c4d.Vector(Matrix[Index][0],Matrix[Index][1],0))
            doc.InsertObject(obj)
            Cyd2 = doc.GetFirstObject()
            Cyd2[c4d.PRIM_CYLINDER_RADIUS] = 0.7
            Cyd2[c4d.PRIM_CYLINDER_SEG] = 4
            RotVector =(0.78,0,0)
            H = EachData[2]
            Cyd2[c4d.PRIM_CYLINDER_HEIGHT]=H
            obj.SetRelPos(c4d.Vector(Matrix[Index][0],-H/2,Matrix[Index][1]))
            obj.SetAbsRot(RotVector)
            print(EachData[4])
            if float(EachData[4])> 0.4 and float(EachData[4])<=0.8:
                mat4 = doc.SearchMaterial("Sup1")
                textag4 = c4d.TextureTag()
                textag4.SetMaterial(mat4)
                textag4[c4d.TEXTURETAG_PROJECTION]=c4d.TEXTURETAG_PROJECTION_UVW
                Cyd2.InsertTag(textag4)
                
            elif float(EachData[4])> 0.8 and float(EachData[4])<=0.96:
                mat5 = doc.SearchMaterial("Sup2")
                textag5 = c4d.TextureTag()
                textag5.SetMaterial(mat5)
                textag5[c4d.TEXTURETAG_PROJECTION]=c4d.TEXTURETAG_PROJECTION_UVW
                Cyd2.InsertTag(textag5)
                c4d.EventAdd()
            else:
                mat6 = doc.SearchMaterial("Sup3")
                textag6 = c4d.TextureTag()
                textag6.SetMaterial(mat6)
                textag6[c4d.TEXTURETAG_PROJECTION]=c4d.TEXTURETAG_PROJECTION_UVW
                Cyd2.InsertTag(textag6)
                c4d.EventAdd()

            
            Index = Index + 1


    return



def main():

    Path1 = "D:\\C4D.csv"
    A = OpenSource(Path1)
    B = GenerateMatrix()
    CreateOb(A,B)


# Execute main()
if __name__=='__main__':
    main()