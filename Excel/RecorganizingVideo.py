import cv2
import os
import datetime
import pandas as pd
import openpyxl as xl
from openpyxl import Workbook
import xlwings as xs
import xlrd
import xlwt
from xlwt import Workbook
import numpy as ny

def User(User_name):
    #Read User name User Excel and Store an array
    workbook = xlrd.open_workbook("C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Face\\User.xls")
    worksheet=workbook.sheet_by_name('sheet1')
    num_rows=worksheet.nrows
    num_colum=worksheet.ncols
    name=[]
   
    for i in range(0,num_rows):
        in_details=[]
        for j in range(0,num_colum):
            in_details.append(worksheet.cell_value(i,j))
        name.append(in_details)
    
    workbook.release_resources()

    check_name=0
    for i in range(0,len(name)):
        for j in range(0,len(name[0])):
            if(name[i][j]==User_name):
                check_name=1

    user_not=1
    if check_name==0:
        print("Not in that person Database")
        user_not=0
        return
    else:
        Excel_UserName_Entry(User_name,user_not)


#Excel Process
def Excel_UserName_Entry(Atten_name,user_check):
    #Get the current date
    date=datetime.datetime.now()
    curr_date="{}.{}.{}".format(date.year,date.month,date.day)
    In="IN-{}h.{}m.{}s".format(date.hour,date.minute,date.second)


    #Old Date Store and write again
    workbook2 = xlrd.open_workbook("C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Face\\Entry1.xls")
    worksheet2=workbook2.sheet_by_name('sheet1')
    num_rows2=worksheet2.nrows
    num_colum2=worksheet2.ncols
    curr_row2 = 0
    curr_cols2 = 0
    details=[]
    in_details=[]

    for i in range(0,num_rows2):
        in_details=[]
        for j in range(0,num_colum2):
            in_details.append(worksheet2.cell_value(i,j))
        details.append(in_details)
    
    
    if user_check==1:
        Not_in_name=0
        for i in range(0,len(details)):
            for j in range(0,len(details[0])):
                if(details[i][j]==Atten_name):
                    Not_in_name=1
    
    if(Not_in_name==0):
        in_details=[len(details[0])]
        in_details[0]=Atten_name
        details.append(in_details)
        print("Name Upadated")
    
    workbook2.release_resources()
    #Write the user name on Entry file upadate
    wb = xlwt.Workbook()
    sheet1=wb.add_sheet("sheet1",cell_overwrite_ok=True)
    Check_curr_date = 0
    Check_curr_name = 0
    for i in range(0,len(details)):
        for j in range(0,len(details[0])):
            sheet1.write(i,j,details[i][j])
            if i==0 and Check_curr_date == 0:
                if details[i][j]==curr_date:
                    Check_curr_date = 1                 
                    Date_index=j
            if j==0 and Check_curr_name == 0:
                if details[i][j]==Atten_name:
                    Name_index=i
                    Check_curr_name = 1


    
    #Check the current date ,If is it not there write current date
    if Check_curr_date==0:
        sheet1.write(0,len(details[0]),curr_date)
        Date_index=len(details[0])-1
        in_details=details[0]
        in_details.append(curr_date)        
        Check_curr_date = 1
        print("Date")
       
    if(Check_curr_date== 1 and Check_curr_name==1 and details[Name_index][Date_index]=="" and details[0][Date_index]==curr_date):
        in_details=details[Name_index]
        in_details[Date_index]=("IN")
        sheet1.write(Name_index,Date_index,In)
            #sheet1.write(Name_index,Date_index,"IN")
        print("Done")
            
    
    wb.save("C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Face\\Entry1.xls")
    return




#Face detect process
def webcam(Unknown):
    #Haarcase File upload
    faceCascade = cv2.CascadeClassifier("C:\\Users\\Uthayakumar Thenujan\\AppData\\Local\\Programs\\Python\\Python310\\Lib\\site-packages\\cv2\\data\\haarcascade_frontalface_default.xml")

    #for our trained face upload
    recognizer=cv2.face.LBPHFaceRecognizer_create()
    recognizer.read("C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Face\\training.yml")

    #Get for the file name as a array
    names=[]

    #Get the file directory and save to names array
    for name in os.listdir("C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Face\\dataset"):
        names.append(name)

    #on the webcamera
    webcam=cv2.VideoCapture(0)

    #Face Detect process loop utill press Q or q
    while True:
        successful_frame_read, frame=webcam.read() #get the input as a frame from webcam
        gray_img=cv2.cvtColor(frame,cv2.COLOR_BGR2GRAY) # chance the frame color to gray
        face_Coordi=faceCascade.detectMultiScale(gray_img,scaleFactor=1.2,minNeighbors=5,minSize=(50,50)) #Detect the face coordination from gray frame

        #Check the every face of the perticular frame in loop    
        for (x,y,w,h) in face_Coordi:          
            id,confidence =recognizer.predict(gray_img[y:y+h, x:x+w]) # predict the our trained face

            if confidence<100: #check the confidence level of the predict
                cv2.rectangle(frame,(x,y),(x+w,y+h),(0,255,0),5)
                cv2.putText(frame,names[id-1],(x,y-4), cv2.FONT_HERSHEY_SIMPLEX, 0.9 , (0,0,255), 2 , cv2.LINE_AA)
                if(Unknown==0):
                    filename="C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Excel\\Emp_Un_Time\\Untime{}.{}Y.{}M.{}D._{}h.{}m.{}s.jpg".format(names[id-1],date.year,date.month,date.day,date.hour,date.minute,date.second)
                    cv2.imwrite(filename,frame)
                    break
                if(name[id-1!=""]):
                    User(names[id-1])
                           
            else:
                cv2.rectangle(frame,(x,y),(x+w,y+h),(0,0,255),5)
                cv2.putText(frame,"Unkonwn",(x,y-8), cv2.FONT_HERSHEY_SIMPLEX, 0.9 , (0,255,0), 2 , cv2.LINE_AA)
                date=datetime.datetime.now()
                if(Unknown==1):
                    filename="C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Excel\\Unknow_Office_Time\\Unkown{}Y.{}M.{}D._{}h.{}m.{}s.jpg".format(date.year,date.month,date.day,date.hour,date.minute,date.second)
                    cv2.imwrite(filename,frame)
                elif(Unknown==0):
                    filename="C:\\Users\\Uthayakumar Thenujan\\OneDrive\\Desktop\\Python\\Excel\\Unknow_Un_Time\\Untime{}Y.{}M.{}D._{}h.{}m.{}s.jpg".format(date.year,date.month,date.day,date.hour,date.minute,date.second)
                    cv2.imwrite(filename,frame)
            
                
        cv2.imshow("Identify Face",frame) # Show the every frame as live
        key=cv2.waitKey(1) #1ms automatically click any key for change the next frame,and get the input key
        if key==81 or key==113: # if isert to q or Q , process will stoped
            break

date=datetime.datetime.now()
timeH=date.hour
timeM=date.minute
timeS=date.second

if((timeH>=6 and timeH<=18 and timeM>=0 and timeS>=0)):
    unknown=1
    
else:
    print("This Not Office time{}.{}.{}".format(date.hour,date.minute,date.second))
    unknown=0

webcam(unknown)
