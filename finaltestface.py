import cv2, sys, numpy, os
size = 4
haar_file = 'haarcascade_frontalface_default.xml'
dataset = 'dataset'
from datetime import datetime
import openpyxl
from identifying_the_testfaces import identifying_the_testfaces

wb = openpyxl.Workbook()
ws = wb.active

def in_out():
    
    cap = cv2.VideoCapture(0)
    right, left = "", ""
    
    data = (
        ("Visitors", "Date", "Time"),
    )
       
    for i in data:
        ws.append(i)
        wb.save("Book1.xlsx")
    
    
        

    while True:
        _, frame1 = cap.read()
        frame1 = cv2.flip(frame1, 1)
        _, frame2 = cap.read()
        frame2 = cv2.flip(frame2, 1)

        diff = cv2.absdiff(frame2, frame1)
        diff = cv2.blur(diff, (5,5))
        gray = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)
        _, threshd = cv2.threshold(gray, 40, 255, cv2.THRESH_BINARY)
        contr, _ = cv2.findContours(threshd, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
        x = 300

        
        
        if len(contr) > 0:
            max_cnt = max(contr, key=cv2.contourArea)
            x,y,w,h = cv2.boundingRect(max_cnt)
            cv2.rectangle(frame1, (x, y), (x+w, y+h), (0,255,0), 2)
            cv2.putText(frame1, "MOTION", (10,80), cv2.FONT_HERSHEY_SIMPLEX, 2, (0,255,0), 2)
            
        if right == "" and left == "":
            if x > 500:
                right = True
            elif x < 200:
                left = True
                
        elif right:
            n = 0
            if x < 200:
                n = n+1
                print("to left")
                x = 300
                right, left = "", ""
                file_name = f"visitors/in/{datetime.now().strftime('%y-%m-%d-%H-%M-%S')}.jpg"
                cv2.imwrite(file_name, frame1)
                
                data = (
                    ("In", f"{datetime.now().strftime('%d-%m-%y')}",f"{datetime.now().strftime('%H:%M:%S')}"),
                )
                for i in data:
                    ws.append(i)
                    wb.save("Book1.xlsx")

            
        elif left:
            if x > 500:
                print("to right")
                x = 300
                right, left = "", ""
                file_name = f"visitors/out/{datetime.now().strftime('%d-%m-%y-%H-%M-%S')}.jpg"
                cv2.imwrite(file_name, frame1)
                data = (
                    ("Out", f"{datetime.now().strftime('%d-%m-%y')}",f"{datetime.now().strftime('%H:%M:%S')}"),
                )
                for i in data:
                    ws.append(i)
                    wb.save("Book1.xlsx")
            
        cv2.imshow("", frame1)

        
        
        
        
        k = cv2.waitKey(1)
        
        if k == 27:
            cap.release()
            cv2.destroyAllWindows()
            break

