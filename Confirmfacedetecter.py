import cv2
from datetime import datetime
import openpyxl
from keras.models import load_model
import numpy as np
from total_data import total_data


wb = openpyxl.Workbook()
ws = wb.active


def in_out():

    model = load_model("keras_Model.h5", compile=False)
    class_names = open("labels.txt", "r").readlines()
    
    cap = cv2.VideoCapture(0)
    right, left = "", ""

    

    data = (
        ("Visitor", "Name", "Date", "Time"),
    )
    for i in data:
        ws.append(i)
        wb.save("Book1.xlsx")

    unknown_in_count = 0
    unknown_out_count = 0    
    in_count = 0 
    out_count = 0

    while True:

        ret, image = cap.read()
        image = cv2.resize(image, (224, 224), interpolation=cv2.INTER_AREA)
        image = np.asarray(image, dtype=np.float32).reshape(1, 224, 224, 3)
        image = (image / 127.5) - 1

        prediction = model.predict(image)
        index = np.argmax(prediction)
        class_name = class_names[index]
        confidence_score = prediction[0][index]

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
            
            if x < 200:
                print("to left")
                x = 300
                right, left = "", ""
                file_name = f"visitors/in/{datetime.now().strftime('%y-%m-%d-%H-%M-%S')}.jpg"
                cv2.imwrite(file_name, frame1)

                in_count += 1

                if (np.round(confidence_score * 100)==100):
                    #print("Class:", class_name[2:], end="")
                    person = class_name[2:]
                    print(person)
                    data = (
                        ("In", person, f"{datetime.now().strftime('%d-%m-%y')}",f"{datetime.now().strftime('%H:%M:%S')}"),
                    )
                    
                else:
                    person = "unknown"
                    print(person)
                    data = (
                        ("In", person, f"{datetime.now().strftime('%d-%m-%y')}",f"{datetime.now().strftime('%H:%M:%S')}"),
                    )
                    unknown_in_count += 1
                
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

                out_count += 1

                if (np.round(confidence_score * 100)==100):
                    #print("Class:", class_name[2:], end="")
                    person = class_name[2:]
                    print(person)
                    data = (
                        ("Out", person, f"{datetime.now().strftime('%d-%m-%y')}",f"{datetime.now().strftime('%H:%M:%S')}"),
                    )
                else:
                    person = "unknown"
                    print(person)
                    data = (
                        ("Out", person, f"{datetime.now().strftime('%d-%m-%y')}",f"{datetime.now().strftime('%H:%M:%S')}"),
                    )
                    unknown_out_count += 1

                for i in data:
                    ws.append(i)
                    wb.save("Book1.xlsx")
            
        cv2.imshow("", frame1)
        
        
        k = cv2.waitKey(1)
        
        if k == 27:
            cap.release()
            cv2.destroyAllWindows()
            break

    total_data(in_count, out_count, unknown_in_count, unknown_out_count)

