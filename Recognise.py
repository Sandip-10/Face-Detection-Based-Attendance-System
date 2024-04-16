import os
import cv2
import numpy as np
import xlwrite
import time
from playsound import playsound

start = time.time()
period = 8
face_cas = cv2.CascadeClassifier('E:\\WorksSpace\\Others_Projects\\Face_Project\\haarcascade_frontalface_default.xml')
cap = cv2.VideoCapture(0)
recognizer = cv2.face.LBPHFaceRecognizer_create()
recognizer.read('E:\\WorksSpace\\Others_Projects\\Face_Project\\trainer.yml')
flag = 0
filename = 'filename'
dict = {'item1': 1}
font = cv2.FONT_HERSHEY_SIMPLEX

# Create directory if it doesn't exist
directory = 'attendance_files'
if not os.path.exists(directory):
    os.makedirs(directory)

while True:
    ret, img = cap.read()
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    faces = face_cas.detectMultiScale(gray, 1.3, 7)

    for (x, y, w, h) in faces:
        roi_gray = gray[y:y + h, x:x + w]
        cv2.rectangle(img, (x, y), (x + w, y + h), (255, 0, 0), 2)
        id, conf = recognizer.predict(roi_gray)

        if conf < 50:
            if id == 1:
                id = 'Sandip Karmakar'
                if str(id) not in dict:
                    filename = xlwrite.output('attendance', 'class1', 1, id, 'yes')
                    dict[str(id)] = str(id)
            elif id == 2:
                id = 'Reju Karmakar'
                if str(id) not in dict:
                    filename = xlwrite.output('attendance', 'class1', 2, id, 'yes')
                    dict[str(id)] = str(id)
            elif id == 3:
                id = 'Chiranjib Nag'
                if str(id) not in dict:
                    filename = xlwrite.output('attendance', 'class1', 3, id, 'yes')
                    dict[str(id)] = str(id)
        else:
            id = 'Unknown, cannot recognize'
            flag = flag + 1
            break

        cv2.putText(img, str(id) + " " + str(conf), (x, y - 10), font, 0.55, (120, 255, 120), 1)

    cv2.imshow('frame', img)

    if flag == 10:
        playsound('transactionSound.mp3')
        print("Transaction Blocked")
        break
    if time.time() > start + period:
        break
    if cv2.waitKey(100) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
