from tkinter import *
from tkinter import ttk
from PIL import Image, ImageTk
import cv2
import numpy as np
import face_recognition
import csv
import os
from datetime import datetime
from twilio.rest import Client
import xlsxwriter
 

class Face_Recognition_System:
    def _init_(self, root):
        self.root = root
        self.root.geometry("1530x790+0+0")
        self.root.title("Face Recognition System")
        # Set current date
        self.current_date = datetime.now().strftime("%Y-%m-%d")  
        self.attendance_folder = "attendance"  # Folder name for storing attendance sheets
        self.create_attendance_folder()  # Create the attendance folder if it doesn't exist


        # First image
        img = Image.open(r"C:\Users\HP\Pictures\New folder\img1.jpg")
        img = img.resize((500, 130))
        self.photoimg = ImageTk.PhotoImage(img)
        f_lbl = Label(self.root, image=self.photoimg)
        f_lbl.place(x=0, y=0, width=500, height=130)

        # Second Image
        img1 = Image.open(r"C:\Users\HP\Pictures\New folder\img2.png")
        img1 = img1.resize((500, 130))
        self.photoimg1 = ImageTk.PhotoImage(img1)
        f_lbl = Label(self.root, image=self.photoimg1)
        f_lbl.place(x=500, y=0, width=500, height=130)

        # Third image
        img2 = Image.open(r"C:\Users\HP\Pictures\New folder\img3.png")
        img2 = img2.resize((500, 130))
        self.photoimg2 = ImageTk.PhotoImage(img2)
        f_lbl = Label(self.root, image=self.photoimg2)
        f_lbl.place(x=1000, y=0, width=550, height=130)

        # Background Image
        img3 = Image.open(r"C:\Users\HP\Pictures\New folder\backimg.jpg")
        img3 = img3.resize((1530, 710))
        self.photoimg3 = ImageTk.PhotoImage(img3)
        bg_img = Label(self.root, image=self.photoimg3)
        bg_img.place(x=0, y=130, width=1530, height=710)

        title_lbl = Label(bg_img, text="FACE RECOGNITION ATTENDANCE SYSTEM", font=("times new roman", 35, "bold"), bg="white", fg="red")
        title_lbl.place(x=0, y=0, width=1530, height="45")

        # Take attendance button
        img4 = Image.open(r"C:\Users\HP\Pictures\New folder\img4.jpg")
        img4 = img4.resize((220, 220))
        self.photoimg4 = ImageTk.PhotoImage(img4)

        b1 = Button(bg_img, image=self.photoimg4, command=self.student_details, cursor="hand2")
        b1.place(x=100, y=100, width=400, height=250)

        b1_1 = Button(bg_img, text="Student details", command=self.student_details, cursor="hand2", font=("times new roman", 15, "bold"), bg="dark blue", fg="white")
        b1_1.place(x=100, y=310, width=400, height=40)

        # Add new user
        img5 = Image.open(r"C:\Users\HP\Pictures\New folder\img5.png")
        img5 = img5.resize((220, 220))
        self.photoimg5 = ImageTk.PhotoImage(img5)

        b2 = Button(bg_img, image=self.photoimg5, command=self.face_detector, cursor="hand2")
        b2.place(x=550, y=100, width=400, height=250)

        b2_1 = Button(bg_img, text="Face Detector", command=self.face_detector, cursor="hand2", font=("times new roman", 15, "bold"), bg="dark blue", fg="white")
        b2_1.place(x=550, y=310, width=400, height=40)
        
        
        #exit button
        img7 = Image.open(r"C:\Users\HP\Pictures\New folder\Exit-button-icon-png.jpg")
        img7 = img7.resize((220, 220))
        self.photoimg7 = ImageTk.PhotoImage(img7)

        b2 = Button(bg_img, image=self.photoimg7, command=self.exit_program, cursor="hand2")
        b2.place(x=550, y=380, width=400, height=200)

        b2_1 = Button(bg_img, text="EXIT", command=self.exit_program, cursor="hand2", font=("times new roman", 15, "bold"), bg="dark blue", fg="white")
        b2_1.place(x=550, y=580, width=400, height=40)

        # View attendance
        img6 = Image.open(r"C:\Users\HP\Pictures\New folder\img6.png")
        img6 = img6.resize((220, 220))
        self.photoimg6 = ImageTk.PhotoImage(img6)

        b3 = Button(bg_img, image=self.photoimg6, command=self.view_attendance, cursor="hand2")
        b3.place(x=1000, y=100, width=400, height=250)

        b3_1 = Button(bg_img, text="Attendance", command=self.view_attendance, cursor="hand2", font=("times new roman", 15, "bold"), bg="dark blue", fg="white")
        b3_1.place(x=1000, y=310, width=400, height=40)
        
        

        # Initialize face recognition variables
        self.path = "D:\myenv\Training_images"
        self.images = []
        self.classNames = []
        self.myList = os.listdir(self.path)

        # Load images and class names for face recognition
        for cl in self.myList:
            curImg = cv2.imread(f'{self.path}/{cl}')
            self.images.append(curImg)
            self.classNames.append(os.path.splitext(cl)[0])

        self.encodeListKnown = self.findEncodings(self.images)
        print('Encoding Complete')

#         self.account_sid = 'ACce9b4f3617569ad22d0e8ac63093ce30'
#         self.auth_token = '3ae2825fa2af721b0b5f77fe2f126ccc'
#         self.twilio_phone_number = '+13158030991'
#         self.client = Client(self.account_sid, self.auth_token)

        # Create an Excel workbook and add a worksheet
        #self.workbook = xlsxwriter.Workbook(self.current_date + '.xlsx')
        #self.worksheet = self.workbook.add_worksheet()
        self.workbook = xlsxwriter.Workbook(os.path.join(self.attendance_folder, f'{self.current_date}.xlsx'))
        self.worksheet = self.workbook.add_worksheet()

        # Write column headers
#         self.worksheet.write(0, 0, 'Roll Number')
#         self.worksheet.write(0, 1, 'Name')
#         self.worksheet.write(0, 2, 'Time')

        self.row = 1  # Initialize row counter
    def create_attendance_folder(self):
        if not os.path.exists(self.attendance_folder):
            os.makedirs(self.attendance_folder)


    def findEncodings(self, images):
        encodeList = []
        for img in images:
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            encode = face_recognition.face_encodings(img)[0]
            encodeList.append(encode)
        return encodeList

    def markAttendance(self, roll_no, name):
        current_time = datetime.now().strftime('%H:%M:%S')
        hour = int(datetime.now().strftime('%H'))
        minute = int(datetime.now().strftime('%M'))
        current_minutes = hour * 60 + minute

        # Check if it is between 12 am to 12 pm (morning)
        if 0 <= hour < 10 or (hour == 10 and minute < 36):
            lect_column = "Lect 1"
            with open(os.path.join(self.attendance_folder, f'{self.current_date}.csv'), 'a+', newline='') as f:
                f.seek(0)
                myDataList = f.readlines()
                nameList = []

                if len(myDataList) > 0:
                    for line in myDataList:
                        entry = line.split(',')
                        if entry[0] == roll_no and entry[2].strip()[:8] == current_time[:8]:
                    # Attendance for the same student within the same hour already marked
                            print(f"Attendance for {name} with roll number {roll_no} already marked in the current hour.")
                            return

                        nameList.append(entry[1].strip())
            
                if name not in nameList:
                    now = datetime.now()
                    dtString = now.strftime('%H:%M:%S')
                    writer = csv.writer(f)

                    if f.tell() == 0:
                        writer.writerow(["Roll No.", "Name", "Lect 1", "Lect 2"])

                    row_data = [roll_no, name, "",""]

            # Add empty column values for the column that is not being updated
            
                    if lect_column == "Lect 1":
                        row_data[2]=dtString

                    writer.writerow(row_data)
                               

#                 message = self.client.messages.create(
#                     to='+919370572782',
#                     from_=self.twilio_phone_number,
#                     body=f"Your child {name} with roll number {roll_no} was present in college today."
#             )

                    print(f"Attendance marked for {name} with roll number {roll_no} at {current_minutes}")
                else:
                    print(f"{name} already marked present today.")
 #elif lect_column =="Lect 2":
 #row_data[3]=dtString    


        else:
            lect_column="Lect 2"
            with open(os.path.join(self.attendance_folder, f'{self.current_date}.csv'), 'a+', newline='') as f:
                f.seek(0)
                myDataList = f.readlines()
                nameList = []

                if len(myDataList) > 0:
                    for line in myDataList:
                        entry = line.split(',')
                        if entry[0] == roll_no and entry[2].strip()[:8] == current_time[:8]:
                    # Attendance for the same student within the same hour already marked
                            print(f"Attendance for {name} with roll number {roll_no} already marked in the current hour.")
                            return

                        nameList.append(entry[1].strip())
            
                if name not in nameList:
                    now = datetime.now()
                    dtString = now.strftime('%H:%M:%S')
                    writer = csv.writer(f)

                    if f.tell() == 0:
                        writer.writerow(["Roll No.", "Name", "Lect 1", "Lect 2"])

                    row_data = [roll_no, name, "",""]

            # Add empty column values for the column that is not being updated
            
                    if lect_column == "Lect 2":
                        row_data[3]=dtString

                    writer.writerow(row_data)
                               

#                 message = self.client.messages.create(
#                     to='+919370572782',
#                     from_=self.twilio_phone_number,
#                     body=f"Your child {name} with roll number {roll_no} was present in college today."
#             )

                    print(f"Attendance marked for {name} with roll number {roll_no} at {current_minutes}")
            #else:
                #print(f"{name} already marked present today.")

            

                if name in nameList:
                #if hour >= 9:
                    print(hour)
                    now = datetime.now()
                    dtString = now.strftime('%H:%M:%S')
                    writer = csv.writer(f)

                    if f.tell() == 0:
                        writer.writerow(["Roll No.", "Name", "Lect 1", "Lect 2"])

                    csv_file_path = os.path.join(self.attendance_folder, f'{self.current_date}.csv')

                    with open(os.path.join(self.attendance_folder, f'{self.current_date}.csv'), 'r', newline='') as file:
                        csv_data = list(csv.reader(file))
        
        # Iterate through rows and find the matching student
                        for row in csv_data[1:]:
                            if row[0] == roll_no and row[1] == name:
                # Check if Lecture 2 is already marked
                                if row[3]:
                                    print(f"Lecture 2 attendance for {name} with roll number {roll_no} already marked")
                                    return

                # Update the Lecture 2 time in the CSV data
                                row[3] = dtString
                                break

                    with open(csv_file_path, 'w', newline='') as file:
                        writer = csv.writer(file)
                        writer.writerows(csv_data)

                    print(f"Lecture 2 attendance marked for {name} with roll number {roll_no} at {dtString}")

                               

                    

    def face_detector(self):
        cap = cv2.VideoCapture(0)

        while True:
            success, img = cap.read()
            imgS = cv2.resize(img, (0, 0), None, 0.25, 0.25)
            imgS = cv2.cvtColor(imgS, cv2.COLOR_BGR2RGB)

            facesCurFrame = face_recognition.face_locations(imgS)
            encodesCurFrame = face_recognition.face_encodings(imgS, facesCurFrame)

            for encodeFace, faceLoc in zip(encodesCurFrame, facesCurFrame):
                matches = face_recognition.compare_faces(self.encodeListKnown, encodeFace)
                faceDis = face_recognition.face_distance(self.encodeListKnown, encodeFace)
                matchIndex = np.argmin(faceDis)

                if matches[matchIndex]:
                    name = self.classNames[matchIndex]
                    roll_no = name.split('_')[0]
                    name = name.split('_')[1]
                    name = name.upper()
                    y1, x2, y2, x1 = faceLoc
                    y1, x2, y2, x1 = y1 * 4, x2 * 4, y2 * 4, x1 * 4
                    cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), 2)
                    cv2.rectangle(img, (x1, y2 - 35), (x2, y2), (0, 255, 0), cv2.FILLED)
                    cv2.putText(img, name, (x1 + 6, y2 - 6), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 2)

                    self.markAttendance(roll_no, name)

                    # Write data to Excel worksheet
                    #self.worksheet.write(self.row, 0, roll_no)
                    #self.worksheet.write(self.row, 1, name)
                    #self.worksheet.write(self.row, 2, datetime.now().strftime('%H:%M:%S'))
                    self.row += 1

                else:
                    # Unknown face
                    name = 'Unknown'
                    roll_no = 'N/A'
                    y1, x2, y2, x1 = faceLoc
                    y1, x2, y2, x1 = y1 * 4, x2 * 4, y2 * 4, x1 * 4
                    cv2.rectangle(img, (x1, y1), (x2, y2), (0, 0, 255), 2)
                    cv2.rectangle(img, (x1, y2 - 35), (x2, y2), (0, 0, 255), cv2.FILLED)
                    cv2.putText(img, name, (x1 + 6, y2 - 6), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 2)

            cv2.imshow('Webcam', img)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        cap.release()
        cv2.destroyAllWindows()

    def view_attendance(self):
        #os.startfile(r'C:\Users\HP')
        os.startfile(os.path.abspath(self.attendance_folder))

    def student_details(self):
        os.startfile('D:\myenv\Training_images')
        
    def exit_program(self):
        #self.workbook.close()
        self.root.destroy()


root = Tk()
obj = Face_Recognition_System(root)
root.mainloop()