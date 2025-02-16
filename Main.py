#Khai báo các thư viện
import tkinter as tk 
from tkinter import Label 
import cv2 #Xử lý hình ảnh từ camera
import face_recognition #Nhận diện khuôn mặt.
import numpy as np #Tính toán
import os #Làm việc với file
from PIL import Image, ImageTk #Chuyển đổi ảnh 
from datetime import datetime #Lấy thời gian hiện tại.
import pandas as pd #: Xử lý file Excel chứa dữ liệu chấm công.
from unidecode import unidecode # Chuyển đổi chữ có dấu sang không dấu.

#Tạo khung trên khuôn mặt
def khung(x, y, z, frame, left, top, right, bottom, name):
    cv2.rectangle(frame, (left, top), (right, bottom), (x, y, z), 2)
    cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (x, y, z), 2)
#Tính khoảng cách Euclidean giữa hai điểm trên mặt
def euclidean_distance(point1, point2):
    return np.linalg.norm(np.array(point1) - np.array(point2))
#Kiểm tra đi trễ
def check_time(result):
    current_time = datetime.now().strftime("%H:%M")
    #Thời gian quy định
    preset_time = "7:00" 
    # So sánh thời gian hiện tại với thời gian quy định
    if current_time > preset_time:
        result = 1#Đi trễ
    return result
# Đọc và mã hóa khuôn mặt từ các ảnh trong kho dữ liệu
def load_known_faces(directory):
    known_face_encodings = []
    known_face_names = []

    for filename in os.listdir(directory):
        if filename.endswith(('.jpg', '.jpeg', '.png')):
            # Đọc ảnh và chuyển đổi sang RGB
            image_path = os.path.join(directory, filename)
            image = face_recognition.load_image_file(image_path)
            image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)

            # Mã hóa khuôn mặt
            face_encodings = face_recognition.face_encodings(image_rgb)
            if face_encodings:
                known_face_encodings.append(face_encodings[0])
                known_face_names.append(os.path.splitext(filename)[0])
    
    return known_face_encodings, known_face_names

#Hiển thị thông tin 
def TEXT(Text, name, X, COLOR, Y):
    SNL = df[name][X]
    EX = tk.Label(window, text=Text + str(int(SNL))  , font=("Arial", 25))
    EX.config(fg = COLOR)
    EX.place(x = 750, y = Y)
#Reset dữ liệu hàng tháng
def reset():
    current_day = datetime.today().strftime("%d")
    if(int(current_day) == 1):
        df.iloc[0, 1:] = 0
        df.iloc[1, 1:] = 0
        df.iloc[2, 1:] = 0
    df.to_excel("D:\\Face_Recognition\\pic\\BangChamCong.xlsx", index=False)

# Hàm cập nhật và hiển thị hình ảnh trong Tkinter
def update_frame():
    ret, frame = cap.read()
    
    if ret:
        # Giữ nguyên ảnh BGR từ OpenCV
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)  # Bỏ dòng này
        
        # Phát hiện các khuôn mặt trong ảnh từ camera
        face_locations = face_recognition.face_locations(frame)
        face_encodings = face_recognition.face_encodings(frame, face_locations)
        face_landmarks_list = face_recognition.face_landmarks(frame, face_locations)

        for (top, right, bottom, left), face_encoding, face_landmarks in zip(face_locations, face_encodings, face_landmarks_list):
            # So sánh khuôn mặt từ camera với các khuôn mặt đã biết trong kho dữ liệu
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
            name = "Unknown"
                
                
            face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
            best_match_index = np.argmin(face_distances)

            #xử lí 
            if matches[best_match_index]:
                #tên và khung 
                name = known_face_names[best_match_index]
                khung(0, 255, 0, frame, left, top, right, bottom, unidecode(name))
                #Tên
                Name = tk.Label(window, text="Tên: " + name + "                                                       ", font=("Arial", 25))
                Name.place(x = 750, y = 150)
                #Thời gian chấm công
                current_time = datetime.now().strftime("%H:%M")
                result = 0
                result = check_time(result)
                time = tk.Label(window, text="Thời Gian Chấm Công: " + current_time , font=("Arial", 25))
                time.place(x = 750, y = 200)
                if(result == 0):
                    TT = tk.Label(window, text="Đúng Giờ      " , font=("Arial", 25))
                    TT.config(fg = "green")
                    time.config(fg="green")
                else:
                    TT = tk.Label(window, text="Đi Trễ     " , font=("Arial", 25))
                    TT.config(fg = "red")
                    time.config(fg="red")
                TT.place(x = 750, y = 350) 
                #Dữ liệu
                current_date = datetime.now().strftime("%d-%m-%Y")
                if(df[name][2] != current_date):
                    df[name][2] = current_date
                    df[name][0] += 1
                    df[name][1] += result
                    df.to_excel("D:\\Face_Recognition\\pic\\BangChamCong.xlsx", index=False)
                else :
                    c = tk.Label(window, text= "Đã Chấm Công", font=("Arial", 25))
                    c.config(fg = "green")
                    c.place(x = 750, y = 300)
                TEXT("Số Ngày Làm: ", name, 0, "green", 450)
                TEXT("Số Ngày Đi Trễ: ", name, 1, "red", 500)
            else:
                khung(255, 0, 0, frame, left, top, right, bottom, name)
        # Giảm kích thước hình ảnh (thay đổi kích thước theo tỷ lệ)
        img_pil = Image.fromarray(frame)
        img_resized = img_pil.resize((int(img_pil.width * 0.935), int(img_pil.height * 0.817)))
        # Chuyển đổi thành ImageTk để hiển thị trong Tkinter
        img_tk = ImageTk.PhotoImage(img_resized)
        label.img_tk = img_tk  # Giữ đối tượng ảnh để tránh thu gom rác
        label.config(image=img_tk)
    # Gọi lại hàm sau mỗi 10ms để tiếp tục cập nhật hình ảnh
    label.after(10, update_frame)
# Đường dẫn đến kho dữ liệu ảnh
known_faces_directory = 'D:/Face_Recognition/pic'
# Tải và mã hóa khuôn mặt từ kho dữ liệu
known_face_encodings, known_face_names = load_known_faces(known_faces_directory)

# Tạo cửa sổ Tkinter
window = tk.Tk()
window.title("Face Recognition")
window.geometry("1200x600")

# Mở camera
cap = cv2.VideoCapture(0)


background_image = Image.open("D:/Face_Recognition/pic/background.jpg")  # Đổi đường dẫn đến hình nền của bạn
background_image = background_image.resize((1200, 600), Image.Resampling.LANCZOS)  # Resize để khớp với cửa sổ
bg_image_tk = ImageTk.PhotoImage(background_image)

# Thêm background vào cửa sổ
background_label = Label(window, image=bg_image_tk)
background_label.place(x=0, y=0, relwidth=1, relheight=1)  # Đặt full màn hình

# Đảm bảo các widget quan trọng xuất hiện trên background
label = Label(window, bd=0, highlightthickness=0)
label.place(x=50, y=90)
    
# Đọc dữ liệu từ file Excel
file_path = 'D:\Face_Recognition\pic\BangChamCong.xlsx'  # Đường dẫn đến file Excel
df = pd.read_excel(file_path)  # Đọc toàn bộ dữ liệu từ sheet đầu tiên
    
reset()
    
# Bắt đầu cập nhật hình ảnh từ camera
update_frame()

# Bắt đầu vòng lặp chính của Tkinter
window.mainloop()

# Giải phóng tài nguyên khi cửa sổ đóng
cap.release()
cv2.destroyAllWindows()
