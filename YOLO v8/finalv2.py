import pandas as pd
import time
from datetime import datetime
import cv2
from ultralytics import YOLO

# ... Other parts of your code ...

URL = 0
excel_file_path = "object_detection.xlsx"
model_path = "D:/Vibehai/IPD/YOLO v8/best3.pt"
model = YOLO(model_path)
cap = cv2.VideoCapture(URL)
class_list = ['Instagram', 'None', 'Book']
current_datetime = datetime.now()

# Format the date as "dd mm yy"
formatted_date = current_datetime.strftime("%d %m %y")
# Format the time as "hh min sec"
formatted_time = current_datetime.strftime("%H %M %S")


def pred():

    start_time = time.time()  # Get the current time in seconds
    end_time = start_time + 3  # Calculate the end time (3 seconds from the start)
    list = []
    list_dict = []
    while time.time() < end_time:
        ret, frame = cap.read()

        if not ret:
            print("Obstruction in Input")

        # Perform object detection using YOLO
        results = model(frame)

        a=results[0].boxes.data
        px=pd.DataFrame(a).astype("float")
    #   
        for _,row in px.iterrows():
            current_datetime = datetime.now()

            # Format the date as "dd mm yy"
            formatted_date = current_datetime.strftime("%d %m %y")

            # Format the time as "hh min sec"
            formatted_time = current_datetime.strftime("%H %M %S")

    #       print(row)
            d=int(row[5])
            c=class_list[d]
            z = {"Task": c, "Time": formatted_time, "Date": formatted_date} #where to add
            if c:
                list_dict.append(z)

    if list_dict:
        return list_dict
    else:
        return "No Detections"
    
#     ret, frame = cap.read()

#     if not ret:
#         print("Obstruction in Input")

#     # Perform object detection using YOLO
#     results = model(frame)

#     a=results[0].boxes.data
#     px=pd.DataFrame(a).astype("float")
# #   
#     for index,row in px.iterrows():
# #       print(row)
#         d=int(row[5])
#         c=class_list[d]
#         list.append(c)
#     return list

            


# Create an empty DataFrame to store the data
data_log = pd.DataFrame(columns=["Time", "Returned String"])


# Load existing data from the Excel file
try:
    existing_data = pd.read_excel(excel_file_path)
except pd.errors.EmptyDataError:
    existing_data = pd.DataFrame(columns=["Task", "Time", "Date"])

try:
    while True:
        # Every 5 minutes
        time.sleep(3)
        returned_dict = pred()

        # ... Other parts of your code ...

        if returned_dict == 'No Detections':
            current_datetime = datetime.now()
            formatted_date = current_datetime.strftime("%d %m %y")
            formatted_time = current_datetime.strftime("%H %M %S")
            save_dict = [{"Task": "No Detections", "Time": formatted_time, "Date": formatted_date}]
        else:
            save_dict = returned_dict

        # Append the data to the existing DataFrame
        new_data = pd.DataFrame(save_dict)
        existing_data = existing_data._append(new_data, ignore_index=True)

        # Save the DataFrame to the Excel file, overwriting the old sheet
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            existing_data.to_excel(writer, index=False, sheet_name='Sheet1')

except KeyboardInterrupt:
    print("Data logging stopped.")
