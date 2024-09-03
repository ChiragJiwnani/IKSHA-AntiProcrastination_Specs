#RUN ON LOCAL
from ultralytics import YOLO
#from ultralytics.yolo.v8.detect.predict import DetectionPredictor
import cv2
from PIL import Image
model = YOLO("D:/Vibehai/IPD/YOLO v8/best3.pt")

# model = YOLO("C:/Users/dhruv/Downloads/weights-20230821T152722Z-001/weights/best.pt")
# accepts all formats - image/dir/Path/URL/video/PIL/ndarray. 0 for webcam
results = model.predict(source="0", show = True)

print(results)
