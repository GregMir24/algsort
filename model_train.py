from ultralytics import YOLO
import torch

model = YOLO('yolo12n.pt')
results = model.train(data='dataset/data.yaml', epochs=20,imgsz = 640)