import pickle
import cv2
import numpy as np
from ultralytics import YOLO
import os

# Path to the model file
MODEL_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'pest_detection_model_2.pkl')

class PestDetectionModel:
    def __init__(self):
        # Load the model
        try:
            with open(MODEL_PATH, 'rb') as file:
                self.model = pickle.load(file)
            self.model_type = "sklearn"
        except:
            # Fallback to YOLO if the pickled model fails
            self.model = YOLO('yolov8n.pt')
            self.model_type = "yolo"
        
        # Class names for detected pests
        self.class_names = [
            "Aphid", "Armyworm", "Beetle", "Bollworm", "Grasshopper", 
            "Leafhopper", "Mite", "Mosquito", "Stem Borer", "Thrips"
        ]
    
    def preprocess_image(self, image):
        """Preprocess the image for the model."""
        if self.model_type == "sklearn":
            # Convert to grayscale and resize
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            resized = cv2.resize(gray, (128, 128))
            # Flatten the image
            flattened = resized.flatten() / 255.0
            return flattened
        else:
            # YOLO model doesn't need preprocessing
            return image
    
    def detect(self, image):
        """Detect pests in the given image."""
        processed_image = self.preprocess_image(image)
        
        if self.model_type == "sklearn":
            # Make prediction with sklearn model
            prediction = self.model.predict([processed_image])
            result = {
                self.class_names[prediction[0]]: 1
            }
        else:
            # Make prediction with YOLO model
            results = self.model(image)
            
            # Process the results
            result = {}
            for r in results:
                boxes = r.boxes
                for box in boxes:
                    cls = int(box.cls[0])
                    conf = float(box.conf[0])
                    if cls < len(self.class_names) and conf > 0.5:
                        pest_name = self.class_names[cls]
                        if pest_name in result:
                            result[pest_name] += 1
                        else:
                            result[pest_name] = 1
        
        return result
