import os
import pickle
import cv2
import numpy as np
from ultralytics import YOLO
import pandas as pd
from datetime import datetime

# Path to the pest detection model
MODEL_PATH = os.path.join(os.path.dirname(__file__), 'pest_detection_model_2.pkl')

# Class names for detected pests
CLASS_NAMES = [
    "Aphid", "Armyworm", "Beetle", "Bollworm", "Grasshopper", 
    "Leafhopper", "Mite", "Mosquito", "Stem Borer", "Thrips"
]

class PestDetector:
    def __init__(self, model_path=MODEL_PATH):
        # Try to load pickled model, fallback to YOLO
        try:
            with open(model_path, 'rb') as file:
                self.model = pickle.load(file)
            self.model_type = "sklearn"
        except:
            # Use YOLO instead of TensorFlow as requested
            self.model = YOLO('yolov8n.pt')
            self.model_type = "yolo"
    
    def preprocess_image(self, image):
        """Preprocess image based on model type."""
        if self.model_type == "sklearn":
            # Convert to grayscale and resize for sklearn model
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            resized = cv2.resize(gray, (128, 128))
            flattened = resized.flatten() / 255.0
            return flattened
        else:
            # YOLO doesn't need preprocessing
            return image
    
    def detect(self, image):
        """Detect pests in the image."""
        processed_image = self.preprocess_image(image)
        
        if self.model_type == "sklearn":
            # Use sklearn model prediction
            prediction = self.model.predict([processed_image])
            result = {
                CLASS_NAMES[prediction[0]]: 1
            }
        else:
            # Use YOLO for detection
            results = self.model(image)
            
            # Process YOLO results
            result = {}
            for r in results:
                boxes = r.boxes
                for box in boxes:
                    cls = int(box.cls[0])
                    conf = float(box.conf[0])
                    
                    # Only include confident detections
                    if conf > 0.5 and cls < len(CLASS_NAMES):
                        pest_name = CLASS_NAMES[cls]
                        if pest_name in result:
                            result[pest_name] += 1
                        else:
                            result[pest_name] = 1
        
        return result
    
    def update_excel(self, detections, excel_path):
        """Update Excel with detection results in real-time."""
        try:
            # Read existing Excel file or create new one
            if os.path.exists(excel_path):
                df = pd.read_excel(excel_path, sheet_name='Pest Detection Data')
            else:
                # Create new DataFrame if file doesn't exist
                df = pd.DataFrame({
                    'Pest Type': CLASS_NAMES,
                    'Count': [0] * len(CLASS_NAMES),
                    'Last Updated': [''] * len(CLASS_NAMES),
                    'Location': [''] * len(CLASS_NAMES)
                })
            
            # Update counts and timestamps
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for pest, count in detections.items():
                # Find the row for this pest
                idx = df[df['Pest Type'] == pest].index
                if len(idx) > 0:
                    # Update existing pest
                    df.at[idx[0], 'Count'] += count
                    df.at[idx[0], 'Last Updated'] = timestamp
                else:
                    # Add new pest if not in the list
                    new_row = pd.DataFrame({
                        'Pest Type': [pest],
                        'Count': [count],
                        'Last Updated': [timestamp],
                        'Location': ['']
                    })
                    df = pd.concat([df, new_row], ignore_index=True)
            
            # Save to Excel with multiple sheets
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Pest Detection Data', index=False)
                
                # Create visualization data
                df[['Pest Type', 'Count']].to_excel(writer, sheet_name='Visualization', index=False)
            
            return True
        
        except Exception as e:
            print(f"Error updating Excel: {e}")
            return False


def process_video_stream(camera_index=0, excel_path=None):
    """Process video stream for real-time pest detection."""
    # Set default Excel path if not provided
    if not excel_path:
        documents_path = os.path.expanduser("~/Documents")
        excel_path = os.path.join(documents_path, "pest_detection_data.xlsx")
    
    # Initialize detector
    detector = PestDetector()
    
    # Open webcam
    cap = cv2.VideoCapture(camera_index)
    if not cap.isOpened():
        print("Error: Could not open camera")
        return
    
    # Process frames
    frame_count = 0
    try:
        while True:
            # Read frame
            ret, frame = cap.read()
            if not ret:
                break
            
            # Process every 10 frames to reduce load
            if frame_count % 10 == 0:
                # Detect pests
                detections = detector.detect(frame)
                
                # Update Excel if pests detected
                if detections:
                    detector.update_excel(detections, excel_path)
                    
                    # Draw detection results on frame
                    for pest, count in detections.items():
                        cv2.putText(
                            frame, f"{pest}: {count}", (10, 30), 
                            cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2
                        )
            
            # Show frame with detections
            cv2.imshow("Pest Detection", frame)
            
            # Exit on 'q' key
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
            
            frame_count += 1
    
    finally:
        # Release resources
        cap.release()
        cv2.destroyAllWindows()


if __name__ == "__main__":
    import argparse
    
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Real-time Pest Detection")
    parser.add_argument("--camera", type=int, default=0, help="Camera index to use")
    parser.add_argument("--excel", type=str, help="Path to Excel file for results")
    
    args = parser.parse_args()
    
    # Start real-time detection
    process_video_stream(args.camera, args.excel)
