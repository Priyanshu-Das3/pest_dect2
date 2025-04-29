from flask import Blueprint, request, jsonify
import cv2
import numpy as np
from app.model import PestDetectionModel
from app.validation import validate_image
from app.excel_integration import update_excel_data

main_bp = Blueprint('main', __name__)

# Initialize the model
model = PestDetectionModel()

@main_bp.route('/detect', methods=['POST'])
def detect_pests():
    """API endpoint for pest detection."""
    # Check if image was uploaded
    if 'image' not in request.files:
        return jsonify({'error': 'No image provided'}), 400
    
    image_file = request.files['image']
    
    # Validate the image
    if not validate_image(image_file):
        return jsonify({'error': 'Invalid image format'}), 400
    
    # Read and process the image
    image_bytes = image_file.read()
    nparr = np.frombuffer(image_bytes, np.uint8)
    image = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    
    # Detect pests
    detection_results = model.detect(image)
    
    # Update Excel with real-time detection data
    update_excel_data(detection_results)
    
    return jsonify({
        'success': True,
        'detections': detection_results
    })

@main_bp.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    return jsonify({'status': 'healthy'})
