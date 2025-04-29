def validate_image(file):
    """Validate that the uploaded file is an image with an allowed extension."""
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}
    
    # Check if the file has a filename
    if file.filename == '':
        return False
    
    # Check if the file has an allowed extension
    ext = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else ''
    return ext in ALLOWED_EXTENSIONS
