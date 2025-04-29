from flask import Flask
from app.excel_integration import init_excel_connector

def create_app():
    app = Flask(__name__)
    
    # Initialize Excel connection for real-time updates
    # Only init Excel on local environment, not on render
    import os
    if os.environ.get('RENDER') != 'true':
        init_excel_connector()
    
    # Register blueprints
    from app.routes import main_bp
    app.register_blueprint(main_bp)
    
    return app
