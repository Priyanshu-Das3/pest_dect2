from app import create_app
from app.excel_integration import cleanup

app = create_app()

if __name__ == '__main__':
    try:
        print("Starting Pest Detection API with Excel integration...")
        app.run(host='0.0.0.0', port=5000, debug=False)
    finally:
        # Clean up resources when the app is shutting down
        cleanup()
