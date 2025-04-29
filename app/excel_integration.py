import os
import threading
import time
import pandas as pd
from datetime import datetime

# Global variables for Excel integration
excel_file_path = None
detection_data = {}
update_thread = None
running = True

def init_excel_connector():
    """Initialize the Excel connection for real-time updates."""
    global excel_file_path, update_thread
    
    # Set path to the Excel file (create in user's documents folder by default)
    documents_path = os.path.expanduser("~/Documents")
    excel_file_path = os.path.join(documents_path, "pest_detection_data.xlsx")
    
    # Create Excel file if it doesn't exist
    if not os.path.exists(excel_file_path):
        create_excel_file()
    
    # Start the update thread
    update_thread = threading.Thread(target=excel_update_loop)
    update_thread.daemon = True
    update_thread.start()

def create_excel_file():
    """Create a new Excel file with the required structure."""
    # Create a DataFrame with pest types
    pest_types = [
        "Aphid", "Armyworm", "Beetle", "Bollworm", "Grasshopper", 
        "Leafhopper", "Mite", "Mosquito", "Stem Borer", "Thrips"
    ]
    
    # Create initial data
    data = {
        'Pest Type': pest_types,
        'Count': [0] * len(pest_types),
        'Last Updated': [''] * len(pest_types),
        'Location': [''] * len(pest_types)
    }
    
    df = pd.DataFrame(data)
    
    # Save to Excel (two sheets - one for data, one for visualization)
    with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Pest Detection Data', index=False)
        df[['Pest Type', 'Count']].to_excel(writer, sheet_name='Visualization', index=False)

def update_excel_data(new_detections, location="Default"):
    """Update the detection data with new detections."""
    global detection_data
    
    # Add timestamps to new detections
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Update counts
    for pest, count in new_detections.items():
        if pest in detection_data:
            detection_data[pest]['Count'] += count
        else:
            detection_data[pest] = {
                'Count': count, 
                'Last Updated': timestamp,
                'Location': location
            }
        
        # Update timestamp
        detection_data[pest]['Last Updated'] = timestamp

def excel_update_loop():
    """Background thread to periodically update the Excel file."""
    global running, detection_data, excel_file_path
    
    while running:
        # Only update if there's data and the excel file exists
        if detection_data and os.path.exists(excel_file_path):
            try:
                # Read the current Excel file
                df = pd.read_excel(excel_file_path, sheet_name='Pest Detection Data')
                
                # Update the DataFrame with new detection data
                for pest, data in detection_data.items():
                    # Find the row index for this pest
                    idx = df[df['Pest Type'] == pest].index
                    if len(idx) > 0:
                        df.at[idx[0], 'Count'] = data['Count']
                        df.at[idx[0], 'Last Updated'] = data['Last Updated']
                        if 'Location' in data:
                            df.at[idx[0], 'Location'] = data['Location']
                
                # Save the updated DataFrame back to Excel
                with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Pest Detection Data', index=False)
                    
                    # Update the visualization sheet
                    df[['Pest Type', 'Count']].to_excel(writer, sheet_name='Visualization', index=False)
            
            except Exception as e:
                print(f"Error updating Excel: {e}")
        
        # Sleep for a while before the next update
        time.sleep(2)

def cleanup():
    """Clean up resources when the application is shutting down."""
    global running, update_thread
    
    running = False
    if update_thread:
        update_thread.join(timeout=1)
