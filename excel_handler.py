import os
import sys
import xlwings as xw
import pandas as pd
from datetime import datetime
import time

class ExcelRealTimeInterface:
    def __init__(self):
        self.excel_path = self._get_excel_path()
        self.app = None
        self.wb = None
    
    def _get_excel_path(self):
        """Get the path to the Excel file."""
        documents_path = os.path.expanduser("~/Documents")
        return os.path.join(documents_path, "pest_detection_data.xlsx")
    
    def start(self):
        """Start the Excel interface."""
        try:
            # Initialize Excel application
            self.app = xw.App(visible=True)
            
            # Open or create workbook
            if os.path.exists(self.excel_path):
                self.wb = self.app.books.open(self.excel_path)
            else:
                self._create_new_workbook()
            
            # Add formulas for real-time updates
            self._setup_real_time_dashboard()
            
            print(f"Excel interface running. File: {self.excel_path}")
            return True
        
        except Exception as e:
            print(f"Error starting Excel interface: {e}")
            return False
    
    def _create_new_workbook(self):
        """Create a new Excel workbook with required structure."""
        self.wb = self.app.books.add()
        
        # Create the data sheet
        pest_types = [
            "Aphid", "Armyworm", "Beetle", "Bollworm", "Grasshopper", 
            "Leafhopper", "Mite", "Mosquito", "Stem Borer", "Thrips"
        ]
        
        data = {
            'Pest Type': pest_types,
            'Count': [0] * len(pest_types),
            'Last Updated': [''] * len(pest_types),
            'Location': [''] * len(pest_types)
        }
        
        df = pd.DataFrame(data)
        
        # Create data sheet
        data_sheet = self.wb.sheets.add('Pest Detection Data')
        data_sheet.range('A1').options(index=False).value = df
        
        # Create visualization sheet
        viz_sheet = self.wb.sheets.add('Visualization')
        viz_sheet.range('A1').options(index=False).value = df[['Pest Type', 'Count']]
        
        # Save the workbook
        self.wb.save(self.excel_path)
    
    def _setup_real_time_dashboard(self):
        """Set up the dashboard sheet with real-time formulas."""
        # Check if Dashboard sheet exists, create if not
        try:
            dashboard = self.wb.sheets['Dashboard']
        except:
            dashboard = self.wb.sheets.add('Dashboard')
        
        # Clear the dashboard
        dashboard.clear()
        
        # Add title
        dashboard.range('A1').value = "Real-Time Pest Detection Dashboard"
        dashboard.range('A1').font.size = 16
        dashboard.range('A1').font.bold = True
        
        # Add a refresh timestamp cell
        dashboard.range('A3').value = "Last Refresh:"
        dashboard.range('B3').formula = "=NOW()"
        dashboard.range('B3').number_format = "yyyy-mm-dd hh:mm:ss"
        
        # Add table headers
        dashboard.range('A5').value = [['Pest Type', 'Count', 'Last Updated']]
        dashboard.range('A5:C5').font.bold = True
        
        # Add data from the data sheet with INDIRECT formula for real-time updates
        # This creates a reference to the data sheet that updates automatically
        sheet_name = "'Pest Detection Data'"
        
        # Get pest types from data sheet
        data_sheet = self.wb.sheets['Pest Detection Data']
        pest_types = data_sheet.range('A2').expand('down').value
        
        if not isinstance(pest_types, list):
            pest_types = [pest_types]
        
        # Add dynamic references to each pest type
        for i, pest in enumerate(pest_types):
            row = 6 + i
            dashboard.range(f'A{row}').value = pest
            
            # Use INDIRECT to create dynamic references to the data sheet
            count_cell = f'INDIRECT("{sheet_name}!B{i+2}")'
            updated_cell = f'INDIRECT("{sheet_name}!C{i+2}")'
            
            dashboard.range(f'B{row}').formula = f"={count_cell}"
            dashboard.range(f'C{row}').formula = f"={updated_cell}"
        
        # Add a chart
        chart_row = 6 + len(pest_types) + 2
        dashboard.range(f'A{chart_row}').value = "Pest Detection Count"
        
        # Create a chart from the data
        chart = dashboard.charts.add()
        chart.chart_type = 'column'
        chart.set_source_data(dashboard.range(f'A6:B{5+len(pest_types)}'))
        chart.left = dashboard.range(f'D6').left
        chart.top = dashboard.range(f'D6').top
        chart.width = 400
        chart.height = 300
        
        # Format the chart
        chart.api[1].SetElement(2)  # Show title
        chart.api[1].ChartTitle.Text = "Pest Counts"
        
        # Auto-fit columns
        dashboard.autofit()
        
        # Save the workbook
        self.wb.save()
    
    def stop(self):
        """Stop the Excel interface."""
        if self.wb:
            self.wb.save()
            self.wb.close()
        if self.app:
            self.app.quit()

def main():
    # Create and start the Excel interface
    interface = ExcelRealTimeInterface()
    if interface.start():
        try:
            # Keep the script running
            print("Press Ctrl+C to exit...")
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("Stopping Excel interface...")
        finally:
            interface.stop()

if __name__ == "__main__":
    main()
