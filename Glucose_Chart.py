from datetime import datetime
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import os
import json
from mpl_toolkits.mplot3d import Axes3D  # Required for 3D plotting

# --------------------------
# Load Excel file and extract data
# --------------------------

# Get the script's directory (useful for EXE packaging)
script_dir = os.path.dirname(os.path.abspath(__file__))
json_file_path = os.path.join(script_dir, "Libre2excel.json")

# Load JSON paths dynamically
with open(json_file_path, "r") as json_file:
    save_paths = json.load(json_file)
        
    # Extract paths
excel_file = save_paths["excel_file"]
image_file = save_paths["image_file"]
json_file = save_paths["json_file"]
    
print("Loaded Paths:", save_paths)

# Define file paths
file_path = excel_file
image_path = image_file

# Read the Excel file, extracting data from Sheet2
df = pd.read_excel(file_path, sheet_name="Glycèmie De Richard Perreault", skiprows=3)
print(df.head())  # Verify if column A now contains actual data

print(df.columns)

# Extract date and glucose columns
dates_raw = df.loc[0:999, "Date"]  # A5:A1000
glucose_readings_raw = df.loc[0:999, "Unnamed: 5"]  # F5:F1000

print(glucose_readings_raw)

# Convert dates to proper datetime format
parsed_dates = []
for d in dates_raw:
    try:
        parsed_dates.append(datetime.strptime(str(d), "%Y-%m-%d %H:%M:%S"))  # Corrected format
    except Exception as e:
        print(f"Error parsing date '{d}': {e}")
        parsed_dates.append(None)
        
# Remove rows with missing values
filtered_data = [(date, glucose) for date, glucose in zip(parsed_dates, glucose_readings_raw) if date and pd.notna(glucose)]
sorted_data = sorted(filtered_data, key=lambda x: x[0])

print("Filtered Data:", filtered_data)

# Unzip sorted data
dates_sorted, glucose_sorted = zip(*sorted_data)

# --------------------------
# Define positions and properties for bars (3D)
# --------------------------
ind = np.arange(len(glucose_sorted))  # x positions, one per reading
width = 0.35   # width of each bar
depth = 0.5    # depth of each bar

# Determine conditional colors for each bar
colors = ['red' if val > 7 else 'blue' if val < 3 else 'green' for val in glucose_sorted]

# --------------------------
# Create a 3D bar chart
# --------------------------
fig = plt.figure()
ax = fig.add_subplot(111, projection='3d')

ax.bar3d(ind, 
         np.zeros_like(ind), 
         np.zeros_like(ind), 
         width, 
         depth, 
         glucose_sorted, 
         color=colors, 
         shade=True)

# --------------------------
# Setting custom x-axis ticks: One label per month in order
# --------------------------
tick_positions = []
tick_labels = []
last_month = None

for i, d in enumerate(dates_sorted):
    current_month = d.strftime('%Y-%m')
    if current_month != last_month:
        tick_positions.append(i + width / 2)
        tick_labels.append(d.strftime('%b %Y'))
        last_month = current_month

ax.set_xticks(tick_positions)
ax.set_xticklabels(tick_labels, rotation=45, fontsize=10)

# --------------------------
# Labeling and formatting axes
# --------------------------
ax.set_zlabel('Glucose Reading')
plt.title('Glucose Readings per Day (3D)')

# Hide dummy y-axis ticks as we have no meaningful y variable.
ax.set_yticks([])

plt.tight_layout()

# --------------------------
# Save the chart as a PNG file before showing it
# --------------------------
output_image = image_file
plt.savefig(output_image, dpi=300)
plt.show()

import win32com.client

# Connect to Excel application
excel = win32com.client.Dispatch("Excel.Application")

# Check if the file is already open
for wb in excel.Workbooks:
    if wb.FullName.lower().find(file_path.lower().split("/")[-1])!=-1:
        
        ws = wb.Sheets("Glycèmie De Richard Perreault")  # Select the correct sheet
        
        # **Remove existing images**
        for shape in ws.Shapes:
            if shape.Type == 13:  # 13 = Image type
                shape.Delete()
                
        image_path = os.path.normpath(image_file)
                
        # **Insert the new image at K27 with custom size**
        ws.Shapes.AddPicture(image_path, 1, 1, ws.Range("K27").Left, ws.Range("K27").Top, 640,400)  # Adjust size

        # Save the file (while keeping it open)
        wb.Save()
        print("Image inserted successfully!")
        break
else:
    print("Excel file is not open. Please open it first.")