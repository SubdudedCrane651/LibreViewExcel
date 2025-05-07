import pandas as pd
import xlwings as xw
import tkinter as tk
from tkinter import filedialog, messagebox
import json
import os  # Needed to open the file automatically
import sys

def process_csv():
    # Get the script's directory (useful for EXE packaging)
    # script_dir = os.path.dirname(os.path.abspath(__file__))
    # json_file_path = os.path.join(script_dir, "Libre2excel.json")
    
    # Get the script or executable directory
    script_dir = os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))

    # Ensure we exclude `.venv` in the path
    if ".venv" in script_dir:
        script_dir = os.path.dirname(script_dir)  # Move one level up

    json_file_path = os.path.join(script_dir, "Libre2excel.json")

    print("Filtered default directory:", script_dir)
    print("JSON file path:", json_file_path)


    # Load JSON paths dynamically
    with open(json_file_path, "r") as json_file:
        save_paths = json.load(json_file)
        
    # Extract paths
    excel_file = save_paths["excel_file"]
    image_file = save_paths["image_file"]
    json_file = save_paths["json_file"]
    
    print("Loaded Paths:", save_paths)
    
    # Open file dialog to select CSV file
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])

    if not file_path:
        messagebox.showerror("Error", "No file selected!")
        return

    # Load the CSV file, skipping metadata rows
    df = pd.read_csv(file_path, encoding="utf-8-sig", skiprows=1)

    # Strip any leading/trailing spaces in column names
    df.columns = df.columns.str.strip()

    # Convert timestamps to datetime format
    df['Device Timestamp'] = pd.to_datetime(df['Device Timestamp'], errors='coerce', dayfirst=True)

    # Extract the date for grouping
    df['Date'] = df['Device Timestamp'].dt.date

    # Define time categorization function
    def categorize_time(timestamp):
        if pd.isna(timestamp): 
            return None
        hour = timestamp.hour
        if 1 <= hour < 9:
            return "Morning (1:00-9:00)"
        elif 9 <= hour < 12:
            return "Before Lunch (9:01-12:00)"
        elif 12 <= hour < 18:
            return "Before Dinner (12:01-18:00)"
        elif 18 <= hour < 24:
            return "Evening (18:01-24:00)"
        return None

    # Apply time categorization
    df['Time Period'] = df['Device Timestamp'].apply(categorize_time)

    # Filter relevant columns
    df = df[['Date', 'Time Period', 'Historic Glucose mmol/L']]

    # Calculate average glucose values for each day & time period
    pivot_df = df.groupby(['Date', 'Time Period'])['Historic Glucose mmol/L'].mean().reset_index()

    # Reshape data into columns for each time period per day
    final_df = pivot_df.pivot(index='Date', columns='Time Period', values='Historic Glucose mmol/L')

    # Reset index to ensure the Date column is retained
    final_df.reset_index(inplace=True)

    # Sort columns in the correct order
    ordered_columns = ["Date", "Morning (1:00-9:00)", "Before Lunch (9:01-12:00)", "Before Dinner (12:01-18:00)", "Evening (18:01-24:00)"]
    final_df = final_df.reindex(columns=ordered_columns)

    # Format numbers to one decimal place
    final_df = final_df.round(1)

    # Open existing macro-enabled Excel file
    output_file = excel_file
    try:
        wb = xw.Book(output_file)  # Open existing workbook
    except FileNotFoundError:
        messagebox.showerror("Error", f"File '{output_file}' not found! Please create it first.")
        return

    sheet = wb.sheets[0]  # Select first sheet

    # Clear previous data
    sheet.range("A1").expand().clear_contents()

    # Write new DataFrame to Excel
    sheet.range("A1").value = final_df

    # Save and close
    wb.save(output_file)
    #wb.close()
    
    app = wb.app
    app.quit()  # Fully closes Excel if no other workbooks are open

    # Show completion message
    messagebox.showinfo("Success", f"Data successfully updated in {output_file} and opened!")
    
    # Open the Excel file automatically
    os.startfile(output_file)

# Create GUI window
root = tk.Tk()
root.title("LibreView CSV Processor")
root.geometry("400x200")

# Create a button to process the file
btn_select = tk.Button(root, text="Select CSV File", command=process_csv)
btn_select.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()