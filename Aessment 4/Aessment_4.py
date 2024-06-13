import tkinter as tk
from tkinter import messagebox
import pyodbc
from datetime import datetime

# Connect to the MS Access database
try:
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\2022000130\Downloads\Company_Data.accdb;') 
    cursor = conn.cursor()
except Exception as e:
    messagebox.showerror("Connection Error", f"Failed to connect to the database: {e}")
    exit()

# Function to print all records
def print_all_records():
    cursor.execute("SELECT * FROM Company_data1")
    records = cursor.fetchall()
    display_records(records)

# Function to display records in the text widget
def display_records(records):
    text_output.delete("1.0", tk.END)
    for i, record in enumerate(records, start=1):
        text_output.insert(tk.END, f"Record {i}:\n")
        for value in record:
            text_output.insert(tk.END, f"  {value}\n")
        text_output.insert(tk.END, "\n")

# Function to print records with positive revenue growth
def print_positive_growth():
    cursor.execute("SELECT * FROM Company_data1 WHERE revenue_growth > 0")
    records = cursor.fetchall()
    display_records(records)

# Function to query record by date
def query_record_by_date():
    input_date = entry_date.get()
    try:
        datetime.strptime(input_date, "%Y-%m-%d")  # Validate date format
        cursor.execute(f"SELECT * FROM Company_data1 WHERE date = #{input_date}#")
        records = cursor.fetchall()
        if records:
            display_records(records)
        else:
            messagebox.showinfo("No Records", "No matching record found for the entered date.")
    except ValueError:
        messagebox.showerror("Invalid Date", "Please enter a valid date in YYYY-MM-DD format.")

# Function to count companies established between two dates
def count_companies_between_dates():
    start_date_str = entry_start_date.get()
    end_date_str = entry_end_date.get()
    try:
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d")

        cursor.execute(f"""
            SELECT * FROM Company_data1
            WHERE established_date BETWEEN #{start_date_str}# AND #{end_date_str}#
        """)
        
        records = cursor.fetchall()
        display_records(records)
    except ValueError:
        messagebox.showerror("Invalid Date", "Please enter valid dates in YYYY-MM-DD format.")

# Function to quit the program
def quit_program():
    root.destroy()
    exit()

# GUI development

# Create the main window
root = tk.Tk()
root.title("Company Data Management")

# Create and place widgets
frame = tk.Frame(root)
frame.pack(pady=10)

btn_print_all = tk.Button(frame, text="Print All Records", command=print_all_records)
btn_print_all.grid(row=0, column=0, padx=5, pady=5)

btn_positive_growth = tk.Button(frame, text="Print Positive Growth", command=print_positive_growth)
btn_positive_growth.grid(row=0, column=1, padx=5, pady=5)

label_date = tk.Label(frame, text="Date (YYYY-MM-DD):")
label_date.grid(row=1, column=0, padx=5, pady=5)
entry_date = tk.Entry(frame)
entry_date.grid(row=1, column=1, padx=5, pady=5)
btn_query_date = tk.Button(frame, text="Query by Date", command=query_record_by_date)
btn_query_date.grid(row=1, column=2, padx=5, pady=5)

label_start_date = tk.Label(frame, text="Start Date (YYYY-MM-DD):")
label_start_date.grid(row=2, column=0, padx=5, pady=5)
entry_start_date = tk.Entry(frame)
entry_start_date.grid(row=2, column=1, padx=5, pady=5)

label_end_date = tk.Label(frame, text="End Date (YYYY-MM-DD):")
label_end_date.grid(row=2, column=2, padx=5, pady=5)
entry_end_date = tk.Entry(frame)
entry_end_date.grid(row=2, column=3, padx=5, pady=5)

btn_count_dates = tk.Button(frame, text="Count Companies Between Dates", command=count_companies_between_dates)
btn_count_dates.grid(row=2, column=4, padx=5, pady=5)

btn_quit = tk.Button(frame, text="Quit", command=quit_program)
btn_quit.grid(row=3, columnspan=5, pady=5)

text_output = tk.Text(root, height=20, width=80)
text_output.pack(pady=10)

# Run the application
root.mainloop()
