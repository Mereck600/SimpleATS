import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Function to load the Excel file
def load_excel_file():
    global df  # Ensure we're updating the global df variable
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path)  # This updates the global df variable
            display_data()  # Display the data immediately after loading
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {e}")
            df = None  # Ensure df is None if an error occurs
    else:
        df = None  # If no file is selected, set df to None

# Function to filter data based on user inputs
def filter_data():
    if df is None:
        messagebox.showerror("Error", "No data loaded.")
        return

    name_filter = name_entry.get()
    location_filter = location_entry.get()
    expertise_filter = expertise_entry.get()
    city_filter = city_entry.get()

    # Apply filters on the dataframe
    filtered_data = df.copy()

    if name_filter:
        filtered_data = filtered_data[filtered_data['Name'].str.contains(name_filter, case=False, na=False)]
    if location_filter:
        filtered_data = filtered_data[filtered_data['State'].str.contains(location_filter, case=False, na=False)]
    if expertise_filter:
        filtered_data = filtered_data[filtered_data['Function'].str.contains(expertise_filter, case=False, na=False)]
    if city_filter:
        filtered_data = filtered_data[filtered_data['City'].str.contains(city_filter, case=False, na=False)]

    # Display filtered results in the Treeview
    for row in tree.get_children():
        tree.delete(row)

    for _, row in filtered_data.iterrows():
        tree.insert("", "end", values=row.tolist())

# Function to display data in the Treeview widget
def display_data():
    if df is None:
        messagebox.showerror("Error", "No data loaded.")
        return

    # Clear previous entries in the Treeview
    for row in tree.get_children():
        tree.delete(row)

    # Insert all rows into the Treeview
    for _, row in df.iterrows():
        tree.insert("", "end", values=row.tolist())

# Initialize Tkinter window
root = tk.Tk()
root.title("Mereck's Genius ATS System")
root.configure(bg='lightblue')
# Create GUI components
frame = ttk.Frame(root)
frame.pack(padx=20, pady=20)

# Load Button
load_button = ttk.Button(frame, text="Load Excel File", command=load_excel_file)
load_button.grid(row=0, column=0, pady=10)

# Filters Section
ttk.Label(frame, text="Name Filter:").grid(row=1, column=0, sticky="w", pady=5)
name_entry = ttk.Entry(frame)
name_entry.grid(row=1, column=1, pady=5)

ttk.Label(frame, text="State Filter :").grid(row=2, column=0, sticky="w", pady=5)
location_entry = ttk.Entry(frame)
location_entry.grid(row=2, column=1, pady=5)

ttk.Label(frame, text="Expertise Filter:").grid(row=3, column=0, sticky="w", pady=5)
expertise_entry = ttk.Entry(frame)
expertise_entry.grid(row=3, column=1, pady=5)

ttk.Label(frame, text="City Filter:").grid(row=4, column=0, sticky="w", pady=5)
city_entry = ttk.Entry(frame)
city_entry.grid(row=4, column=1, pady=5)

# Filter Button
filter_button = ttk.Button(frame, text="Apply Filters", command=filter_data)
filter_button.grid(row=5, column=0, columnspan=2, pady=10)

# Display Button
display_button = ttk.Button(frame, text="Display All Data", command=display_data)
display_button.grid(row=6, column=0, columnspan=2, pady=10)

# Frame to contain the Treeview and Scrollbar
tree_frame = ttk.Frame(root)
tree_frame.pack(padx=20, pady=20)

# Columns for the Treeview, adding Phone and City columns
columns = ['Name', 'Expertise', 'Company','Phone', 'City' ]  # Swap City and Phone columns

# Create the Treeview widget
tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
tree.pack(side="left", fill="both", expand=True)

# Add scrollbar to the Treeview
scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
scrollbar.pack(side="right", fill="y")

# Attach the scrollbar to the Treeview
tree.configure(yscrollcommand=scrollbar.set)

# Set the headings for the columns
for col in columns:
    tree.heading(col, text=col)

# Global variable to store dataframe
df = None

root.mainloop()
