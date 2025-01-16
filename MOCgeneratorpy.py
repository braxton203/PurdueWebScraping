import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext

# Load the Excel file into a DataFrame
file_path = "PurdueTest.xlsx"  # Update with file path
df = pd.read_excel(file_path, index_col=0)  # Assuming chemicals are in the first column

# Function to display the compatibility table
def show_table(chemical):
    if chemical in df.index:
        # Extract the row corresponding to the selected chemical
        data = df.loc[[chemical]]
        
        # Convert the row to a nicely formatted string
        table_text = data.to_string(index=False)
        
        # Clear the text area and display the table
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, f"Compatibility of {chemical} with materials:\n\n{table_text}")
    else:
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, f"Chemical {chemical} not found in the data.")

# Set up the GUI
def create_dropdown():
    root = tk.Tk()
    root.title("Chemical Compatibility Checker")
    
    # Label
    label = tk.Label(root, text="Select a chemical:")
    label.pack(pady=10)
    
    # Create a dropdown list of chemicals
    chemical_var = tk.StringVar()
    chemical_dropdown = ttk.Combobox(root, textvariable=chemical_var)
    chemical_dropdown['values'] = df.index.tolist()  # Populate dropdown with chemical names
    chemical_dropdown.pack(pady=10)
    
    # Create a button to show the compatibility table
    def on_button_click():
        selected_chemical = chemical_var.get()
        show_table(selected_chemical)
    
    button = tk.Button(root, text="Show Compatibility", command=on_button_click)
    button.pack(pady=10)
    
    # Text area to display the table
    global text_area
    text_area = scrolledtext.ScrolledText(root, width=100, height=20, wrap=tk.WORD)
    text_area.pack(pady=10)
    
    root.mainloop()

# Run the dropdown interface
create_dropdown()
