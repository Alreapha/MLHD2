import tkinter as tk
from tkinter import ttk
import pandas as pd

def main():
    # Create the main window
    root = tk.Tk()
    root.title("Excel Data Viewer")
    root.geometry("1980x1000")
    
    # Configure root grid layout
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    
    # Create a frame for the table
    table_frame = tk.Frame(root)
    table_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
    
    # Configure table_frame grid
    table_frame.grid_rowconfigure(0, weight=1)
    table_frame.grid_columnconfigure(0, weight=1)
    
    try:
        # Load data from Excel file
        df = pd.read_excel("mission_log_test.xlsx")
        
        # Get column names from the DataFrame
        columns = list(df.columns)
        
        # Create the table with ttk.Treeview
        table = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        
        # Define the column headings based on Excel columns
        for col in columns:
            table.heading(col, text=col)
            table.column(col, width=100, anchor=tk.CENTER)
        
        # Add data from Excel
        for _, row in df.iterrows():
            values = [str(val) if pd.notna(val) else "" for val in row]
            table.insert("", tk.END, values=values)
            
    except Exception as e:
        print(f"Error loading Excel file: {e}")
    
    # Add scrollbars
    y_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=table.yview)
    x_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=table.xview)
    table.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
    
    # Grid layout for the table and scrollbars
    table.grid(row=0, column=0, sticky="nsew")
    y_scrollbar.grid(row=0, column=1, sticky="ns")
    x_scrollbar.grid(row=1, column=0, sticky="ew")
    
    # Create a button frame
    button_frame = tk.Frame(root)
    button_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
    
    # Add a button
    def button_action():
        selected = table.selection()
        if selected:
            selected_item = table.item(selected, "values")
            print(f"Selected: {selected_item}")
        else:
            print("No item selected")
    
    button = tk.Button(button_frame, text="Process Selection", command=button_action)
    button.pack(side=tk.RIGHT)
    
    # Run the main event loop
    root.mainloop()

if __name__ == "__main__":
    main()