import tkinter as tk
from tkinter import ttk
import configparser
import pandas as pd
import threading
from tkinter import messagebox

# Read configuration from config.config
config = configparser.ConfigParser()
config.read('config.config')

#Constants
DEBUG = config.getboolean('DEBUGGING', 'DEBUG', fallback=False)


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
    
    # Create the table with ttk.Treeview
    try:
        # Determine which file to use
        if DEBUG:
            excel_file = "mission_log_test.xlsx"
        else:
            excel_file = "mission_log.xlsx"
        
        # Create a label to show loading status
        status_label = tk.Label(table_frame, text="Loading data...", font=("Arial", 12))
        status_label.grid(row=0, column=0, sticky="nsew")
        
        # Create the table structure (but don't load data yet)
        table = ttk.Treeview(table_frame, show="headings", selectmode="extended")
        
        # Function to load data in background
        def load_data():
            try:
                # Read only the header first to set up columns
                df_header = pd.read_excel(excel_file, nrows=0)
                columns = list(df_header.columns)
                
                # Configure columns on the main thread
                root.after(0, lambda: setup_columns(columns))
                
                # Read the entire Excel file
                chunk_size = 1000
                df = pd.read_excel(excel_file)
                
                chunk_count = 0
                total_rows = len(df)
                for i in range(0, total_rows, chunk_size):
                    chunk = df.iloc[i:i+chunk_size]
                    batch = []
                    for _, row in chunk.iterrows():
                        values = [str(val) if pd.notna(val) else "" for val in row]
                        batch.append(values)
                    
                    # Update UI in the main thread with this batch
                    chunk_count += 1
                    root.after(0, lambda b=batch, c=chunk_count: insert_batch(b, c, chunk_size))
                
                # Final UI update when complete
                root.after(0, lambda: status_label.grid_forget())
                
            except Exception as e:
                root.after(0, lambda e=e: show_error(f"Error loading Excel file: {e}"))
        
        def setup_columns(columns):
            table["columns"] = columns
            for col in columns:
                table.heading(col, text=col)
                table.column(col, width=100, anchor=tk.CENTER)
        
        def insert_batch(batch, chunk_num, chunk_size):
            for values in batch:
                table.insert("", tk.END, values=values)
            status_label.config(text=f"Loading data... (Loaded {chunk_num * chunk_size} rows)")
        
        def show_error(message):
            status_label.grid_forget()
            messagebox.showerror("Error", message)
            print(message)
        
        # Start the loading thread
        threading.Thread(target=load_data, daemon=True).start()
            
    except Exception as e:
        print(f"Error setting up table: {e}")
    
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
            selected_items = [table.item(item, "values") for item in selected]
            print(f"Selected {len(selected_items)} items:")
            for item in selected_items:
                print(item)
        else:
            print("No item selected")
    
    # Add a quit button
    quit_button = tk.Button(button_frame, text="Quit", command=root.quit)
    quit_button.pack(side=tk.RIGHT)

    # Add a process selection button
    button = tk.Button(button_frame, text="Process Selection", command=button_action)
    button.pack(side=tk.RIGHT)

    # Add filters section
    # Global variables to track filter selections
    filters = {
        'enemy_type': 'All',
        'Enemy Subfaction': 'All',
        'sector': 'All',
        'planet': 'All'
    }
    
    # Function to filter data based on all selected filters
    def filter_data():
        # Clear the table
        for item in table.get_children():
            table.delete(item)
        
        # Reload data based on all selected filters
        try:
            df = pd.read_excel(excel_file)
            
            # Apply enemy type filter
            if filters['enemy_type'] != 'All':
                df = df[df['Enemy Type'] == filters['enemy_type']]
            
            # Apply subfaction filter
            if filters['Enemy Subfaction'] != 'All':
                df = df[df['Enemy Subfaction'] == filters['Enemy Subfaction']]
            # Apply sector filter
            if filters['sector'] != 'All':
                df = df[df['Sector'] == filters['sector']]
                
            # Apply planet filter
            if filters['planet'] != 'All':
                df = df[df['Planet'] == filters['planet']]
            
            # Insert filtered data into table
            # Insert filtered data into table
            for _, row in df.iterrows():
                values = [str(val) if pd.notna(val) else "" for val in row]
                table.insert("", tk.END, values=values)
                
        except Exception as e:
            show_error(f"Error filtering data: {e}")

    # Enemy Type filter
    enemy_label = tk.Label(button_frame, text="Select enemy Type:")
    enemy_label.pack(side=tk.LEFT)
    enemy_var = tk.StringVar()
    enemy_dropdown = ttk.Combobox(button_frame, textvariable=enemy_var)
    enemy_dropdown['values'] = ['All', 'Automatons', 'Illuminate', 'Terminids']
    enemy_dropdown.current(0)
    enemy_dropdown.pack(side=tk.LEFT)
    
    def on_enemy_select(event):
        filters['enemy_type'] = enemy_var.get()
        filter_data()
    
    enemy_dropdown.bind("<<ComboboxSelected>>", on_enemy_select)
    
    # Subfaction filter
    subfaction_label = tk.Label(button_frame, text="Select subfaction:")
    subfaction_label.pack(side=tk.LEFT, padx=(10, 0))
    subfaction_var = tk.StringVar()
    subfaction_dropdown = ttk.Combobox(button_frame, textvariable=subfaction_var)
    subfaction_dropdown['values'] = ['All', 'Terminid Horde', 'Predator Strain', 'Spore Burst Strain', 
                                    'Automaton Legion', 'Jet Brigade', 'Incineration Corps', 
                                    'Jet Brigade & Incineration Corps', 'Illuminate Cult']
    subfaction_dropdown.current(0)
    subfaction_dropdown.pack(side=tk.LEFT)
    
    def on_subfaction_select(event):
        filters['Enemy Subfaction'] = subfaction_var.get()
        filter_data()
    
    subfaction_dropdown.bind("<<ComboboxSelected>>", on_subfaction_select)
    
    # Sector filter
    sector_label = tk.Label(button_frame, text="Select sector:")
    sector_label.pack(side=tk.LEFT, padx=(10, 0))
    sector_var = tk.StringVar()
    sector_dropdown = ttk.Combobox(button_frame, textvariable=sector_var)
    sector_dropdown['values'] = ['All', 'Akira Sector', 'Alstrad Sector', 'Altus Sector', 'Andromeda Sector', 'Arturion Sector', 'Barnard Sector', 'Borgus Sector', 'Cancri Sector', 'Cantolus Sector', 'Celeste Sector', 'Draco Sector', 'Falstaff Sector', 'Farsight Sector', 'Ferris Sector', 'Gallux Sector',
                                'Gellert Sector', 'Gothmar Sector', 'Guang Sector', 'Hanzo Sector', 'Hawking Sector', 'Hydra Sector', 'Idun Sector', 'Iptus Sector', 'Jin Xi Sector', 'Kelvin Sector', 'Korpus Sector', 'L\'estrade Sector', 'Lacaille Sector', 'Leo Sector', 'Marspira Sector', 'Meridian Sector',
                                'Mirin Sector', 'Morgon Sector', 'Nanos Sector', 'Omega Sector', 'Orion Sector', 'Quintus Sector', 'Rictus Sector', 'Rigel Sector', 'Sagan Sector', 'Saleria Sector', 'Severin Sector', 'Sol System', 'Sten Sector', 'Talus Sector', 'Tanis Sector', 'Tarragon Sector', 'Theseus Sector',
                                'Trigon Sector', 'Umlaut Sector', 'Ursa Sector', 'Valdis Sector', 'Xi Tauri Sector', 'Xzar Sector', 'Ymir Sector']
    sector_dropdown.current(0)
    sector_dropdown.pack(side=tk.LEFT)
    
    def on_sector_select(event):
        filters['sector'] = sector_var.get()
        filter_data()
    
    sector_dropdown.bind("<<ComboboxSelected>>", on_sector_select)

    #planet filter
    planet_label = tk.Label(button_frame, text="Select planet:")
    planet_label.pack(side=tk.LEFT, padx=(10, 0))
    planet_var = tk.StringVar()
    planet_dropdown = ttk.Combobox(button_frame, textvariable=planet_var)
    planet_dropdown['values'] = ['All','Alaraph', 'Alathfar XI', 'Andar', 'Asperoth Prime', 'Keid', 'Kneth Port', 'Klaka 5', 'Kraz', 'Pathfinder V', 'Klen Dahth II', 'Widow\'s Harbor', 'New Haven', 'Pilen V', 'Charbal-VII', 'Charon Prime', 'Martale', 'Marfark', 'Matar Bay', 'Mortax Prime', 'Kirrik', 'Wilford Station', 'Arkturus',
                                'Pioneer II', 'Electra Bay', 'Deneb Secundus', 'Fornskogur II', 'Veil', 'Marre IV', 'Midasburg', 'Darrowsport', 'Hydrofall Prime', 'Ursica XI', 'Achird III', 'Achernar Secundus', 'Darius II', 'Prosperity Falls', 'Cerberus IIIc', 'Effluvia', 'Seyshel Beach', 'Fort Sanctuary', 'Kelvinor', 'Martyr\'s Bay',
                                'Freedom Peak', 'Viridia Prime', 'Obari', 'Sulfura', 'Nublaria I', 'Krakatwo', 'Ivis', 'Slif', 'Moradesh', 'Meridia', 'Crimsica', 'Estanu', 'Fori Prime', 'Bore Rock', 'Esker', 'Socorro III', 'Erson Sands', 'Prasa', 'Pollux 31', 'Polaris Prime', 'Pherkad Secundus', 'Grand Errant', 'Hadar', 'Haldus', 'Zea Rugosia',
                                'Herthon Secundus', 'Kharst', 'Bashyr', 'Rasp', 'Acubens Prime', 'Adhara', 'Afoyay Bay', 'Minchir', 'Mintoria', 'Blistica', 'Zzaniah Prime', 'Zosma', 'Okul VI', 'Solghast', 'Diluvia', 'Elysian Meadows', 'Alderidge Cove', 'Bellatrix', 'Botein', 'Khandark', 'Heze Bay', 'Alairt III', 'Alamak VII', 'New Stockholm', 'Ain-5',
                                'Mordia 9', 'Euphoria III', 'Skitter', 'Kuma', 'Aesir Pass', 'Vernen Wells', 'Menkent', 'Wraith', 'Atrama', 'Myradesh', 'Maw', 'Providence', 'Primordia', 'Krakabos', 'Iridica', 'Valgaard', 'Ratch', 'Acamar IV', 'Pandion-XXIV', 'Gacrux', 'Phact Bay', 'Gar Haren', 'Gatria', 'Zegema Paradise', 'Fort Justice', 'New Kiruna',
                                'Igla', 'Emeria', 'Crucible', 'Volterra', 'Caramoor', 'Alta V', 'Inari', 'Navi VII', 'Omicron', 'Nabatea Secundus', 'Gemstone Bluffs', 'Epsilon Phoencis VI', 'Enuliale', 'Disapora X', 'Lesath', 'Penta', 'Chort Bay', 'Choohe', 'Ras Algethi', 'Propus', 'Halies Port', 'Haka', 'Curia', 'Barabos', 'Fenmire', 'Tarsh', 'Mastia',
                                'Emorath', 'Ilduna Prime', 'Baldrick Prime', 'Liberty Ridge', 'Hellmire', 'Nivel 43', 'Zagon Prime', 'Oshaune', 'Myrium', 'Eukoria', 'Regnus', 'Mog', 'Dolph', 'Julheim', 'Bekvam III', 'Duma Tyr', 'Setia', 'Senge 23', 'Seasse', 'Hydrobius', 'Karlia', 'Terrek', 'Azterra', 'Fort Union', 'Cirrus', 'Heeth', 'Angel\'s Venture',
                                'Veld', 'Termadon', 'Stor Tha Prime', 'Spherion', 'Stout', 'Leng Secundus', 'Valmox', 'Iro', 'Grafmere', 'Kerth Secundus', 'Parsh', 'Oasis', 'Genesis Prime', 'Rogue 5', 'RD-4', 'Hesoe Prime', 'Hort', 'Rirga Bay', 'Oslo Station', 'Gunvald', 'Borea', 'Calypso', 'Outpost 32', 'Reaf', 'Irulta', 'Maia', 'Malevelon Creek', 'Durgen',
                                'Ubanea', 'Tibit', 'Super Earth', 'Mars', 'Trandor', 'Peacock', 'Partion', 'Overgoe Prime', 'Azur Secundus', 'Shallus', 'Shelt', 'Gaellivare', 'Imber', 'Claorell', 'Vog-Sojoth', 'Clasa', 'Yed Prior', 'Zefia', 'Demiurg', 'East Iridium Trading Bay', 'Brink-2', 'Osupsam', 'Canopus', 'Bunda Secundus', 'The Weir', 'Kuper', 'Caph', 'Castor',
                                'Tien Kwan', 'Lastofe', 'Varylia 5', 'Choepessa IV', 'Ustotu', 'Troost', 'Vandalon IV', 'Erata Prime', 'Fenrir III', 'Turing', 'Skaash', 'Acrab XI', 'Acrux IX', 'Gemma', 'Merga IV', 'Merak', 'Cyberstan', 'Aurora Bay', 'Mekbuda', 'Videmitarix Prime', 'Skat Bay', 'Sirius', 'Siemnot', 'Shete', 'Mort', 'P\u00F6pli IX', 'Ingmar', 'Mantes',
                                'Draupnir', 'Meissa', 'Wasat', 'X-45', 'Vega Bay', 'Wezen']
    def on_planet_select(event):
        filters['planet'] = planet_var.get()
        filter_data()
    planet_dropdown.current(0)
    planet_dropdown.pack(side=tk.LEFT)

    planet_dropdown.bind("<<ComboboxSelected>>", on_planet_select)

    # Add a button to clear filters
    def clear_filters():
        filters['enemy_type'] = 'All'
        filters['Enemy Subfaction'] = 'All'
        filters['sector'] = 'All'
        filters['planet'] = 'All'
        
        enemy_dropdown.current(0)
        subfaction_dropdown.current(0)
        sector_dropdown.current(0)
        planet_dropdown.current(0)
        
        filter_data()
    clear_button = tk.Button(button_frame, text="Clear Filters", command=clear_filters)
    clear_button.pack(side=tk.LEFT, padx=(10, 0))
    
    # Run the main event loop
    root.mainloop()

if __name__ == "__main__":
    main()