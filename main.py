"""
CREDITS TO DEAN FOR THE STUPID AMOUNT OF DATA HE PROVIDED FOR THE JSON FILES
CREDITS TO ADAM FOR THE SCRIPT AND THE GUI
"""

import tkinter as tk
from tkinter import ttk, messagebox
import requests
import os
from datetime import datetime, timezone, timedelta
import json
import pandas as pd
import logging
from typing import Dict, List, Optional
from pypresence import Presence
import time
import configparser

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Read configuration from config.config
config = configparser.ConfigParser()
config.read('config.config')

# Constants
DEBUG = True
DATE_FORMAT = "%d-%m-%Y %H:%M:%S"
SETTINGS_FILE = 'user_settings.json'
EXCEL_FILE_PROD = config['Excel Location']['PROD']
EXCEL_FILE_TEST = config['Excel Location']['TEST']


DISCORD_CLIENT_ID = config['Discord']['DISCORD_CLIENT_ID']
RPC_UPDATE_INTERVAL = 15  # seconds

# Discord webhook configuration
WEBHOOK_URLS = {
    'PROD': config['Webhooks']['PROD'].split(','),  # Split comma-separated URLs into list
    'TEST': config['Webhooks']['TEST'].split(',')
}
ACTIVE_WEBHOOK = WEBHOOK_URLS['TEST'] if DEBUG else WEBHOOK_URLS['PROD']

# Enemy icons and colors from config
ENEMY_ICONS = {
    "Automatons": config['EnemyIcons']['Automatons'],
    "Terminids": config['EnemyIcons']['Terminids'],
    "Illuminate": config['EnemyIcons']['Illuminate']
}

DIFFICULTY_ICONS = {
    "1 - TRIVIAL": config['DifficultyIcons']['1 - TRIVIAL'],
    "2 - EASY": config['DifficultyIcons']['2 - EASY'],
    "3 - MEDIUM": config['DifficultyIcons']['3 - MEDIUM'],
    "4 - CHALLENGING": config['DifficultyIcons']['4 - CHALLENGING'],
    "5 - HARD": config['DifficultyIcons']['5 - HARD'],
    "6 - EXTREME": config['DifficultyIcons']['6 - EXTREME'],
    "7 - SUICIDE MISSION": config['DifficultyIcons']['7 - SUICIDE MISSION'],
    "8 - IMPOSSIBLE": config['DifficultyIcons']['8 - IMPOSSIBLE'],
    "9 - HELLDIVE": config['DifficultyIcons']['9 - HELLDIVE'],
    "10 - SUPER HELLDIVE": config['DifficultyIcons']['10 - SUPER HELLDIVE']
}
SYSTEM_COLORS = {
    "Automatons": config['SystemColors']['Automatons'],
    "Terminids": config['SystemColors']['Terminids'],
    "Illuminate": config['SystemColors']['Illuminate']
}

def get_enemy_icon(enemy_type: str) -> str:
    """Get the Discord emoji icon for an enemy type."""
    return ENEMY_ICONS.get(enemy_type, "NaN")

def get_difficulty_icon(difficulty: str) -> str:
    """Get Difficulty Icons"""
    return DIFFICULTY_ICONS.get(difficulty, "NaN")

def get_system_color(enemy_type: str) -> int:
    """Get the Discord color code for an enemy type."""
    return int(SYSTEM_COLORS.get(enemy_type, "0"))

class MissionLogGUI:
    """GUI application for logging Helldiver 2 mission data."""
    
    def __init__(self, root: tk.Tk) -> None:
        """Initialize the GUI application."""
        self.root = root
        self.root.title("Helldiver Mission Log Manager")
        self.root.resizable(False, False)
        self.root.iconphoto(False, tk.PhotoImage(file='SuperEarth.png'))
        self.settings_file = SETTINGS_FILE
        self._setup_variables()
        self._setup_discord_rpc()  # Add this line
        self._create_main_frame()
        self._setup_ui()
        self.root.after(100, self.load_settings)

        # Add periodic RPC updates
        self.root.after(1000, self._periodic_rpc_update)

    def _periodic_rpc_update(self) -> None:
        """Periodically update Discord Rich Presence."""
        self._update_discord_presence()
        self.root.after(1000, self._periodic_rpc_update)

    def _setup_variables(self) -> None:
        """Initialize tkinter variables with validation."""
        self.sector = tk.StringVar()
        self.planet = tk.StringVar()
        self.mission_type = tk.StringVar()
        self.kills = tk.StringVar()
        self.deaths = tk.StringVar()
        self.enemy_type = tk.StringVar()
        self.Helldivers = tk.StringVar()
        self.mission_category = tk.StringVar()
        self.rating = tk.StringVar(value="Outstanding Patriotism")
        self.level = tk.IntVar()
        self.title = tk.StringVar()
        self.difficulty = tk.StringVar()
        self.MO = tk.BooleanVar()
        self.DSS = tk.BooleanVar()
        self.DSSMod = tk.StringVar()
        self.report_style = tk.StringVar(value='Modern')

        # Add validation for numeric fields
        validate_cmd = self.root.register(self._validate_numeric_input)
        self.kills.trace_add("write", lambda *args: self._validate_field(self.kills))
        self.deaths.trace_add("write", lambda *args: self._validate_field(self.deaths))

    def _validate_numeric_input(self, value: str) -> bool:
        """Validate that input is numeric and within acceptable range."""
        if not value:
            return True
        try:
            return 0 <= int(value) <= 999999
        except ValueError:
            return False

    def _validate_field(self, var: tk.StringVar) -> None:
        """Clear invalid numeric fields."""
        if not self._validate_numeric_input(var.get()):
            var.set("")

    def _create_main_frame(self) -> None:
        """Create the main application frame."""
        self.frame = ttk.Frame(self.root, padding="10")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        style = ttk.Style()
        style.configure('TLabel', font=('Arial', 10))
        style.configure('TButton', font=('Arial', 10, 'bold'))


    def _setup_discord_rpc(self) -> None:
        """Initialize Discord Rich Presence."""
        try:
            self.RPC = Presence(DISCORD_CLIENT_ID)
            self.RPC.connect()
            self.last_rpc_update = 0
            logging.info("Discord Rich Presence initialized successfully")
        except Exception as e:
            logging.error(f"Failed to initialize Discord Rich Presence: {e}")
            self.RPC = None


    def _setup_ui(self) -> None:
        """Set up the complete UI layout."""
        # Create main content frame with padding
        content = ttk.Frame(self.frame, padding=(20, 10))
        content.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        SETime = (datetime.now(timezone.utc) + timedelta(hours=1)).strftime("%H:%M:%S")

        # Mission Information Section
        mission_frame = ttk.LabelFrame(content, text="Mission Information  SEST: {}".format(SETime), padding=10)

        def update_time():
            SETime = (datetime.now(timezone.utc) + timedelta(hours=1)).strftime("%H:%M:%S")
            mission_frame.config(text=f"Mission Information  SEST: {SETime}")

        self.update_time = update_time
        mission_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Load sectors from config
        with open('PlanetSectors.json', 'r') as f:
            sectors_data = json.load(f)
            sector_list = list(sectors_data.keys())

        # Mission Info Grid
        ttk.Label(mission_frame, text="Helldiver:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(mission_frame, textvariable=self.Helldivers, width=30).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(mission_frame, text="Level:").grid(row=0, column=2, sticky=tk.W, pady=5)
        ttk.Entry(mission_frame, textvariable=self.level, width=30).grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(mission_frame, text="Title:").grid(row=0, column=4, sticky=tk.W, pady=5)
        titles = ['CADET', 'SPACE CADET', 'SERGEANT', 'MASTER SERGEANT', 'CHIEF', 'SPACE CHIEF PRIME', 
             'DEATH CAPTAIN', 'MARSHALL', 'STAR MARSHALL', 'ADMIRAL', 'SKULL ADMIRAL', 'FLEET ADMIRAL',
             'ADMIRABLE ADMIRAL', 'COMMANDER', 'GALACTIC COMMANDER', 'HELL COMMANDER', 'GENERAL',
             '5-STAR GENERAL', '10-STAR GENERAL', 'PRIVATE', 'SUPER PRIVATE', 'SUPER CITIZEN',
             'VIPER COMMANDO', 'FIRE SAFETY OFFICER', 'EXPERT EXTERMINATOR', 'FREE OF THOUGHT',
             'ASSAULT INFANTRY', 'SUPER PEDESTRIAN', 'SERVANT OF FREEDOM']
        title_combo = ttk.Combobox(mission_frame, textvariable=self.title, values=titles, state='readonly', width=27)
        title_combo.grid(row=0, column=5, padx=5, pady=5)
        title_combo.set(titles[0])

        ttk.Label(mission_frame, text="Sector:").grid(row=1, column=0, sticky=tk.W, pady=5)
        sector_combo = ttk.Combobox(mission_frame, textvariable=self.sector, values=sector_list, state='readonly', width=27)
        sector_combo.grid(row=1, column=1, padx=5, pady=5)
        sector_combo.set(sector_list[0])

        ttk.Label(mission_frame, text="Planet:").grid(row=2, column=0, sticky=tk.W, pady=5)
        planet_combo = ttk.Combobox(mission_frame, textvariable=self.planet, state='readonly', width=27)
        planet_combo.grid(row=2, column=1, padx=5, pady=5)
        self.sector_combo = sector_combo
        self.planet_combo = planet_combo

        def update_planets(*args):
            selected_sector = self.sector.get()
            planet_list = sectors_data[selected_sector]["planets"]
            planet_combo['values'] = planet_list
            planet_combo.set(planet_list[0])

        sector_combo.bind('<<ComboboxSelected>>', update_planets)
        update_planets()

        # Mission Details Section
        details_frame = ttk.LabelFrame(content, text="Mission Details", padding=10)
        details_frame.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Enemy Type Selection
        ttk.Label(details_frame, text="Enemy Type:").grid(row=0, column=0, sticky=tk.W, pady=5)
        with open('Missions.json', 'r') as f:
            missions_data = json.load(f)
            enemy_types = list(missions_data.keys())
        enemy_combo = ttk.Combobox(details_frame, textvariable=self.enemy_type, values=enemy_types, state='readonly', width=27)
        enemy_combo.grid(row=0, column=1, padx=5, pady=5)
        enemy_combo.set(enemy_types[0])

        ttk.Checkbutton(details_frame, text="Major Order", variable=self.MO).grid(row=0, column=2, padx=5, pady=5)

        ttk.Checkbutton(details_frame, text="DSS Active", variable=self.DSS).grid(row=0, column=3, padx=5, pady=5)

        # DSS Modifier dropdown
        self.dss_frame = ttk.Frame(details_frame)
        self.dss_frame.grid(row=0, column=4, columnspan=2, sticky=tk.W, pady=5)
        ttk.Label(self.dss_frame, text="DSS Modifier:").pack(side=tk.LEFT)
        dss_mods = ["Inactive", "Orbital Blockade", "Heavy Ordinance Distribution", "Eagle Storm"]
        self.DSSMod.set("Inactive")  # Set default value
        self.dss_combo = ttk.Combobox(self.dss_frame, textvariable=self.DSSMod, values=dss_mods, state='readonly', width=27)
        self.dss_combo.pack(side=tk.LEFT, padx=5)
        
        # Initially hide the dropdown
        self.dss_frame.grid_remove()
        
        # Function to toggle DSS modifier visibility
        def toggle_dss_mod(*args):
            if self.DSS.get():
                self.dss_frame.grid()
            else:
                self.DSSMod.set("Inactive")
                self.dss_frame.grid_remove()
            
        self.DSS.trace_add("write", toggle_dss_mod)

        # Mission Campaign Selection
        ttk.Label(details_frame, text="Mission Campaign:").grid(row=1, column=0, sticky=tk.W, pady=5)
        mission_cat_combo = ttk.Combobox(details_frame, textvariable=self.mission_category, state='readonly', width=27)
        mission_cat_combo.grid(row=1, column=1, padx=5, pady=5)

        # Difficulty Selection
        ttk.Label(details_frame, text="Mission Difficulty:").grid(row=2, column=0, sticky=tk.W, pady=5)
        difficulty_combo = ttk.Combobox(details_frame, textvariable=self.difficulty, state='readonly', width=27)
        difficulty_combo.grid(row=2, column=1, padx=5, pady=5)

        # Mission Type Selection
        ttk.Label(details_frame, text="Mission Type:").grid(row=3, column=0, sticky=tk.W, pady=5)
        mission_type_combo = ttk.Combobox(details_frame, textvariable=self.mission_type, state='readonly', width=27)
        mission_type_combo.grid(row=3, column=1, padx=5, pady=5)

        def update_mission_categories(*args):
            enemy = self.enemy_type.get()
            if enemy in missions_data:
                categories = list(missions_data[enemy].keys())
                mission_cat_combo['values'] = categories
                if categories:
                    mission_cat_combo.set(categories[0])
                    update_mission_types()

        def update_mission_types(*args):
            enemy = self.enemy_type.get()
            category = self.mission_category.get()
            
            if enemy in missions_data and category in missions_data[enemy]:
                if missions_data[enemy][category] != "Unknown":
                    difficulties = list(missions_data[enemy][category].keys())
                    difficulty_combo['values'] = difficulties
                    difficulty_combo.set(difficulties[0])
                    
                    # Set initial mission types from first difficulty
                    first_difficulty = difficulties[0]
                    available_missions = missions_data[enemy][category][first_difficulty]
                    mission_type_combo['values'] = available_missions
                    if available_missions:
                        mission_type_combo.set(available_missions[0])
                else:
                    mission_type_combo['values'] = ["No missions available"]
                    mission_type_combo.set("No missions available")
                    difficulty_combo['values'] = ["No difficulties available"]
                    difficulty_combo.set("No difficulties available")

        def update_available_missions(*args):
            enemy = self.enemy_type.get()
            category = self.mission_category.get()
            difficulty = self.difficulty.get()
            
            if (enemy in missions_data and 
                category in missions_data[enemy] and 
                difficulty in missions_data[enemy][category]):
                
                available_missions = missions_data[enemy][category][difficulty]
                mission_type_combo['values'] = available_missions
                if available_missions:
                    mission_type_combo.set(available_missions[0])

        enemy_combo.bind('<<ComboboxSelected>>', update_mission_categories)
        mission_cat_combo.bind('<<ComboboxSelected>>', update_mission_types)
        difficulty_combo.bind('<<ComboboxSelected>>', update_available_missions)
        
        # Initial setup
        update_mission_categories()

        # Statistics Section
        stats_frame = ttk.LabelFrame(content, text="Mission Statistics", padding=10)
        stats_frame.grid(row=2, column=0, padx=5, pady=5, sticky=(tk.W, tk.N, tk.S))

        

        ttk.Label(stats_frame, text="Kills:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(stats_frame, textvariable=self.kills, width=30).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(stats_frame, text="Deaths:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(stats_frame, textvariable=self.deaths, width=30).grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(stats_frame, text="Performance:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ratings = ["Outstanding Patriotism", "Superior Valour", "Honourable Duty", "Unremarkable Performance", "Disappointing Service", "Disgraceful Conduct"]
        self.rating.set(ratings[0])  # Set default value before creating Combobox
        rating_combo = ttk.Combobox(stats_frame, textvariable=self.rating, values=ratings, state='readonly', width=27)
        rating_combo.grid(row=3, column=1, padx=5, pady=5)

        # Submit Button
        submit_button = ttk.Button(content, text="Submit Mission Report", command=self.submit_data)
        submit_button.grid(row=3, column=0, pady=15)

        # Create a frame for the report style and export section
        bottom_frame = ttk.LabelFrame(content, text="Report Style and Export", padding=10)
        bottom_frame.grid(row=4, column=0, pady=5, sticky=(tk.W, tk.E))

        # Report Style section (left side)
        style_frame = ttk.Frame(bottom_frame)
        style_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(style_frame, text="Report Style:").pack(side=tk.LEFT)

        # Export section (right side)
        export_frame = ttk.LabelFrame(bottom_frame, text="Exporting", padding=10)
        export_frame.pack(side=tk.RIGHT, padx=5)

        self.report_style = tk.StringVar()
        report_styles = ['Modern', 'Fax']
        style_combo = ttk.Combobox(style_frame, textvariable=self.report_style, values=report_styles, state='readonly', width=27)
        style_combo.pack(side=tk.LEFT, padx=5)
        style_combo.set(report_styles[0])

        #export button
        export_button = ttk.Button(export_frame, text="Export Excel Data to Webhook", command=lambda: os.system('python sub.py'))
        export_button.grid(row=6, column=0, pady=15)


    def _update_discord_presence(self) -> None:
        """Update Discord Rich Presence with current mission information."""
        if not hasattr(self, 'RPC') or self.RPC is None:
            return

        current_time = time.time()
        if current_time - self.last_rpc_update < RPC_UPDATE_INTERVAL:
            return

        try:
            helldiver = self.Helldivers.get() or "Unknown Helldiver"
            sector = self.sector.get() or "No Sector"
            planet = self.planet.get() or "No Planet"
            enemytype = self.enemy_type.get() or "Unknown Enemy"
            level = self.level.get() or 0
            title = self.title.get() or "No Title"

            # Map enemy types to Discord asset names
            enemy_assets = {
                "Automatons": "bots",
                "Terminids": "bugs",
                "Illuminate": "squids"
            }
            
            small_image = enemy_assets.get(enemytype, "unknown")

            self.RPC.update(
                state=f"Sector: {sector}\nPlanet: {planet}",
                details=f"Helldiver: {helldiver} Level: {level} | {title}",
                large_image="superearth",
                large_text="Helldivers 2",
                small_image=small_image,
                small_text=f"Fighting: {enemytype}",
            )
            self.last_rpc_update = current_time
        except Exception as e:
            logging.error(f"Failed to update Discord Rich Presence: {e}")

    def load_settings(self) -> None:
        """Load user settings from file."""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                    self._apply_settings(settings)
        except Exception as e:
            self._show_error(f"Error loading settings: {e}")

    def _apply_settings(self, settings: Dict) -> None:
        """Apply loaded settings to UI elements."""
        self.Helldivers.set(settings.get('helldiver', ''))
        self.level.set(settings.get('level', 0))
        self.difficulty.set(settings.get('difficulty', '1 - TRIVIAL'))
        self.mission_category.set(settings.get('campaign', 'Ivasion'))
        self.DSS.set(settings.get('DSS', False))


        if settings.get('DSS'):
            self.DSSMod.set(settings.get('DSSMod', 'None'))
        
        # For mission type
        if settings.get('mission'):
            self.mission_type.set(settings.get('mission'))

        title_combo = self.frame.winfo_children()[0].winfo_children()[0].winfo_children()[5]
        if settings.get('title') in title_combo['values']:
            self.title.set(settings.get('title'))
        if settings.get('sector') in self.sector_combo['values']:
            self.sector.set(settings.get('sector'))
            self.root.update()
            self.sector_combo.event_generate('<<ComboboxSelected>>')
            
            if settings.get('planet') in self.planet_combo['values']:
                self.planet.set(settings.get('planet'))

    def save_settings(self) -> None:
        """Save current settings to file."""
        settings = {
            'helldiver': self.Helldivers.get(),
            'level': self.level.get(),
            'title': self.title.get(),
            'sector': self.sector.get(),
            'planet': self.planet.get(),
            'difficulty': self.difficulty.get(),
            'mission': self.mission_type.get(),
            'DSS': self.DSS.get(),
            'DSSMod': self.DSSMod.get(),
            'campaign': self.mission_category.get()
        }
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            self._show_error(f"Error saving settings: {e}")

    def submit_data(self) -> None:
        """Handle mission report submission."""
        if not self._validate_submission():
            return
            
        self.save_settings()
        self.update_time()
        data = self._collect_mission_data()

        if self._save_to_excel(data):
            if self._send_to_discord(data):
                self._show_success("Mission report submitted successfully!")
            
    def _validate_submission(self) -> bool:
        """Validate all required fields before submission."""
        try:
            # Validate numeric fields
            level = int(self.level.get())
            kills = int(self.kills.get())
            deaths = int(self.deaths.get())

            if level < 1 or level > 150:  # Add reasonable level range
                self._show_error("Level must be between 1 and 150")
                return False

            if kills < 0 or kills > 10000:
                self._show_error("Invalid number of kills")
                return False

            if deaths < 0 or deaths > 1000:
                self._show_error("Invalid number of deaths")
                return False

            # Validate required text fields
            if not self.Helldivers.get().strip():
                self._show_error("Helldiver name is required")
                return False

            if not self.mission_type.get().strip():
                self._show_error("Mission type is required")
                return False

            return True
        except ValueError:
            self._show_error("Invalid numeric input")
            return False

    def _show_error(self, message: str) -> None:
        """Display error message to user."""
        messagebox.showerror("Error", message)

    def _show_success(self, message: str) -> None:
        """This is a useless feature..."""

    def _collect_mission_data(self) -> Dict:
        """Collect all mission data into a dictionary."""

        print(self.DSSMod.get())
        return {
            'Helldivers': self.Helldivers.get(),
            'Level': self.level.get(),
            'Title': self.title.get(),
            'Sector': self.sector.get(),
            'Planet': self.planet.get(),
            'Enemy Type': self.enemy_type.get(),
            'Major Order': self.MO.get(),
            'DSS Active': self.DSS.get(),
            'DSS Modifier': self.DSSMod.get() or "Inactive",
            'Mission Category': self.mission_category.get(),
            'Mission Type': self.mission_type.get(),
            'Difficulty': self.difficulty.get(),
            'Kills': self.kills.get(),
            'Deaths': self.deaths.get(),
            'Rating': self.rating.get(),
            'Time': datetime.now().strftime(DATE_FORMAT)
        }


        # Save to Excel
        self._save_to_excel(data)

        # Send to Discord
        self._send_to_discord(data)

    def _save_to_excel(self, data: Dict) -> bool:
        """Save mission data to Excel file with backup."""
        excel_file = EXCEL_FILE_TEST if DEBUG else EXCEL_FILE_PROD

        try:
            # Save new data
            df = pd.DataFrame([data])
            if os.path.exists(excel_file):
                existing_df = pd.read_excel(excel_file)
                df = pd.concat([existing_df, df], ignore_index=True)
            df.to_excel(excel_file, index=False)
            logging.info(f"Successfully saved data to {excel_file}")
            return True

        except Exception as e:
            logging.error(f"Error saving to Excel: {e}")
            self._show_error(f"Error saving to Excel: {e}")
            return False

    def _send_to_discord(self, data: Dict) -> bool:
        """Send mission report to Discord with improved error handling."""
        try:
            Stars = ""
            GoldStar = config['Stars']['GoldStar']
            GreyStar = config['Stars']['GreyStar']
            if self.rating.get() == "Outstanding Patriotism":
                Stars = "{}{}{}{}{}".format(GoldStar, GoldStar, GoldStar, GoldStar, GoldStar)
            elif self.rating.get() == "Superior Valour":
                Stars = "{}{}{}{}{}".format(GoldStar, GoldStar, GoldStar, GoldStar, GreyStar)
            elif self.rating.get() == "Honourable Duty":
                Stars = "{}{}{}{}{}".format(GoldStar, GoldStar, GoldStar, GreyStar, GreyStar)
            elif self.rating.get() == "Unremarkable Performance":
                Stars = "{}{}{}{}{}".format(GoldStar, GoldStar, GreyStar, GreyStar, GreyStar)
            elif self.rating.get() == "Disappointing Service":
                Stars = "{}{}{}{}{}".format(GoldStar, GreyStar, GreyStar, GreyStar, GreyStar)
            else:
                Stars = "{}{}{}{}{}".format(GreyStar, GreyStar, GreyStar, GreyStar, GreyStar)
            date = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            enemy_icon = get_enemy_icon(data['Enemy Type'])
            system_color = get_system_color(data['Enemy Type'])
            diff_icon = get_difficulty_icon(data['Difficulty'])

            # Format the message for Discord
            message_content1 = (
                f"# ================================\n"
                f"> # | **Date:** {date}\n"
                f"> # | **Mission Report for {data['Helldivers']}**\n"
                f"> # --------------------------------\n"
                f"> # | **Sector:** {data['Sector']}\n"
                f"> ## | **Planet:** {data['Planet']}\n"
                f"> ## | **Enemy Type:** {data['Enemy Type']} {enemy_icon}\n"
                f"> ## | **Mission Category:** {data['Mission Category']}\n"
                f"> ## | **Mission Type:** {data['Mission Type']}\n"
                f"> ## | **Mission Difficulty:** {data['Difficulty']}\n"
                f"> # --------------------------------\n"
                f"> ### | **Kills:** {data['Kills']}\n"
                f"> ### | **Deaths:** {data['Deaths']}\n"
                f"# ================================\n"
            )

            message_content2 = {
                "content": None,
                "embeds": [{
                    "title": f"Date: {date}\n> Mission Report for {data['Helldivers']}\n> Level {data['Level']} | {data['Title']}",
                    "description": f"=============================\nSector: {data['Sector']}\n\nPlanet: {data['Planet']}\n\nEnemy Faction: {data['Enemy Type']} {enemy_icon}\n\n Major Order: {data['Major Order']}\n\n DSS Active: {data['DSS Active']}\n\n DSS Modifier: {data['DSS Modifier']}\n\nCampaign: {data['Mission Category']}\n=============================",
                    "color": system_color,
                    "fields": [{
                        "name": "> Mission Statistics",
                        "value": f"=============================\nMission: {data['Mission Type']}\n\n Difficulty: {data['Difficulty']} {diff_icon}\n\nKills: {data['Kills']}\n\nDeaths: {data['Deaths']}\n\n Performance: {data['Rating']}\n\n {Stars}\n============================="
                    }],
                    "author": {
                        "name": "Super Earth Mission Control"
                    },
                    "image": {"url": "https://images-ext-1.discordapp.net/external/9jCMPgdYyRaWcSNmU0JZjnDQD9Lt2awiLxegodvltpc/https/i.ibb.co/qY68vxkS/1f75a494d68eae549179996c4610bda0c22.png"},
                    "thumbnail": {"url": "https://cdn.discordapp.com/attachments/1337173158377033779/1337468193777782845/super-earth-helldivers-svg-logo-v0-0cvbn5nesrvc1.png?ex=67a78dd2&is=67a63c52&hm=3a9d304d5aafbf928ed549190ed427a7c510afc594f54d14ba582ea3e72445e6&"}
                }],
                "attachments": []
            }

            webhook_data = {"content": message_content1} if self.report_style.get() == "Fax" else message_content2
            
            successes = []
            for url in ACTIVE_WEBHOOK:
                try:
                    response = requests.post(url, json=webhook_data)
                    if response.status_code == 204:
                        logging.info(f"Successfully sent to Discord webhook: {url}")
                        successes.append(True)
                    else:
                        logging.error(f"Failed to send to Discord webhook {url}. Status code: {response.status_code}")
                        self._show_error(f"Failed to send to Discord (Status: {response.status_code})")
                        successes.append(False)
                except requests.RequestException as e:
                    logging.error(f"Network error sending to Discord webhook {url}: {e}")
                    self._show_error(f"Failed to connect to Discord webhook")
                    successes.append(False)
                except Exception as e:
                    logging.error(f"Unexpected error sending to Discord webhook {url}: {e}")
                    self._show_error("An unexpected error occurred while sending to Discord")
                    successes.append(False)

            # Return True only if all webhooks succeeded
            return any(successes) if successes else False
        except Exception as e:
            logging.error(f"Error preparing Discord message: {e}")
            self._show_error("Error preparing Discord message")
            return False

    def export_data(self) -> None:
        """Export Excel data to webhook."""
        excel_file = "mission_log_test.xlsx" if DEBUG else "mission_log.xlsx"
        try:
            if not os.path.exists(excel_file):
                self._show_error("No Excel file found to export")
                return

            df = pd.read_excel(excel_file)
            for _, row in df.iterrows():
                data = row.to_dict()
                self._send_to_discord(data)

            self._show_success("Excel data exported successfully!")
        except Exception as e:
            self._show_error(f"Error exporting data: {e}")
            logging.error(f"Error during Excel export: {e}")

    def __del__(self) -> None:
        """Clean up resources on deletion."""
        if hasattr(self, 'RPC') and self.RPC is not None:
            try:
                self.RPC.close()
            except:
                pass

if __name__ == "__main__":
    root = tk.Tk()
    app = MissionLogGUI(root)
    root.mainloop()