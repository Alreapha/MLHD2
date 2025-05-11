"""
CREDITS TO DEAN FOR THE STUPID AMOUNT OF DATA HE PROVIDED FOR THE JSON FILES
CREDITS TO ADAM FOR THE SCRIPT AND THE GUI
"""

import tkinter as tk
from tkinter import ttk, messagebox
import requests
from datetime import datetime, timezone, timedelta
import json
import pandas as pd
import logging
from typing import Dict, List, Optional
from pypresence import Presence
import time
import configparser
import threading
import os
import subprocess
import random
import re

# Read configuration from config.config
config = configparser.ConfigParser()
config.read('config.config')

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Constants
DEBUG = config.getboolean('DEBUGGING', 'DEBUG', fallback=False)
VERSION = "1.3.125"
DATE_FORMAT = "%d-%m-%Y %H:%M:%S"
SETTINGS_FILE = 'user_settings.json'
EXCEL_FILE_TEST = 'mission_log_test.xlsx'
EXCEL_FILE_PROD = 'mission_log.xlsx'

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
    "Illuminate": config['EnemyIcons']['Illuminate'],
    "Observing": config['EnemyIcons']['Observation'],
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

# Planet Icons
PLANET_ICONS = {
    "Super Earth": config['PlanetIcons']['Human Homeworld'],
    "Cyberstan": config['PlanetIcons']['Automaton Homeworld'],
    "Malevelon Creek": config['PlanetIcons']['Malevelon Creek'],
    "Calypso": config['PlanetIcons']['Calypso'],
    "Diaspora X": config['PlanetIcons']['Gloom'],
    "Enuliale": config['PlanetIcons']['Gloom'],
    "Epsilon Phoencis VI": config['PlanetIcons']['Gloom'],
    "Gemstone Bluffs": config['PlanetIcons']['Gloom'],
    "Nabatea Secundus": config['PlanetIcons']['Gloom'],
    "Navi VII": config['PlanetIcons']['Gloom'],
    "Azur Secundus": config['PlanetIcons']['Gloom'],
    "Erson Sands": config['PlanetIcons']['Gloom'],
    "Nivel 43": config['PlanetIcons']['Gloom'],
    "Zagon Prime": config['PlanetIcons']['Gloom'],
    "Hellmire": config['PlanetIcons']['Gloom'],
    "Omicron": config['PlanetIcons']['Gloom'],
    "Oshaune": config['PlanetIcons']['Gloom'],
    "Fori Prime": config['PlanetIcons']['Gloom']
}

# Campaign Icons
CAMPAIGN_ICONS = {
    "Defense": config['CampaignIcons']['Defense'],
    "Liberation": config['CampaignIcons']['Liberation'],
    "Invasion": config['CampaignIcons']['Invasion']
}

# Mission Icons
MISSION_ICONS = {
    "Terminate Illegal Broadcast": config['MissionIcons']['Terminate Illegal Broadcast'],
    "Pump Fuel To ICBM": config['MissionIcons']['Pump Fuel To ICBM'],
    "Upload Escape Pod Data": config['MissionIcons']['Upload Escape Pod Data'],
    "Spread Democracy": config['MissionIcons']['Spread Democracy'],
    "Conduct Geological Survey": config['MissionIcons']['Conduct Geological Survey'],
    "Launch ICBM": config['MissionIcons']['Launch ICBM'],
    "Retrieve Valuable Data": config['MissionIcons']['Retrieve Valuable Data'],
    "Blitz: Search and Destroy": config['MissionIcons']['Blitz Search and Destroy'],
    "PLACEHOLDER": config['MissionIcons']['PLACEHOLDER'],
    "Emergency Evacuation": config['MissionIcons']['Emergency Evacuation'],
    "Retrieve Essential Personnel": config['MissionIcons']['Retrieve Essential Personnel'],
    "Evacuate High-Value Assets": config['MissionIcons']['Evacuate High-Value Assets'],
    "Eliminate Brood Commanders": config['MissionIcons']['Eliminate Brood Commanders'],
    "Eliminate Chargers": config['MissionIcons']['Eliminate Chargers'],
    "Eliminate Impaler": config['MissionIcons']['Eliminate Impaler'],
    "Eliminate Bile Titans": config['MissionIcons']['Eliminate Bile Titans'],
    "Activate E-710 Pumps": config['MissionIcons']['Activate E-710 Pumps'],
    "Purge Hatcheries": config['MissionIcons']['Purge Hatcheries'],
    "Enable E-710 Extraction": config['MissionIcons']['Enable E-710 Extraction'],
    "Nuke Nursery": config['MissionIcons']['Nuke Nursery'],
    "Activate Terminid Control System": config['MissionIcons']['Activate Terminid Control System'],
    "Deactivate Terminid Control System": config['MissionIcons']['Deactivate Terminid Control System'],
    "Deploy Dark Fluid": config['MissionIcons']['Deploy Dark Fluid'],
    "Eradicate Terminid Swarm": config['MissionIcons']['Eradicate Terminid Swarm'],
    "Destroy Transmission Network": config['MissionIcons']['Destroy Transmission Network'],
    "Eliminate Devastators": config['MissionIcons']['Eliminate Devastators'],
    "Eliminate Automaton Hulks": config['MissionIcons']['Eliminate Automaton Hulks'],
    "Eliminate Automaton Factory Strider": config['MissionIcons']['Eliminate Automaton Factory Strider'],
    "Sabotage Supply Bases": config['MissionIcons']['Sabotage Supply Bases'],
    "Sabotage Air Base": config['MissionIcons']['Sabotage Air Base'],
    "Eradicate Automaton Forces": config['MissionIcons']['Eradicate Automaton Forces'],
    "Destroy Command Bunkers": config['MissionIcons']['Destroy Command Bunkers'],
    "Neutralize Orbital Defenses": config['MissionIcons']['Neutralize Orbital Defenses'],
    "Evacuate Colonists": config['MissionIcons']['Evacuate Colonists'],
    "Retrieve Recon Craft Intel": config['MissionIcons']['Retrieve Recon Craft Intel'],
    "Free Colony": config['MissionIcons']['Free Colony'],
    "Blitz: Destroy Illuminate Warp Ships": config['MissionIcons']['Blitz Destroy Illuminate Warp Ships'],
    "Destroy Harvesters": config['MissionIcons']['Destroy Harvesters'],
    "Extract Research Probe Data": config['MissionIcons']['Extract Research Probe Data'],
    "Collect Meteorological Data": config['MissionIcons']['Collect Meteorological Data'],
    "Collect Gloom-Infused Oil": config['MissionIcons']['Collect Gloom-Infused Oil'],
    "Blitz: Secure Research Site": config['MissionIcons']['Blitz Secure Research Site'],
    "Collect Gloom Spore Readings": config['MissionIcons']['Collect Gloom Spore Readings'],
    "Chart Terminid Tunnels": config['MissionIcons']['Chart Terminid Tunnels']
}

# Biome banners for Planets
BIOME_BANNERS = {
    "Propus": config['BiomeBanners']['Desert Dunes'],
    "Klen Dahth II": config['BiomeBanners']['Desert Dunes'],
    "Outpost 32": config['BiomeBanners']['Desert Dunes'],
    "Lastofe": config['BiomeBanners']['Desert Dunes'],
    "Diaspora X": config['BiomeBanners']['Desert Dunes'],
    "Zagon Prime": config['BiomeBanners']['Desert Dunes'],
    "Osupsam": config['BiomeBanners']['Desert Dunes'],
    "Mastia": config['BiomeBanners']['Desert Dunes'],
    "Caramoor": config['BiomeBanners']['Desert Dunes'],
    "Heze Bay": config['BiomeBanners']['Desert Dunes'],
    "Viridia Prime": config['BiomeBanners']['Desert Dunes'],
    "Durgen": config['BiomeBanners']['Desert Dunes'],
    "Phact Bay": config['BiomeBanners']['Desert Dunes'],
    "Keid": config['BiomeBanners']['Desert Dunes'],
    "Zzaniah Prime": config['BiomeBanners']['Desert Dunes'],
    "Choohe": config['BiomeBanners']['Desert Dunes'],
    "Pilen V": config['BiomeBanners']['Desert Cliffs'],
    "Zea Rugosia": config['BiomeBanners']['Desert Cliffs'],
    "Myradesh": config['BiomeBanners']['Desert Cliffs'],
    "Azur Secundus": config['BiomeBanners']['Desert Cliffs'],
    "Erata Prime": config['BiomeBanners']['Desert Cliffs'],
    "Mortax Prime": config['BiomeBanners']['Desert Cliffs'],
    "Cerberus IIIc": config['BiomeBanners']['Desert Cliffs'],
    "Ustotu": config['BiomeBanners']['Desert Cliffs'],
    "Erson Sands": config['BiomeBanners']['Desert Cliffs'],
    "Canopus": config['BiomeBanners']['Desert Cliffs'],
    "Hydrobius": config['BiomeBanners']['Desert Cliffs'],
    "Polaris Prime": config['BiomeBanners']['Desert Cliffs'],
    "Darrowsport": config['BiomeBanners']['Acidic Badlands'],
    "Darius II": config['BiomeBanners']['Acidic Badlands'],
    "Chort Bay": config['BiomeBanners']['Acidic Badlands'],
    "Leng Secundus": config['BiomeBanners']['Acidic Badlands'],
    "Rirga Bay": config['BiomeBanners']['Acidic Badlands'],
    "Shete": config['BiomeBanners']['Acidic Badlands'],
    "Skaash": config['BiomeBanners']['Acidic Badlands'],
    "Wraith": config['BiomeBanners']['Acidic Badlands'],
    "Slif": config['BiomeBanners']['Acidic Badlands'],
    "Wilford Station": config['BiomeBanners']['Acidic Badlands'],
    "Botein": config['BiomeBanners']['Acidic Badlands'],
    "Wasat": config['BiomeBanners']['Acidic Badlands'],
    "Esker": config['BiomeBanners']['Acidic Badlands'],
    "Charbal-VII": config['BiomeBanners']['Acidic Badlands'],
    "Kraz": config['BiomeBanners']['Rocky Canyons'],
    "Hydrofall Prime": config['BiomeBanners']['Rocky Canyons'],
    "Myrium": config['BiomeBanners']['Rocky Canyons'],
    "Vernen Wells": config['BiomeBanners']['Rocky Canyons'],
    "Calypso": config['BiomeBanners']['Rocky Canyons'],
    "Achird III": config['BiomeBanners']['Rocky Canyons'],
    "Azterra": config['BiomeBanners']['Rocky Canyons'],
    "Senge 23": config['BiomeBanners']['Rocky Canyons'],
    "Emeria": config['BiomeBanners']['Rocky Canyons'],
    "Fori Prime": config['BiomeBanners']['Rocky Canyons'],
    "Mekbuda": config['BiomeBanners']['Rocky Canyons'],
    "Effluvia": config['BiomeBanners']['Rocky Canyons'],
    "Pioneer II": config['BiomeBanners']['Rocky Canyons'],
    "Castor": config['BiomeBanners']['Rocky Canyons'],
    "Prasa": config['BiomeBanners']['Rocky Canyons'],
    "Kuma": config['BiomeBanners']['Rocky Canyons'],
	"Widow's Harbor": config['BiomeBanners']['Moon'],
	"RD-4": config['BiomeBanners']['Moon'],
	"Claorell": config['BiomeBanners']['Moon'],
	"Maia": config['BiomeBanners']['Moon'],
	"Curia": config['BiomeBanners']['Moon'],
	"Sirius": config['BiomeBanners']['Moon'],
	"Rasp": config['BiomeBanners']['Moon'],
	"Terrek": config['BiomeBanners']['Moon'],
	"Dolph": config['BiomeBanners']['Moon'],
	"Fenrir III": config['BiomeBanners']['Moon'],
	"Zosma": config['BiomeBanners']['Moon'],
	"Euphoria III": config['BiomeBanners']['Moon'],
	"Primordia": config['BiomeBanners']['Volcanic Jungle'],
	"Rogue 5": config['BiomeBanners']['Volcanic Jungle'],
	"Alta V": config['BiomeBanners']['Volcanic Jungle'],
	"Mantes": config['BiomeBanners']['Volcanic Jungle'],
	"Gaellivare": config['BiomeBanners']['Volcanic Jungle'],
	"Meissa": config['BiomeBanners']['Volcanic Jungle'],
	"Spherion": config['BiomeBanners']['Volcanic Jungle'],
	"Kirrik": config['BiomeBanners']['Volcanic Jungle'],
	"Baldrick Prime": config['BiomeBanners']['Volcanic Jungle'],
	"Zegema Paradise": config['BiomeBanners']['Volcanic Jungle'],
	"Irulta": config['BiomeBanners']['Volcanic Jungle'],
	"Regnus": config['BiomeBanners']['Volcanic Jungle'],
	"Navi VII": config['BiomeBanners']['Volcanic Jungle'],
	"Oasis": config['BiomeBanners']['Volcanic Jungle'],
	"Pollux 31": config['BiomeBanners']['Volcanic Jungle'],
	"Aesir Pass": config['BiomeBanners']['Deadlands'],
	"Alderidge Cove": config['BiomeBanners']['Deadlands'],
	"Penta": config['BiomeBanners']['Deadlands'],
	"Ain-5": config['BiomeBanners']['Deadlands'],
	"Skat Bay": config['BiomeBanners']['Deadlands'],
	"Alaraph": config['BiomeBanners']['Deadlands'],
	"Veil": config['BiomeBanners']['Deadlands'],
	"Troost": config['BiomeBanners']['Deadlands'],
	"Haka": config['BiomeBanners']['Deadlands'],
	"Nivel 43": config['BiomeBanners']['Deadlands'],
	"Pandion-XXIV": config['BiomeBanners']['Deadlands'],
	"Cirrus": config['BiomeBanners']['Deadlands'],
	"Mort": config['BiomeBanners']['Deadlands'],
	"Iridica": config['BiomeBanners']['Ethereal Jungle'],
	"Seyshel Beach": config['BiomeBanners']['Ethereal Jungle'],
	"Ursica XI": config['BiomeBanners']['Ethereal Jungle'],
	"Acubens Prime": config['BiomeBanners']['Ethereal Jungle'],
	"Fort Justice": config['BiomeBanners']['Ethereal Jungle'],
	"Sulfura": config['BiomeBanners']['Ethereal Jungle'],
	"Alamak VII": config['BiomeBanners']['Ethereal Jungle'],
	"Tibit": config['BiomeBanners']['Ethereal Jungle'],
	"Mordia 9": config['BiomeBanners']['Ethereal Jungle'],
	"Emorath": config['BiomeBanners']['Ethereal Jungle'],
	"Shallus": config['BiomeBanners']['Ethereal Jungle'],
	"Vindemitarix Prime": config['BiomeBanners']['Ethereal Jungle'],
	"Zefia": config['BiomeBanners']['Ethereal Jungle'],
	"Bekvam III": config['BiomeBanners']['Ethereal Jungle'],
	"Turing": config['BiomeBanners']['Ethereal Jungle'],
	"New Haven": config['BiomeBanners']['Ionic Jungle'],
	"Prosperity Falls": config['BiomeBanners']['Ionic Jungle'],
	"Veld": config['BiomeBanners']['Ionic Jungle'],
	"Malevelon Creek": config['BiomeBanners']['Ionic Jungle'],
	"Siemnot": config['BiomeBanners']['Ionic Jungle'],
	"Alairt III": config['BiomeBanners']['Ionic Jungle'],
	"Merak": config['BiomeBanners']['Ionic Jungle'],
	"Gemma": config['BiomeBanners']['Ionic Jungle'],
	"Minchir": config['BiomeBanners']['Ionic Jungle'],
	"Kuper": config['BiomeBanners']['Ionic Jungle'],
	"Brink-2": config['BiomeBanners']['Ionic Jungle'],
	"Peacock": config['BiomeBanners']['Ionic Jungle'],
	"Genesis Prime": config['BiomeBanners']['Ionic Jungle'],
	"New Kiruna": config['BiomeBanners']['Icy Glaciers'],
	"Borea": config['BiomeBanners']['Icy Glaciers'],
	"Marfark": config['BiomeBanners']['Icy Glaciers'],
	"Epsilon Phoencis VI": config['BiomeBanners']['Icy Glaciers'],
	"Kelvinor": config['BiomeBanners']['Icy Glaciers'],
	"Vog-Sojoth": config['BiomeBanners']['Icy Glaciers'],
	"Alathfar XI": config['BiomeBanners']['Icy Glaciers'],
	"Okul VI": config['BiomeBanners']['Icy Glaciers'],
	"Julheim": config['BiomeBanners']['Icy Glaciers'],
	"Hadar": config['BiomeBanners']['Icy Glaciers'],
	"Mog": config['BiomeBanners']['Icy Glaciers'],
	"Vandalon IV": config['BiomeBanners']['Icy Glaciers'],
	"Arkturus": config['BiomeBanners']['Icy Glaciers'],
	"Hesoe Prime": config['BiomeBanners']['Icy Glaciers'],
	"Vega Bay": config['BiomeBanners']['Icy Glaciers'],
	"New Stockholm": config['BiomeBanners']['Icy Glaciers'],
	"Heeth": config['BiomeBanners']['Icy Glaciers'],
	"Choepessa IV": config['BiomeBanners']['Boneyard'],
	"Martyr's Bay": config['BiomeBanners']['Boneyard'],
	"Lesath": config['BiomeBanners']['Boneyard'],
	"Cyberstan": config['BiomeBanners']['Boneyard'],
	"Deneb Secundus": config['BiomeBanners']['Boneyard'],
	"Acrux IX": config['BiomeBanners']['Boneyard'],
	"Inari": config['BiomeBanners']['Boneyard'],
	"Estanu": config['BiomeBanners']['Boneyard'],
	"Stor Tha Prime": config['BiomeBanners']['Boneyard'],
	"Halies Port": config['BiomeBanners']['Boneyard'],
	"Oslo Station": config['BiomeBanners']['Boneyard'],
	"Igla": config['BiomeBanners']['Boneyard'],
	"Krakatwo": config['BiomeBanners']['Boneyard'],
	"Grafmere": config['BiomeBanners']['Boneyard'],
	"Eukoria": config['BiomeBanners']['Boneyard'],
	"Tien Kwan": config['BiomeBanners']['Boneyard'],
	"Pathfinder V": config['BiomeBanners']['Plains'],
	"Fort Union": config['BiomeBanners']['Plains'],
	"Volterra": config['BiomeBanners']['Plains'],
	"Gemstone Bluffs": config['BiomeBanners']['Plains'],
	"Acamar IV": config['BiomeBanners']['Plains'],
	"Achernar Secundus": config['BiomeBanners']['Plains'],
	"Electra Bay": config['BiomeBanners']['Plains'],
	"Afoyay Bay": config['BiomeBanners']['Plains'],
	"Matar Bay": config['BiomeBanners']['Plains'],
	"Reaf": config['BiomeBanners']['Plains'],
	"Termadon": config['BiomeBanners']['Plains'],
	"Fenmire": config['BiomeBanners']['Plains'],
	"The Weir": config['BiomeBanners']['Plains'],
	"Bellatrix": config['BiomeBanners']['Plains'],
	"Oshaune": config['BiomeBanners']['Plains'],
	"Varylia 5": config['BiomeBanners']['Plains'],
	"Hort": config['BiomeBanners']['Plains'],
	"Draupnir": config['BiomeBanners']['Plains'],
	"Obari": config['BiomeBanners']['Plains'],
	"Mintoria": config['BiomeBanners']['Plains'],
	"Midasburg": config['BiomeBanners']['Tundra'],
	"Demiurg": config['BiomeBanners']['Tundra'],
	"Kerth Secundus": config['BiomeBanners']['Tundra'],
	"Aurora Bay": config['BiomeBanners']['Tundra'],
	"Martale": config['BiomeBanners']['Tundra'],
	"Crucible": config['BiomeBanners']['Tundra'],
	"Shelt": config['BiomeBanners']['Tundra'],
	"Trandor": config['BiomeBanners']['Tundra'],
	"Andar": config['BiomeBanners']['Tundra'],
	"Diluvia": config['BiomeBanners']['Tundra'],
	"Bunda Secundus": config['BiomeBanners']['Tundra'],
	"Ilduna Prime": config['BiomeBanners']['Tundra'],
	"Omicron": config['BiomeBanners']['Tundra'],
	"Ras Algethi": config['BiomeBanners']['Tundra'],
	"Duma Tyr": config['BiomeBanners']['Tundra'],
	"Adhara": config['BiomeBanners']['Scorched Moor'],
	"Hellmire": config['BiomeBanners']['Scorched Moor'],
	"Imber": config['BiomeBanners']['Scorched Moor'],
	"Menkent": config['BiomeBanners']['Scorched Moor'],
	"Blistica": config['BiomeBanners']['Scorched Moor'],
	"Herthon Secundus": config['BiomeBanners']['Scorched Moor'],
	"Pöpli IX": config['BiomeBanners']['Scorched Moor'],
	"Partion": config['BiomeBanners']['Scorched Moor'],
	"Wezen": config['BiomeBanners']['Scorched Moor'],
	"Marre IV": config['BiomeBanners']['Scorched Moor'],
	"Karlia": config['BiomeBanners']['Scorched Moor'],
	"Maw": config['BiomeBanners']['Scorched Moor'],
	"Kneth Port": config['BiomeBanners']['Scorched Moor'],
	"Grand Errant": config['BiomeBanners']['Scorched Moor'],
	"Fort Sanctuary": config['BiomeBanners']['Ionic Crimson'],
	"Elysian Meadows": config['BiomeBanners']['Ionic Crimson'],
	"Acrab XI": config['BiomeBanners']['Ionic Crimson'],
	"Enuliale": config['BiomeBanners']['Ionic Crimson'],
	"Liberty Ridge": config['BiomeBanners']['Ionic Crimson'],
	"Stout": config['BiomeBanners']['Ionic Crimson'],
	"Gatria": config['BiomeBanners']['Ionic Crimson'],
	"Freedom Peak": config['BiomeBanners']['Ionic Crimson'],
	"Ubanea": config['BiomeBanners']['Ionic Crimson'],
	"Valgaard": config['BiomeBanners']['Ionic Crimson'],
	"Valmox": config['BiomeBanners']['Ionic Crimson'],
	"Overgoe Prime": config['BiomeBanners']['Ionic Crimson'],
	"Providence": config['BiomeBanners']['Ionic Crimson'],
	"Kharst": config['BiomeBanners']['Ionic Crimson'],
	"Gunvald": config['BiomeBanners']['Ionic Crimson'],
	"Yed Prior": config['BiomeBanners']['Ionic Crimson'],
	"Ingmar": config['BiomeBanners']['Ionic Crimson'],
	"Crimsica": config['BiomeBanners']['Ionic Crimson'],
	"Charon Prime": config['BiomeBanners']['Ionic Crimson'],
	"Clasa": config['BiomeBanners']['Basic Swamp'],
	"Seasse": config['BiomeBanners']['Basic Swamp'],
	"Parsh": config['BiomeBanners']['Basic Swamp'],
	"East Iridium Trading Bay": config['BiomeBanners']['Basic Swamp'],
	"Gacrux": config['BiomeBanners']['Basic Swamp'],
	"Barabos": config['BiomeBanners']['Basic Swamp'],
	"Ivis": config['BiomeBanners']['Basic Swamp'],
	"Fornskogur II": config['BiomeBanners']['Basic Swamp'],
	"Nabatea Secundus": config['BiomeBanners']['Basic Swamp'],
	"Haldus": config['BiomeBanners']['Basic Swamp'],
	"Caph": config['BiomeBanners']['Basic Swamp'],
	"Bore Rock": config['BiomeBanners']['Basic Swamp'],
	"X-45": config['BiomeBanners']['Basic Swamp'],
	"Pherkad Secundus": config['BiomeBanners']['Basic Swamp'],
	"Krakabos": config['BiomeBanners']['Basic Swamp'],
	"Asperoth Prime": config['BiomeBanners']['Basic Swamp'],
	"Atrama": config['BiomeBanners']['Haunted Swamp'],
	"Setia": config['BiomeBanners']['Haunted Swamp'],
	"Tarsh": config['BiomeBanners']['Haunted Swamp'],
	"Gar Haren": config['BiomeBanners']['Haunted Swamp'],
	"Merga IV": config['BiomeBanners']['Haunted Swamp'],
	"Ratch": config['BiomeBanners']['Haunted Swamp'],
	"Bashyr": config['BiomeBanners']['Haunted Swamp'],
	"Nublaria I": config['BiomeBanners']['Haunted Swamp'],
	"Solghast": config['BiomeBanners']['Haunted Swamp'],
	"Iro": config['BiomeBanners']['Haunted Swamp'],
	"Socorro III": config['BiomeBanners']['Haunted Swamp'],
	"Khandark": config['BiomeBanners']['Haunted Swamp'],
	"Klaka 5": config['BiomeBanners']['Haunted Swamp'],
	"Skitter": config['BiomeBanners']['Haunted Swamp'],
    "Angel's Venture": config['BiomeBanners']['Fractured Planet'],
    "Moradesh": config['BiomeBanners']['Fractured Planet'],
    "Meridia": config['BiomeBanners']['Black Hole']
}

# Enemy icons for Subfactions
SUBFACTION_ICONS = {
    "Automaton Legion": config['SubfactionIcons']['AutomatonLegion'],
    "Terminid Horde": config['SubfactionIcons']['TerminidHorde'],
    "Illuminate Cult": config['SubfactionIcons']['IlluminateCult'],
    "Jet Brigade": config['SubfactionIcons']['JetBrigade'],
    "Predator Strain": config['SubfactionIcons']['PredatorStrain'],
    "Incineration Corps": config['SubfactionIcons']['IncinerationCorps'],
    "Jet Brigade & Incineration Corps": config['SubfactionIcons']['JetBrigadeIncinerationCorps'],
    "Spore Burst Strain": config['SubfactionIcons']['SporeBurstStrain'],
}

# DSS icons for Modifiers
DSS_ICONS = {
    "Eagle Storm": config['MiscIcon']['Eagle Storm'],
    "Orbital Blockade": config['MiscIcon']['Orbital Blockade'],
    "Heavy Ordnance Distribution": config['MiscIcon']['Heavy Ordnance Distribution']
}

def get_enemy_icon(enemy_type: str) -> str:
    """Get the Discord emoji icon for an enemy type."""
    return ENEMY_ICONS.get(enemy_type, "NaN")

def get_difficulty_icon(difficulty: str) -> str:
    """Get Difficulty Icons"""
    return DIFFICULTY_ICONS.get(difficulty, "NaN")

def get_planet_icon(planet: str) -> str:
    """Get Planet Icons"""
    return PLANET_ICONS.get(planet, "")

def get_system_color(enemy_type: str) -> int:
    """Get the Discord color code for an enemy type."""
    return int(SYSTEM_COLORS.get(enemy_type, "0"))

def get_campaign_icon(mission_category: str) -> str:
    """Get Campaign Icons"""
    return CAMPAIGN_ICONS.get(mission_category, "")

def get_mission_icon(mission_type: str) -> str:
    """Get Mission Icons"""
    return MISSION_ICONS.get(mission_type, "")

def get_biome_banner(planet: str) -> str:
    """Get Biome Banners"""
    return BIOME_BANNERS.get(planet, "")

def get_dss_icon(dss_modifier: str) -> str:
    """Get DSS Icons"""
    return DSS_ICONS.get(dss_modifier, "")

def normalize_subfaction_name(subfaction: str) -> str:
    """Normalize subfaction name to match config keys."""
    # Remove extra spaces and convert to title case
    normalized = " ".join(subfaction.split()).title()
    # Add any specific replacements
    replacements = {
        "Jet Brigade": "JetBrigade",
        "Predator Strain": "PredatorStrain",
        "Incineration Corps": "IncinerationCorps",
        "Jet Brigade & Incineration Corps": "JetBrigadeIncinerationCorps",
        "Spore Burst Strain": "SporeBurstStrain"
        # Add more mappings as needed
    }
    return replacements.get(normalized, normalized)

def get_subfaction_icon(subfaction_type: str) -> str:
    """Get the Discord emoji icon for subfaction type."""
    # Direct lookup without normalization
    icon = SUBFACTION_ICONS.get(subfaction_type, "NaN")
    logging.info(f"Getting subfaction icon for '{subfaction_type}', found: {icon}")
    if icon == "NaN":
        icon = ""
    return icon

def total_missions():
    df = pd.read_excel('mission_log_test.xlsx') if DEBUG else pd.read_excel('mission_log.xlsx')
    total_rows = len(df)
    return total_rows

class MissionLogGUI:
    """GUI application for logging Helldiver 2 mission data."""
    
    def __init__(self, root: tk.Tk) -> None:
        """Initialize the GUI application."""
        self.root = root
        if DEBUG:
            self.root.title("Helldiver Mission Log Manager V-{} DEBUG:{}".format(VERSION, DEBUG))
        else:
            self.root.title("Helldiver Mission Log Manager V-{}".format(VERSION))
        
        self.root.resizable(False, False)
        
        # Load icon in a separate thread
        def load_icon():
            try:
                icon = tk.PhotoImage(file='SuperEarth.png')
                self.root.after(0, lambda: self.root.iconphoto(False, icon))
            except Exception as e:
                logging.error(f"Failed to load icon: {e}")
        
        threading.Thread(target=load_icon, daemon=True).start()
        
        self.settings_file = SETTINGS_FILE
        self._setup_variables()
        self._setup_discord_rpc()
        self._create_main_frame()
        self._setup_ui()
        
        # Delay loading settings
        self.root.after(100, self.load_settings)
        
        # Start RPC updates after a short delay
        self.root.after(1000, self._periodic_rpc_update)

    def _periodic_rpc_update(self) -> None:
        """Periodically update Discord Rich Presence."""
        try:
            self._update_discord_presence()
        except Exception as e:
            logging.error(f"Error in periodic RPC update: {e}")
        finally:
            self.root.after(RPC_UPDATE_INTERVAL * 1000, self._periodic_rpc_update)

    def _setup_variables(self) -> None:
        """Initialize tkinter variables with validation."""
        self.sector = tk.StringVar()
        self.planet = tk.StringVar()
        self.mission_type = tk.StringVar()
        self.kills = tk.StringVar()
        self.deaths = tk.StringVar()
        self.enemy_type = tk.StringVar()
        self.subfaction_type = tk.StringVar()
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
        self.note = tk.StringVar()
        self.shipName1 = tk.StringVar()
        self.shipName2 = tk.StringVar()
        self.FullShipName = tk.StringVar()

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
        def init_rpc():
            try:
                self.RPC = Presence(DISCORD_CLIENT_ID)
                self.RPC.connect()
                self.last_rpc_update = time.time()  # Initialize the timestamp
                logging.info("Discord Rich Presence initialized successfully")
            except Exception as e:
                logging.error(f"Failed to initialize Discord Rich Presence: {e}")
                self.RPC = None
                self.last_rpc_update = 0  # Set default value even if connection fails

        # Run Discord RPC initialization in a separate thread
        threading.Thread(target=init_rpc, daemon=True).start()

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
        ttk.Label(mission_frame, text="Destroyer Name:").grid(row=0, column=0, sticky=tk.W, pady=5)

        self.shipName1s = ["SES Adjudicator", "SES Advocate", "SES Aegis", "SES Agent", "SES Arbiter", "SES Banner", "SES Beacon", "SES Blade", "SES Bringer", "SES Champion", "SES Citizen", "SES Claw", "SES Colossus", "SES Comptroller", "SES Courier", "SES Custodian", "SES Dawn", "SES Defender", "SES Diamond", "SES Distributor", "SES Dream", "SES Elected Representative", "SES Emperor", "SES Executor", "SES Eye", "SES Father", "SES Fist", "SES Flame", "SES Force", "SES Forerunner", "SES Founding Father", "SES Gauntlet", "SES Giant", "SES Guardian", "SES Halo", "SES Hammer", "SES Harbinger", "SES Herald", "SES Judge", "SES Keeper", "SES King", "SES Knight", "SES Lady", "SES Legislator", "SES Leviathan", "SES Light", "SES Lord", "SES Magistrate", "SES Marshall", "SES Martyr", "SES Mirror", "SES Mother", "SES Octagon", "SES Ombudsman", "SES Panther", "SES Paragon", "SES Patriot", "SES Pledge", "SES Power", "SES Precursor", "SES Pride", "SES Prince", "SES Princess", "SES Progenitor", "SES Prophet", "SES Protector", "SES Purveyor", "SES Queen", "SES Ranger", "SES Reign", "SES Representative", "SES Senator", "SES Sentinel", "SES Shield", "SES Soldier", "SES Song", "SES Soul", "SES Sovereign", "SES Spear", "SES Stallion", "SES Star", "SES Steward", "SES Superintendent", "SES Sword", "SES Titan", "SES Triumph", "SES Warrior", "SES Whisper", "SES Will", "SES Wings"]
        self.shipName2s = ["of Allegiance", "of Audacity", "of Authority", "of Battle", "of Benevolence", "of Conquest", "of Conviction", "of Conviviality", "of Courage", "of Dawn", "of Democracy", "of Destiny", "of Destruction", "of Determination", "of Equality", "of Eternity", "of Family Values", "of Fortitude", "of Freedom", "of Glory", "of Gold", "of Honour", "of Humankind", "of Independence", "of Individual Merit", "of Integrity", "of Iron", "of Judgement", "of Justice", "of Law", "of Liberty", "of Mercy", "of Midnight", "of Morality", "of Morning", "of Opportunity", "of Patriotism", "of Peace", "of Perseverance", "of Pride", "of Redemption", "of Science", "of Self-Determination", "of Selfless Service", "of Serenity", "of Starlight", "of Steel", "of Super Earth", "of Supremacy", "of the Constitution", "of the People", "of the Regime", "of the Stars", "of the State", "of Truth", "of Twilight", "of Victory", "of Vigilance", "of War", "of Wrath"]

        self.ship1_combo = ttk.Combobox(mission_frame, textvariable=self.shipName1, values=self.shipName1s, state='readonly', width=27)
        self.ship1_combo.grid(row=0, column=1, padx=5, pady=5)
        self.ship1_combo.set(self.shipName1s[0])

        self.ship2_combo = ttk.Combobox(mission_frame, textvariable=self.shipName2, values=self.shipName2s, state='readonly', width=27)
        self.ship2_combo.grid(row=0, column=2, padx=5, pady=5)
        self.ship2_combo.set(self.shipName2s[0])

        def update_full_ship_name(*args):
            self.FullShipName.set(f"{self.shipName1.get()} {self.shipName2.get()}")

        self.shipName1.trace_add("write", update_full_ship_name)
        self.shipName2.trace_add("write", update_full_ship_name)
        update_full_ship_name()

        ttk.Label(mission_frame, text="Helldiver:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(mission_frame, textvariable=self.Helldivers, width=30).grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(mission_frame, text="Level:").grid(row=2, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(mission_frame, textvariable=self.level, width=10).grid(row=2, column=2, sticky=tk.E, padx=(0,5), pady=5)

        ttk.Label(mission_frame, text="Title:").grid(row=2, column=4, sticky=tk.W, pady=5)
        self.titles = ['CADET', 'SPACE CADET', 'SERGEANT', 'MASTER SERGEANT', 'CHIEF', 'SPACE CHIEF PRIME', 
             'DEATH CAPTAIN', 'MARSHALL', 'STAR MARSHALL', 'ADMIRAL', 'SKULL ADMIRAL', 'FLEET ADMIRAL',
             'ADMIRABLE ADMIRAL', 'COMMANDER', 'GALACTIC COMMANDER', 'HELL COMMANDER', 'GENERAL',
             '5-STAR GENERAL', '10-STAR GENERAL', 'PRIVATE', 'SUPER PRIVATE', 'SUPER CITIZEN',
             'VIPER COMMANDO', 'FIRE SAFETY OFFICER', 'EXPERT EXTERMINATOR', 'FREE OF THOUGHT',
             'ASSAULT INFANTRY', 'SUPER PEDESTRIAN', 'SERVANT OF FREEDOM', 'SUPER SHERIFF', 'DECORATED HERO']
        self.title_combo = ttk.Combobox(mission_frame, textvariable=self.title, state='readonly', width=27)
        self.title_combo['values'] = self.titles
        self.title_combo.grid(row=2, column=5, padx=5, pady=5)
        self.title_combo.set(self.titles[0])

        ttk.Label(mission_frame, text="Sector:").grid(row=3, column=0, sticky=tk.W, pady=5)
        sector_combo = ttk.Combobox(mission_frame, textvariable=self.sector, values=sector_list, state='readonly', width=27)
        sector_combo.grid(row=3, column=1, padx=5, pady=5)
        sector_combo.set(sector_list[0])

        ttk.Label(mission_frame, text="Planet:").grid(row=4, column=0, sticky=tk.W, pady=5)
        planet_combo = ttk.Combobox(mission_frame, textvariable=self.planet, state='readonly', width=27)
        planet_combo.grid(row=4, column=1, padx=5, pady=5)
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

        # Major Order checkbox
        ttk.Label(details_frame, text="Major Order:").grid(row=3, column=2, sticky=tk.W, pady=5)
        ttk.Checkbutton(details_frame, variable=self.MO).grid(row=3, column=3, padx=5, pady=5)

        # DSS Modifier dropdown
        ttk.Label(details_frame, text="DSS Active:").grid(row=1, column=2, sticky=tk.W, pady=5)
        ttk.Checkbutton(details_frame, variable=self.DSS).grid(row=1, column=3, padx=5, pady=5)

        self.dss_frame = ttk.Frame(details_frame)
        self.dss_frame.grid(row=1, column=4, columnspan=2, sticky=tk.W, pady=5)
        ttk.Label(self.dss_frame, text="DSS Modifier:").pack(side=tk.LEFT)
        dss_mods = ["Inactive", "Orbital Blockade", "Heavy Ordnance Distribution", "Eagle Storm"]
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

        # Subfaction Selection
        ttk.Label(details_frame, text="Enemy Subfaction:").grid(row=0, column=2, sticky=tk.W, pady=5)
        subfaction_combo = ttk.Combobox(details_frame, textvariable=self.subfaction_type, state='readonly', width=27)
        subfaction_combo.grid(row=0, column=3, padx=5, pady=5)

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

        # Biome Banner
        def display_image_from_url(biome_banner):
            with urllib.request.urlopen(biome_banner) as u:
                raw_data = u.read()

            planet_image = Image.open(io.BytesIO(raw_data))
            planet_photo = ImageTk.PhotoImage(planet_image)

            planet_label = tk.Label(content, planet_image=planet_photo)
            planet_label.pack()

        def update_subfactions(*args):
            enemy = self.enemy_type.get()
            if enemy in missions_data:
                subfactions = list(missions_data[enemy].keys())
                logging.info(f"Available subfactions for {enemy}: {subfactions}")  # Add logging
                subfaction_combo['values'] = subfactions
            if subfactions:
                subfaction_combo.set(subfactions[0])
                update_mission_categories()

        def update_mission_categories(*args):
            enemy = self.enemy_type.get()
            subfaction = self.subfaction_type.get()
            if enemy in missions_data and subfaction in missions_data[enemy]:
                categories = list(missions_data[enemy][subfaction].keys())
                mission_cat_combo['values'] = categories
            if categories:
                mission_cat_combo.set(categories[0])
                update_mission_types()

        def update_mission_types(*args):
            enemy = self.enemy_type.get()
            subfaction = self.subfaction_type.get()
            category = self.mission_category.get()
            
            if (enemy in missions_data and 
                subfaction in missions_data[enemy] and 
                category in missions_data[enemy][subfaction]):
                
                if missions_data[enemy][subfaction][category] != "Unknown":
                    difficulties = list(missions_data[enemy][subfaction][category].keys())
                    difficulty_combo['values'] = difficulties
                    difficulty_combo.set(difficulties[0])
                    
                    # Set initial mission types from first difficulty
                    first_difficulty = difficulties[0]
                    available_missions = missions_data[enemy][subfaction][category][first_difficulty]
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
            subfaction = self.subfaction_type.get()
            category = self.mission_category.get()
            difficulty = self.difficulty.get()
            
            if (enemy in missions_data and 
                subfaction in missions_data[enemy] and 
                category in missions_data[enemy][subfaction] and 
                difficulty in missions_data[enemy][subfaction][category]):
                
                available_missions = missions_data[enemy][subfaction][category][difficulty]
                mission_type_combo['values'] = available_missions
                if available_missions:
                    mission_type_combo.set(available_missions[0])

        enemy_combo.bind('<<ComboboxSelected>>', update_subfactions)
        subfaction_combo.bind('<<ComboboxSelected>>', update_mission_categories)
        mission_cat_combo.bind('<<ComboboxSelected>>', update_mission_types)
        difficulty_combo.bind('<<ComboboxSelected>>', update_available_missions)

        # Initial setup
        update_subfactions()

        # Statistics and Note Section Container
        stats_note_container = ttk.Frame(content)
        stats_note_container.grid(row=2, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

        # Statistics Section (Left)
        stats_frame = ttk.LabelFrame(stats_note_container, text="Mission Statistics", padding=10)
        stats_frame.pack(side=tk.LEFT, padx=5, fill=tk.BOTH, expand=True)

        # Note Section (Right)
        note_frame = ttk.LabelFrame(stats_note_container, text="Note", padding=10)
        note_frame.pack(side=tk.RIGHT, padx=5, fill=tk.BOTH, expand=True)

        note_entry = tk.Text(note_frame, height=3, width=30)
        note_entry.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        # Add a trace to update self.note when text changes
        def update_note(*args):
            self.note.set(note_entry.get("1.0", "end-1c"))
        note_entry.bind('<KeyRelease>', lambda e: update_note())
        note_frame.columnconfigure(0, weight=1)  # Make the column expand horizontally
        note_frame.rowconfigure(0, weight=1)  # Make the row expand vertically

        ttk.Label(stats_frame, text="Kills:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(stats_frame, textvariable=self.kills, width=30).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(stats_frame, text="Deaths:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(stats_frame, textvariable=self.deaths, width=30).grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(stats_frame, text="Performance:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ratings = ["Outstanding Patriotism", "Superior Valour", "Costly Failure", "Honourable Duty", "Unremarkable Performance", "Disappointing Service", "Disgraceful Conduct"]
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

        #export button by planet
        export_button = ttk.Button(export_frame, text="Export Excel Data to Webhook", command=lambda: os.system('python sub.py'))
        export_button.grid(row=6, column=0, pady=15)

        #export button by faction
        export_button = ttk.Button(export_frame, text="Export Faction Data to Webhook", command=lambda: os.system('python faction.py'))
        export_button.grid(row=7, column=0, pady=15)


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
                "Illuminate": "squids",
                "Observing": "obs"
            }
            
            small_image = enemy_assets.get(enemytype, "unknown")

            if self.enemy_type.get() == "Observing":
                SText = "Observing"
            else:
                SText = "Fighting: " + enemytype


            self.RPC.update(
                state=f"Sector: {sector}\nPlanet: {planet}",
                details=f"Helldiver: {helldiver} Level: {level} | {title}",
                large_image="superearth",
                large_text="Helldivers 2",
                small_image=small_image,
                small_text=f"{SText}",
            )
            self.last_rpc_update = current_time
        except Exception as e:
            logging.error(f"Failed to update Discord Rich Presence: {e}")

    def load_settings(self) -> None:
        """Load user settings from file."""
        def load():
            try:
                if os.path.exists(self.settings_file):
                    with open(self.settings_file, 'r') as f:
                        settings = json.load(f)
                        self.root.after(0, lambda: self._apply_settings(settings))
            except Exception as e:
                self.root.after(0, lambda: self._show_error(f"Error loading settings: {e}"))

        threading.Thread(target=load, daemon=True).start()

    def _apply_settings(self, settings: Dict) -> None:
        """Apply loaded settings to UI elements."""
        self.Helldivers.set(settings.get('helldiver', ''))
        self.level.set(settings.get('level', 0))
        self.difficulty.set(settings.get('difficulty', '1 - TRIVIAL'))
        self.mission_category.set(settings.get('campaign', 'Ivasion'))
        self.DSS.set(settings.get('DSS', False))
        self.shipName1.set(settings.get('shipName1', 'SES Adjudicator'))
        self.shipName2.set(settings.get('shipName2', 'of Allegiance'))


        if settings.get('DSS'):
            self.DSSMod.set(settings.get('DSSMod', 'Inactive'))
        
        # For mission type
        if settings.get('mission'):
            self.mission_type.set(settings.get('mission'))

        if settings.get('title') in self.titles:
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
            'campaign': self.mission_category.get(),
            'subfaction': self.subfaction_type.get(),
            'shipName1': self.shipName1.get(),
            'shipName2': self.shipName2.get()
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

        if self.enemy_type.get() == "Observing":
            self._show_error("ADVISORY: You cannot submit an observation mission")
            return

        if self.mission_type.get() == "No missions available":
            self._show_error("ADVISORY: You cannot submit without selecting a mission")
            return

        if self.planet.get() == "Meridia":
            self._show_error("ADVISORY: Volatile spacetime fluctuations currently prohibit FTL travel to the Meridian Black Hole.")
            return

        if self.planet.get() == "Angel's Venture" or self.planet.get() == "Moradesh":
            self._show_error("ADVISORY: You cannot deploy on a fractured planet")
            return

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
            #create a randint between 1 and 0
            rndint = random.randint(0, 1)

            if level < 1 or level > 150:  # Add reasonable level range
                if rndint == 1:
                    self._show_error("ADVISORY: You are not a Helldiver")
                else:
                    self._show_error("Level must be between 1 and 150")
                return False

            if kills < 0 or kills > 10000:
                if rndint == 1:
                    self._show_error("These kills will be reported to your Democracy Officer... I dear hope you're not lying...")
                else:
                    self._show_error("Invalid number of kills")
                return False

            if deaths < 0 or deaths > 1000:
                if rndint == 1:
                    self._show_error("Surely you didn't die this many times... right?")
                else:
                    self._show_error("Invalid number of deaths")
                return False

            # Validate required text fields
            if not self.Helldivers.get().strip():
                if rndint == 1:
                    self._show_error("I know we're cannon fodder but you could at least give yourself a name... have some dignity!")
                else:
                    self._show_error("Helldiver name is required")
                return False

            if not self.mission_type.get().strip():
                if rndint == 1:
                    self._show_error("Did you sit in your hellpod the entire time?")
                else:
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
        return {
            'Super Destroyer': self.FullShipName.get(),
            'Helldivers': self.Helldivers.get(),
            'Level': self.level.get(),
            'Title': self.title.get(),
            'Sector': self.sector.get(),
            'Planet': self.planet.get(),
            'Enemy Type': self.enemy_type.get(),
            'Enemy Subfaction': self.subfaction_type.get(),
            'Major Order': self.MO.get(),
            'DSS Active': self.DSS.get(),
            'DSS Modifier': self.DSSMod.get(),
            'Mission Category': self.mission_category.get(),
            'Mission Type': self.mission_type.get(),
            'Difficulty': self.difficulty.get(),
            'Kills': self.kills.get(),
            'Deaths': self.deaths.get(),
            'Rating': self.rating.get(),
            'Time': datetime.now().strftime(DATE_FORMAT),
            'Note': self.note.get(),
        }


        # Save to Excel
        self._save_to_excel(data)

        # Send to Discord
        self._send_to_discord(data)

    def _save_to_excel(self, data: Dict) -> bool:
        """Save mission data to Excel file by appending new rows."""
        excel_file = EXCEL_FILE_TEST if DEBUG else EXCEL_FILE_PROD
        
        try:
            # Create new DataFrame with single row of data
            new_data = pd.DataFrame([data])
            
            # If file exists, read and append. Otherwise create new file
            if os.path.exists(excel_file):
                existing_df = pd.read_excel(excel_file)
                updated_df = pd.concat([existing_df, new_data], ignore_index=True)
            else:
                updated_df = new_data
                
            # Save the updated DataFrame back to Excel
            with pd.ExcelWriter(excel_file) as writer:
                updated_df.to_excel(writer, index=False)
                
            logging.info(f"Successfully appended data to {excel_file}")
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
            
            # Map ratings to number of gold stars
            rating_stars = {
                "Outstanding Patriotism": 5,
                "Superior Valour": 4,
                "Costly Failure": 4,
                "Honourable Duty": 3,
                "Unremarkable Performance": 2,
                "Disappointing Service": 1,
                "Disgraceful Conduct": 0
            }
            
            gold_count = rating_stars.get(self.rating.get(), 0)
            Stars = GoldStar * gold_count + GreyStar * (5 - gold_count)
            date = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            enemy_icon = get_enemy_icon(data['Enemy Type'])
            planet_icon = get_planet_icon(data['Planet'])
            if planet_icon == '':
                planet_icon = '<:hd1superearth:1103949794285723658>'
            system_color = get_system_color(data['Enemy Type'])
            diff_icon = get_difficulty_icon(data['Difficulty'])
            subfaction_icon = get_subfaction_icon(data['Enemy Subfaction'])
            campaign_icon = get_campaign_icon(data['Mission Category'])
            mission_icon = get_mission_icon(data['Mission Type'])
            biome_banner = get_biome_banner(data['Planet'])
            dss_icon = get_dss_icon(data['DSS Modifier'])


            # Logic to track streaks (submissions within an hour)
            helldiver_name = data['Helldivers']
            streak_file = 'streak_data.json'
            streak = 1  # Default streak value
            highest_streak = 0  # Default highest streak value
            streak_emoji = ""  # No streak emoji by default

            try:
                # Load streak data for all users
                streak_data = {}
                if os.path.exists(streak_file):
                    with open(streak_file, 'r') as f:
                        streak_data = json.load(f)
                
                user_data = streak_data.get(helldiver_name, {'streak': 0, 'last_time': None})
                
                # Check if this user has previous streak data with valid timestamp
                if user_data['last_time']:
                    last_time = datetime.strptime(user_data['last_time'], "%Y-%m-%d %H:%M:%S")
                    time_diff = datetime.now() - last_time
                    
                    if time_diff.total_seconds() <= 3600:  # Within an hour
                        streak = user_data['streak'] + 1
                        # Add streak emoji based on length
                        if streak >= 30:
                            streak_emoji = "🔥 x" + str(streak) + " WTF!"
                        elif streak >= 24:
                            streak_emoji = "🔥 x" + str(streak) + " TRULY HELLDIVING!"
                        elif streak >= 21:
                            streak_emoji = "🔥 x" + str(streak) + " IMPOSSIBLE!"
                        elif streak >= 18:
                            streak_emoji = "🔥 x" + str(streak) + " SUICIDAL!"
                        elif streak >= 15:
                            streak_emoji = "🔥 x" + str(streak) + " PATRIOTIC!"
                        elif streak >= 12:
                            streak_emoji = "🔥 x" + str(streak) + " DEMOCRATIC!"
                        elif streak >= 9:
                            streak_emoji = "🔥 x" + str(streak) + " LIBERATING!"
                        elif streak >= 6:
                            streak_emoji = "🔥 x" + str(streak) + " SUPER!"
                        elif streak >= 3:
                            streak_emoji = "🔥 x" + str(streak) + " COMMENDABLE!"
                        else:
                            streak_emoji = "🔥 x" + str(streak)
                # Update streak data for this user
                highest_streak = user_data.get('highest_streak', 0)  # Get existing highest streak or default to 0
                if streak > highest_streak:
                    highest_streak = streak
                    
                streak_data[helldiver_name] = {
                    'streak': streak,
                    'highest_streak': highest_streak,
                    'last_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                
                # Save updated streak data
                with open(streak_file, 'w') as f:
                    json.dump(streak_data, f, indent=4)
                    
            except Exception as e:
                logging.error(f"Error managing streak data: {e}")

            total_missions_main = total_missions()

            # Format the message for Discord
            message_content1 = (
                f"# ================================\n"
                f"> # | **Date:** {date}\n"
                f"> # | **Mission Report for {data['Helldivers']}**\n"
                f"> # --------------------------------\n"
                f"> # | **Sector:** {data['Sector']}\n"
                f"> ## | **Planet:** {data['Planet']} {planet_icon}\n"
                f"> ## | **Enemy Type:** {data['Enemy Type']} {enemy_icon}\n"
                f"> ## | **Mission Category:** {data['Mission Category']} {campaign_icon}\n"
                f"> ## | **Mission Type:** {data['Mission Type']} {mission_icon}\n"
                f"> ## | **Mission Difficulty:** {data['Difficulty']}\n"
                f"> # --------------------------------\n"
                f"> ### | **Kills:** {data['Kills']}\n"
                f"> ### | **Deaths:** {data['Deaths']}\n"
                f"# ================================\n"
            )

            UID = config['Discord']['UID']
            MICo = str(data["Major Order"]) + " " + config['MiscIcon']['MO'] if data["Major Order"] else str(data["Major Order"])
            DSSIco = str(data["DSS Active"]) + " " + config['MiscIcon']['DSS'] if data["DSS Active"] else str(data["DSS Active"])

            message_content2 = {
                "content": None,
                "embeds": [{
                    "title": f"{data['Super Destroyer']}\nDeployed {data['Helldivers']}\nLevel {data['Level']} | {data['Title']}\nMission: {total_missions_main}",
                    "description": f"<a:easyshine1:1349110651829747773> <:hd1superearth:1103949794285723658> **Galactic Intel** {planet_icon} <a:easyshine3:1349110648528699422>\n> Sector: {data['Sector']}\n> Planet: {data['Planet']}\n> Major Order: {MICo}\n> DSS Active: {DSSIco}\n> DSS Modifier: {data['DSS Modifier']} {dss_icon}\n\n",
                    "color": system_color,
                    "fields": [{
                        "name": f"<a:easyshine1:1349110651829747773> {enemy_icon} **Enemy Intel** {subfaction_icon} <a:easyshine3:1349110648528699422>",
                        "value": f"> Faction: {data['Enemy Type']}\n> Subfaction: {data['Enemy Subfaction']}\n> Campaign: {data['Mission Category']}\n\n<a:easyshine1:1349110651829747773> {campaign_icon} **Mission Intel** {mission_icon} <a:easyshine3:1349110648528699422>\n> Mission: {data['Mission Type']}\n> Difficulty: {data['Difficulty']} {diff_icon}\n> Kills: {data['Kills']}\n> Deaths: {data['Deaths']}\n> Rating: {data['Rating']}\n\n {Stars}\n"
                    }],
                    "author": {
                        "name": f"Super Earth Mission Report\nDate: {date}",
                        "icon_url": "https://cdn.discordapp.com/attachments/1340508329977446484/1356001307596427564/NwNzS9B.png?ex=67eafa21&is=67e9a8a1&hm=7e204265cbcdeaf96d7b244cd63992c4ef10dc18befbcf2ed39c3a269af14ec0&"
                    },
                    "footer": {
                    "text": f"{UID}  {streak_emoji}",
                    "icon_url": "https://cdn.discordapp.com/attachments/1340508329977446484/1356025859319926784/5cwgI15.png?ex=67eb10fe&is=67e9bf7e&hm=ab6326a9da1e76125238bf3668acac8ad1e43b24947fc6d878d7b94c8a60ab28&"
                    },
                    "image": {"url": f"{biome_banner}"},
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

    def export_excel(self):
        try:
            subprocess.run(['python', 'sub.py'], 
                          shell=False,
                          check=True,
                          capture_output=True)
        except subprocess.CalledProcessError as e:
            self._show_error(f"Export failed: {e}")

if __name__ == "__main__":
    if not (re.match(r'^\d{17,19}$', config['Discord']['UID']) or (DEBUG and config['Discord']['UID'] == '0')):
        print("Please set a valid Discord ID in the config file")
        messagebox.showerror("Error", "Please set a valid Discord ID in the config file")
        os._exit(1)

    root = tk.Tk()
    app = MissionLogGUI(root)
    root.mainloop()