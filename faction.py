import pandas as pd
import configparser
import requests
import json

#Constants
DEBUG = config.getboolean('DEBUGGING', 'DEBUG', fallback=False)

# Read config file
config = configparser.ConfigParser()
config.read('config.config')

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

# Enemy icons and colors from config
ENEMY_ICONS = {
    "Automatons": config['EnemyIcons']['Automatons'],
    "Terminids": config['EnemyIcons']['Terminids'],
    "Illuminate": config['EnemyIcons']['Illuminate'],
    "Observing": config['EnemyIcons']['Observation'],
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

# Read the Excel file
try:
    df = pd.read_excel('mission_log_test.xlsx') if DEBUG else pd.read_excel('mission_log.xlsx')
except FileNotFoundError:
    print("Error: Excel file not found. Please ensure the file exists in the correct location.")
    exit(1)

# Initialize a dictionary to store column totals
sectors = []
planets = []
enemy_types = []
MissionCategory = []
difficulties = []

# Get total number of rows
total_rows = len(df)
max_rating = total_rows * 5
# Initialize counter for rating
total_rating = 0
# Create rating mapping
rating_mapping = {"Outstanding Patriotism": 5, "Superior Valour": 4, "Costly Failure": 4, "Honourable Duty":3, "Unremarkable Performance":2, "Dissapointing Service":1, "Disgraceful Conduct":0}
# Calculate total rating
total_rating = sum(rating_mapping[row["Rating"]] for index, row in df.iterrows() if "Rating" in df.columns and row["Rating"] in rating_mapping)
Rating_Percentage = (total_rating / max_rating) * 100

# Get the user's name and level from the last row of the DataFrame
helldiver_ses = df['Super Destroyer'].iloc[-1] if 'Super Destroyer' in df.columns else "Unknown"
helldiver_name = df['Helldivers'].iloc[-1] if 'Helldivers' in df.columns else "Unknown"
helldiver_level = df['Level'].iloc[-1] if 'Level' in df.columns else 0
helldiver_title = df['Title'].iloc[-1] if 'Title' in df.columns else "Unknown"


if Rating_Percentage >= 90:
    Rating = "Outstanding Patriotism"
elif Rating_Percentage >= 70:
    Rating = "Superior Valour"
elif Rating_Percentage >= 50:
    Rating = "Honourable Duty"
elif Rating_Percentage >= 30:
    Rating = "Unremarkable Performance"
elif Rating_Percentage >= 10:
    Rating = "Dissapointing Service"
else:
    Rating = "Disgraceful Conduct"

# Iterate through each row
for index, row in df.iterrows():
    # Append Sector values to the list
    if "Sector" in df.columns and row["Sector"] not in sectors:
        sectors.append(row["Sector"])

    # Append Planet values to the list
    if "Planet" in df.columns and row["Planet"] not in planets:
        planets.append(row["Planet"])

    # Append Enemy Type values to the list
    if "Enemy Type" in df.columns and row["Enemy Type"] not in enemy_types:
        enemy_types.append(row["Enemy Type"])
    
    # Append Category values to the list
    if "Mission Category" in df.columns and row["Mission Category"] not in MissionCategory:
        MissionCategory.append(row["Mission Category"])
    
    # Append Difficulty values to the list
    if "Difficulty" in df.columns and row["Difficulty"] not in difficulties:
        difficulties.append(row["Difficulty"])

# Initialize lists to store stats for each planet
planet_kills_list = []
planet_deaths_list = []
planet_orders_list = []

for Planets in planets:
    # Filter data for this planet and sum stats
    planet_data = df[df["Planet"] == Planets]
    planet_kills = planet_data["Kills"].sum()
    planet_deaths = planet_data["Deaths"].sum()
    planet_major_orders = planet_data["Major Order"].astype(int).sum()
    planet_last_date = planet_data["Time"].max() if "Time" in df.columns else "No date available"
    planet_deployments = len(planet_data)
    
    # Create dictionaries to store data for each planet if they don't exist
    if 'planet_data_dict' not in locals():
        planet_data_dict = {}
        planet_kills_dict = {}
        planet_deaths_dict = {}
        planet_orders_dict = {}
        planet_last_date_dict = {}
        planet_deployments_dict = {}
    
    # Store data in dictionaries with planet name as key
    planet_data_dict[Planets] = planet_data
    planet_kills_dict[Planets] = planet_kills
    planet_deaths_dict[Planets] = planet_deaths
    planet_orders_dict[Planets] = planet_major_orders
    planet_last_date_dict[Planets] = planet_last_date
    planet_deployments_dict[Planets] = planet_deployments

# Create a DataFrame from the planet stats
planet_stats_df = pd.DataFrame({
    "Planet": planets,
    "Total Kills": [planet_kills_dict[planet] for planet in planets],
    "Total Deaths": [planet_deaths_dict[planet] for planet in planets],
    "Major Orders": [planet_orders_dict[planet] for planet in planets],
    "Last Date": [planet_last_date_dict[planet] for planet in planets]
})

# Discord webhook configuration
WEBHOOK_URLS = {
    'PROD': config['Webhooks']['BAT'].split(','),
    'TEST': config['Webhooks']['TEST'].split(',')
}
ACTIVE_WEBHOOK = WEBHOOK_URLS['TEST'] if DEBUG else WEBHOOK_URLS['PROD']
UID = config['Discord']['UID']

# Get latest note
non_blank_notes = df['Note'].dropna()
latest_note = non_blank_notes.iloc[-1] if not non_blank_notes.empty else "No Quote"

# Get Instances from Data
search_mission = df['Mission Type'].mode()[0]
MissionCount = df.apply(lambda row: row.astype(str).str.contains(search_mission, case=False).sum(), axis=1).sum()
search_campaign = df['Mission Category'].mode()[0]
CampaignCount = df.apply(lambda row: row.astype(str).str.contains(search_campaign, case=False).sum(), axis=1).sum()
search_faction = df['Enemy Type'].mode()[0]
FactionCount = df.apply(lambda row: row.astype(str).str.contains(search_faction, case=False).sum(), axis=1).sum()
search_difficulty = df['Difficulty'].mode()[0]
DifficultyCount = df.apply(lambda row: row.astype(str).str.contains(search_difficulty, case=False).sum(), axis=1).sum()
search_planet = df['Planet'].mode()[0]
PlanetCount = df.apply(lambda row: row.astype(str).str.contains(search_planet, case=False).sum(), axis=1).sum()
search_sector = df['Sector'].mode()[0]
SectorCount = df.apply(lambda row: row.astype(str).str.contains(search_sector, case=False).sum(), axis=1).sum()

# Create embed data
embed_data = {
    "content": None,
    "embeds": [
        {
            "title": "",  # Empty title, will be set below
            "description": f"\"{latest_note}\"\n\n<a:easyshine1:1349110651829747773>  <a:easyshine2:1349110649753698305> Combat Statistics <a:easyshine2:1349110649753698305> <a:easyshine3:1349110648528699422>\n" + 
                        f"> Kills - {df['Kills'].sum()}\n" +
                        f"> Deaths - {df['Deaths'].sum()}\n" +
                        f"> Highest Kills in Mission - {df['Kills'].max()}\n" +

                        f"\n<a:easyshine1:1349110651829747773>  <a:easysuperearth:1343266082881802443> Mission Statistics <a:easysuperearth:1343266082881802443> <a:easyshine3:1349110648528699422>\n" + 
                        f"> Deployments - {len(df)}\n" +
                        f"> Major Order Deployments - {df['Major Order'].astype(int).sum()}\n" +
                        f"> DSS Deployments - {df['DSS Active'].astype(int).sum()}\n" +

                        f"\n<a:easyshine1:1349110651829747773>  <a:easyskullgold:1232018045791375360> Performance Statistics <a:easyskullgold:1232018045791375360> <a:easyshine3:1349110648528699422>\n" +                      
                        f"> Rating - {Rating} | {int(Rating_Percentage)}%\n" +

                        f"\n<a:easyshine1:1349110651829747773>  <:goldstar:1337818552094163034> Favourites <:goldstar:1337818552094163034> <a:easyshine3:1349110648528699422>\n" +     
                        f"> Mission - {df['Mission Type'].mode()[0]} {MISSION_ICONS.get(df['Mission Type'].mode()[0], '')} (x{MissionCount})\n" +
                        f"> Campaign - {df['Mission Category'].mode()[0]} {CAMPAIGN_ICONS.get(df['Mission Category'].mode()[0], '')} (x{CampaignCount})\n" +
                        f"> Faction - {df['Enemy Type'].mode()[0]} {ENEMY_ICONS.get(df['Enemy Type'].mode()[0], '')} (x{FactionCount})\n" +
                        f"> Difficulty - {df['Difficulty'].mode()[0]} {DIFFICULTY_ICONS.get(df['Difficulty'].mode()[0], '')} (x{DifficultyCount})\n" +
                        f"> Planet - {df['Planet'].mode()[0]} {PLANET_ICONS.get(df['Planet'].mode()[0], '')} (x{PlanetCount})\n" +
                        f"> Sector - {df['Sector'].mode()[0]} (x{SectorCount})\n",
            "color": 7257043,
            "author": {"name": "SEAF Battle Record"},
            "footer": {"text": config['Discord']['UID'],"icon_url": "https://cdn.discordapp.com/attachments/1340508329977446484/1356025859319926784/5cwgI15.png?ex=67eb10fe&is=67e9bf7e&hm=ab6326a9da1e76125238bf3668acac8ad1e43b24947fc6d878d7b94c8a60ab28&"},
            "image": {"url": f"{BIOME_BANNERS.get(df['Planet'].mode()[0], '')}"},
            "thumbnail": {"url": "https://i.ibb.co/5g2b9NXb/Super-Earth-Icon.png"}
        },
        {
      "title": "Terminids Campaign Record",
      "description": f"<a:easyshine1:1349110651829747773> <:hd2bugs:1337190441170370693> Terminid Front Statistics <:hd2bugs:1337190441170370693> <a:easyshine3:1349110648528699422>\n" +
                         f"> <:resistance:1370883421496148068> Kills - {df[df['Enemy Type'] == 'Terminids']['Kills'].sum()}\n" +
                         f"> <:helldiverHD:1370887412443648070> Deaths - {df[df['Enemy Type'] == 'Terminids']['Deaths'].sum()}\n" +
                         f"> <:highprioritytarget:1370882658019704903> Highest Kills in Mission - {df[df['Enemy Type'] == 'Terminids']['Kills'].max()}\n\n" +

                         f"> <:deployments:1370887552525139968> Deployments - {df[df['Enemy Type'] == 'Terminids']['Enemy Type'].count().sum()}\n" +
                         f"> <:major_order:1356035244033048788> Major Order Deployments - {df[df['Enemy Type'] == 'Terminids']['Major Order'].astype(int).sum()}\n" +
                         f"> <:dss:1356034406430806036> DSS Deployments - {df[df['Enemy Type'] == 'Terminids']['DSS Active'].astype(int).sum()}\n" +
                         f"> <:lastdeployment:1370887542445965564> Last Deployment - {df[df['Enemy Type'] == 'Terminids']['Time'].max() if 'Time' in df.columns else 'No date available'}\n\n" +

                         f"> <:liberation_campaign:1355955855572602962> Liberations - {df[df['Enemy Type'] == 'Terminids'][df['Mission Category'] == 'Liberation']['Mission Category'].count().sum()}\n" +
                         f"> <:defence_campaign:1355955857480876282> Defenses - {df[df['Enemy Type'] == 'Terminids'][df['Mission Category'] == 'Defense']['Mission Category'].count().sum()}\n" +
                         f"> <:invasion_campaign:1355955853588562202> Invasion - {df[df['Enemy Type'] == 'Terminids'][df['Mission Category'] == 'Invasion']['Mission Category'].count().sum()}\n" +
                         f"> <:highprioritycampaign:1370787949372899328> High-Priority - {df[df['Enemy Type'] == 'Terminids'][df['Mission Category'] == 'High-Priority']['Mission Category'].count().sum()}\n\n",
      
    #   f"<a:easyshine1:1349110651829747773> <:hd2bugs:1337190441170370693> Terminid Horde Statistics <:hd2bugs:1337190441170370693> <a:easyshine3:1349110648528699422>\n" +
    #                      f"> <:resistance:1370883421496148068> Kills - {df[df['Enemy Subfaction'] == 'Terminid Horde']['Kills'].sum()}\n" +
    #                      f"> <:helldiverHD:1370887412443648070> Deaths - {df[df['Enemy Subfaction'] == 'Terminid Horde']['Deaths'].sum()}\n" +
    #                      f"> <:highprioritytarget:1370882658019704903> Highest Kills in Mission - {df[df['Enemy Subfaction'] == 'Terminid Horde']['Kills'].max()}\n\n" +

    #                      f"> <:deployments:1370887552525139968> Deployments - {df[df['Enemy Subfaction'] == 'Terminid Horde']['Enemy Type'].count().sum()}\n" +
    #                      f"> <:major_order:1356035244033048788> Major Order Deployments - {df[df['Enemy Subfaction'] == 'Terminid Horde']['Major Order'].astype(int).sum()}\n" +
    #                      f"> <:dss:1356034406430806036> DSS Deployments - {df[df['Enemy Subfaction'] == 'Terminid Horde']['DSS Active'].astype(int).sum()}\n" +
    #                      f"> <:lastdeployment:1370887542445965564> Last Deployment - {df[df['Enemy Subfaction'] == 'Terminid Horde']['Time'].max() if 'Time' in df.columns else 'No date available'}\n\n" +

    #                      f"> <:liberation_campaign:1355955855572602962> Liberations - {df[df['Enemy Subfaction'] == 'Terminid Horde'][df['Mission Category'] == 'Liberation']['Mission Category'].count().sum()}\n" +
    #                      f"> <:defence_campaign:1355955857480876282> Defenses - {df[df['Enemy Subfaction'] == 'Terminid Horde'][df['Mission Category'] == 'Defense']['Mission Category'].count().sum()}\n" +
    #                      f"> <:invasion_campaign:1355955853588562202> Invasion - {df[df['Enemy Subfaction'] == 'Terminid Horde'][df['Mission Category'] == 'Invasion']['Mission Category'].count().sum()}\n" +
    #                      f"> <:highprioritycampaign:1370787949372899328> High-Priority - {df[df['Enemy Subfaction'] == 'Terminid Horde'][df['Mission Category'] == 'High-Priority']['Mission Category'].count().sum()}\n\n",

      "color": 16761088,
      "image": {
        "url": "https://cdn.discordapp.com/attachments/1340508329977446484/1370786767128760420/terminidBanner.png?ex=6820c429&is=681f72a9&hm=3ca1e122e8063426a3dd1963791aca33ba6343a7a946b06287d344ce6c0f93a0&"
      },
      "thumbnail": {
        "url": "https://i.ibb.co/PspGgJkH/Terminids-Icon.png"
      }
    },
    # {
    #   "title": "Predator Strain Campaign Record",
    #   "description":   f"<a:easyshine1:1349110651829747773> <:predatorstrain:1370887431586582622> Predator Strain Statistics <:predatorstrain:1370887431586582622> <a:easyshine3:1349110648528699422>\n" +
    #                      f"> <:resistance:1370883421496148068> Kills - {df[df['Enemy Subfaction'] == 'Predator Strain']['Kills'].sum()}\n" +
    #                      f"> <:helldiverHD:1370887412443648070> Deaths - {df[df['Enemy Subfaction'] == 'Predator Strain']['Deaths'].sum()}\n" +
    #                      f"> <:highprioritytarget:1370882658019704903> Highest Kills in Mission - {df[df['Enemy Subfaction'] == 'Predator Strain']['Kills'].max()}\n\n" +

    #                      f"> <:deployments:1370887552525139968> Deployments - {df[df['Enemy Subfaction'] == 'Predator Strain']['Enemy Type'].count().sum()}\n" +
    #                      f"> <:major_order:1356035244033048788> Major Order Deployments - {df[df['Enemy Subfaction'] == 'Predator Strain']['Major Order'].astype(int).sum()}\n" +
    #                      f"> <:dss:1356034406430806036> DSS Deployments - {df[df['Enemy Subfaction'] == 'Predator Strain']['DSS Active'].astype(int).sum()}\n" +
    #                      f"> <:lastdeployment:1370887542445965564> Last Deployment - {df[df['Enemy Subfaction'] == 'Predator Strain']['Time'].max() if 'Time' in df.columns else 'No date available'}\n\n" +

    #                      f"> <:liberation_campaign:1355955855572602962> Liberations - {df[df['Enemy Subfaction'] == 'Predator Strain'][df['Mission Category'] == 'Liberation']['Mission Category'].count().sum()}\n" +
    #                      f"> <:defence_campaign:1355955857480876282> Defenses - {df[df['Enemy Subfaction'] == 'Predator Strain'][df['Mission Category'] == 'Defense']['Mission Category'].count().sum()}\n" +
    #                      f"> <:invasion_campaign:1355955853588562202> Invasion - {df[df['Enemy Subfaction'] == 'Predator Strain'][df['Mission Category'] == 'Invasion']['Mission Category'].count().sum()}\n" +
    #                      f"> <:highprioritycampaign:1370787949372899328> High-Priority - {df[df['Enemy Subfaction'] == 'Predator Strain'][df['Mission Category'] == 'High-Priority']['Mission Category'].count().sum()}\n\n" +
      
    #   f"<a:easyshine1:1349110651829747773> <:sporeburststrain:1370787947800166420> Spore Burst Strain Statistics <:sporeburststrain:1370787947800166420> <a:easyshine3:1349110648528699422>\n" +
    #                      f"> <:resistance:1370883421496148068> Kills - {df[df['Enemy Subfaction'] == 'Spore Burst Strain']['Kills'].sum()}\n" +
    #                      f"> <:helldiverHD:1370887412443648070> Deaths - {df[df['Enemy Subfaction'] == 'Spore Burst Strain']['Deaths'].sum()}\n" +
    #                      f"> <:highprioritytarget:1370882658019704903> Highest Kills in Mission - {df[df['Enemy Subfaction'] == 'Spore Burst Strain']['Kills'].max()}\n\n" +

    #                      f"> <:deployments:1370887552525139968> Deployments - {df[df['Enemy Subfaction'] == 'Spore Burst Strain']['Enemy Type'].count().sum()}\n" +
    #                      f"> <:major_order:1356035244033048788> Major Order Deployments - {df[df['Enemy Subfaction'] == 'Spore Burst Strain']['Major Order'].astype(int).sum()}\n" +
    #                      f"> <:dss:1356034406430806036> DSS Deployments - {df[df['Enemy Subfaction'] == 'Spore Burst Strain']['DSS Active'].astype(int).sum()}\n" +
    #                      f"> <:lastdeployment:1370887542445965564> Last Deployment - {df[df['Enemy Subfaction'] == 'Spore Burst Strain']['Time'].max() if 'Time' in df.columns else 'No date available'}\n\n" +

    #                      f"> <:liberation_campaign:1355955855572602962> Liberations - {df[df['Enemy Subfaction'] == 'Spore Burst Strain'][df['Mission Category'] == 'Liberation']['Mission Category'].count().sum()}\n" +
    #                      f"> <:defence_campaign:1355955857480876282> Defenses - {df[df['Enemy Subfaction'] == 'Spore Burst Strain'][df['Mission Category'] == 'Defense']['Mission Category'].count().sum()}\n" +
    #                      f"> <:invasion_campaign:1355955853588562202> Invasion - {df[df['Enemy Subfaction'] == 'Spore Burst Strain'][df['Mission Category'] == 'Invasion']['Mission Category'].count().sum()}\n" +
    #                      f"> <:highprioritycampaign:1370787949372899328> High-Priority - {df[df['Enemy Subfaction'] == 'Spore Burst Strain'][df['Mission Category'] == 'High-Priority']['Mission Category'].count().sum()}\n\n",

    #   "color": 16761088,
    #   "image": {
    #     "url": "https://cdn.discordapp.com/attachments/1340508329977446484/1370786767128760420/terminidBanner.png?ex=6820c429&is=681f72a9&hm=3ca1e122e8063426a3dd1963791aca33ba6343a7a946b06287d344ce6c0f93a0&"
    #   },
    #   "thumbnail": {
    #     "url": "https://i.ibb.co/PspGgJkH/Terminids-Icon.png"
    #   }
    # },
    {
      "title": "Automaton Campaign Record",
      "description": "<a:easyshine1:1349110651829747773> <:hd2bots:1337190442449502208> Automaton Front Statistics <:hd2bots:1337190442449502208> <a:easyshine3:1349110648528699422>\n" +
                         f"> <:resistance:1370883421496148068> Kills - {df[df['Enemy Type'] == 'Automatons']['Kills'].sum()}\n" +
                         f"> <:helldiverHD:1370887412443648070> Deaths - {df[df['Enemy Type'] == 'Automatons']['Deaths'].sum()}\n" +
                         f"> <:highprioritytarget:1370882658019704903> Highest Kills in Mission - {df[df['Enemy Type'] == 'Automatons']['Kills'].max()}\n\n" +

                         f"> <:deployments:1370887552525139968> Deployments - {df[df['Enemy Type'] == 'Automatons']['Enemy Type'].count().sum()}\n" +
                         f"> <:major_order:1356035244033048788> Major Order Deployments - {df[df['Enemy Type'] == 'Automatons']['Major Order'].astype(int).sum()}\n" +
                         f"> <:dss:1356034406430806036> DSS Deployments - {df[df['Enemy Type'] == 'Automatons']['DSS Active'].astype(int).sum()}\n" +
                         f"> <:lastdeployment:1370887542445965564> Last Deployment - {df[df['Enemy Type'] == 'Automatons']['Time'].max() if 'Time' in df.columns else 'No date available'}\n\n" +

                         f"> <:liberation_campaign:1355955855572602962> Liberations - {df[df['Enemy Type'] == 'Automatons'][df['Mission Category'] == 'Liberation']['Mission Category'].count().sum()}\n" +
                         f"> <:defence_campaign:1355955857480876282> Defenses - {df[df['Enemy Type'] == 'Automatons'][df['Mission Category'] == 'Defense']['Mission Category'].count().sum()}\n" +
                         f"> <:invasion_campaign:1355955853588562202> Invasion - {df[df['Enemy Type'] == 'Automatons'][df['Mission Category'] == 'Invasion']['Mission Category'].count().sum()}\n" +
                         f"> <:highprioritycampaign:1370787949372899328> High-Priority - {df[df['Enemy Type'] == 'Automatons'][df['Mission Category'] == 'High-Priority']['Mission Category'].count().sum()}\n\n",
      
    #   "<a:easyshine1:1349110651829747773> <:hd2bots:1337190442449502208> Automaton Legion Statistics <:hd2bots:1337190442449502208> <a:easyshine3:1349110648528699422>\n" +
    #   "> Kills - \n" +
    #   "> Deaths - \n" +
    #   "> Highest Kills in Mission - \n" +
    #   "> Deployments - \n" +
    #   "> Major Order Deployments - \n" +
    #   "> DSS Deployments - \n" +
    #   "> Last Deployment - \n\n" +
      
    #   "<a:easyshine1:1349110651829747773> <:jetbrigade:1370887479506374736> Jet Brigade Statistics <:jetbrigade:1370887479506374736> <a:easyshine3:1349110648528699422>\n" +
    #   "> Kills - \n" +
    #   "> Deaths - \n" +
    #   "> Highest Kills in Mission - \n" +
    #   "> Deployments - \n" +
    #   "> Major Order Deployments - \n" +
    #   "> DSS Deployments - \n" +
    #   "> Last Deployment - \n\n" +
      
    #   "<a:easyshine1:1349110651829747773> <:incinerationcorps:1355266318705627365> Incineration Corps Statistics <:incinerationcorps:1355266318705627365> <a:easyshine3:1349110648528699422>\n" +
    #   "> Kills - \n" +
    #   "> Deaths - \n" +
    #   "> Highest Kills in Mission - \n" +
    #   "> Deployments - \n" +
    #   "> Major Order Deployments - \n" +
    #   "> DSS Deployments - \n" +
    #   "> Last Deployment - \n\n" +
      
    #   "<a:easyshine1:1349110651829747773> <:jetbrigade:1370887479506374736> Jet Brigade & Incineration Corps Stats <:incinerationcorps:1355266318705627365> <a:easyshine3:1349110648528699422>\n" +
    #   "> Kills - \n" +
    #   "> Deaths - \n" +
    #   "> Highest Kills in Mission - \n" +
    #   "> Deployments - \n" +
    #   "> Major Order Deployments - \n" +
    #   "> DSS Deployments - \n" +
    #   "> Last Deployment -",

      "color": 16739693,
      "image": {
        "url": "https://cdn.discordapp.com/attachments/1340508329977446484/1370786778465964193/automatonBanner.png?ex=6820c42b&is=681f72ab&hm=63213a37d29cfc25661737c7d20867ebea272fffc9e830116c32ef7df3cf1a24&"
      },
      "thumbnail": {
        "url": "https://i.ibb.co/bgNp2q73/Automatons-Icon.png"
      }
    },
    {
      "title": "Illuminate Campaign Record",
      "description": "<a:easyshine1:1349110651829747773> <:hd2squids:1337190439979319306> Illuminate Cult Statistics <:hd2squids:1337190439979319306> <a:easyshine3:1349110648528699422>\n" +
                         f"> <:resistance:1370883421496148068> Kills - {df[df['Enemy Type'] == 'Illuminate']['Kills'].sum()}\n" +
                         f"> <:helldiverHD:1370887412443648070> Deaths - {df[df['Enemy Type'] == 'Illuminate']['Deaths'].sum()}\n" +
                         f"> <:highprioritytarget:1370882658019704903> Highest Kills in Mission - {df[df['Enemy Type'] == 'Illuminate']['Kills'].max()}\n\n" +

                         f"> <:deployments:1370887552525139968> Deployments - {df[df['Enemy Type'] == 'Illuminate']['Enemy Type'].count().sum()}\n" +
                         f"> <:major_order:1356035244033048788> Major Order Deployments - {df[df['Enemy Type'] == 'Illuminate']['Major Order'].astype(int).sum()}\n" +
                         f"> <:dss:1356034406430806036> DSS Deployments - {df[df['Enemy Type'] == 'Illuminate']['DSS Active'].astype(int).sum()}\n" +
                         f"> <:lastdeployment:1370887542445965564> Last Deployment - {df[df['Enemy Type'] == 'Illuminate']['Time'].max() if 'Time' in df.columns else 'No date available'}\n\n" +

                         f"> <:liberation_campaign:1355955855572602962> Liberations - {df[df['Enemy Type'] == 'Illuminate'][df['Mission Category'] == 'Liberation']['Mission Category'].count().sum()}\n" +
                         f"> <:defence_campaign:1355955857480876282> Defenses - {df[df['Enemy Type'] == 'Illuminate'][df['Mission Category'] == 'Defense']['Mission Category'].count().sum()}\n" +
                         f"> <:invasion_campaign:1355955853588562202> Invasion - {df[df['Enemy Type'] == 'Illuminate'][df['Mission Category'] == 'Invasion']['Mission Category'].count().sum()}\n" +
                         f"> <:highprioritycampaign:1370787949372899328> High-Priority - {df[df['Enemy Type'] == 'Illuminate'][df['Mission Category'] == 'High-Priority']['Mission Category'].count().sum()}\n\n",

      "color": 9003210,
      "image": {
        "url": "https://cdn.discordapp.com/attachments/1340508329977446484/1370786787441774632/illuminateBanner.png?ex=6820c42e&is=681f72ae&hm=bc4d9e9f89bcae58521b9af1558816ecb0c336bee108862725663b87e5bb6079&"
      },
      "thumbnail": {
        "url": "https://i.ibb.co/wr4Nm5HT/Illuminate-Icon.png"
      }
    }
    ],
    "attachments": []
}

# Update the embed title with name and level
embed_data["embeds"][0]["title"] = f"{helldiver_ses}\nHelldiver: {helldiver_name}\nLevel {helldiver_level} | {helldiver_title}"

# Enemy type specific embeds with icons
enemy_icons = {
    "Automatons": {
        "emoji": config['EnemyIcons']['Automatons'],
        "color": int(config['SystemColors']['Automatons']),
        "url": "https://i.ibb.co/bgNp2q73/Automatons-Icon.png"
    },
    "Terminids": {
        "emoji": config['EnemyIcons']['Terminids'],
        "color": int(config['SystemColors']['Terminids']),
        "url": "https://i.ibb.co/PspGgJkH/Terminids-Icon.png"
    },
    "Illuminate": {
        "emoji": config['EnemyIcons']['Illuminate'],
        "color": int(config['SystemColors']['Illuminate']),
        "url": "https://i.ibb.co/wr4Nm5HT/Illuminate-Icon.png"
    }
}

enemy_banners = {
    "Automatons": {
        "url": "https://cdn.discordapp.com/attachments/1340508329977446484/1370786778465964193/automatonBanner.png?ex=6820c42b&is=681f72ab&hm=63213a37d29cfc25661737c7d20867ebea272fffc9e830116c32ef7df3cf1a24&"
    },
    "Terminids": {
        "url": "https://cdn.discordapp.com/attachments/1340508329977446484/1370786767128760420/terminidBanner.png?ex=6820c429&is=681f72a9&hm=3ca1e122e8063426a3dd1963791aca33ba6343a7a946b06287d344ce6c0f93a0&"
    },
    "Illuminate": {
        "url": "https://cdn.discordapp.com/attachments/1340508329977446484/1370786787441774632/illuminateBanner.png?ex=6820c42e&is=681f72ae&hm=bc4d9e9f89bcae58521b9af1558816ecb0c336bee108862725663b87e5bb6079&"
    }
}

# Group data by enemy type (faction)
faction_stats = {}
for enemy_type in enemy_types:
    faction_data = df[df["Enemy Type"] == enemy_type]
    if not faction_data.empty:
        faction_stats[enemy_type] = {
            "total_kills": faction_data["Kills"].sum(),
            "total_deaths": faction_data["Deaths"].sum(),
            "total_deployments": len(faction_data),
            "major_orders": faction_data["Major Order"].astype(int).sum(),
            "last_deployment": faction_data["Time"].max() if "Time" in df.columns else "No date available",
            "planets": faction_data["Planet"].unique().tolist()
        }

# Add enemy-specific embeds
# for enemy_type, stats in faction_stats.items():
#     faction_description = f"{enemy_icons.get(enemy_type, {'emoji': ''})['emoji']} **{enemy_type} Front Statistics**\n" + \
#         f"> Deployments - {stats['total_deployments']}\n" + \
#         f"> Major Order Deployments - {stats['major_orders']}\n" + \
#         f"> Kills - {stats['total_kills']}\n" + \
#         f"> Deaths - {stats['total_deaths']}\n" + \
#         f"> Last Deployment - {stats['last_deployment']}\n\n"

#     embed_data["embeds"].append({
#         "title": f"{enemy_type} Campaign Record",
#         "description": faction_description,
#         "color": enemy_icons.get(enemy_type, {"color": 7257043})["color"],
#         "thumbnail": {"url": enemy_icons.get(enemy_type, {"url": ""})["url"]},
#         "image": {"url": enemy_banners.get(enemy_type, {"url": ""})["url"]}
#     })

# embed_data = {
#     "content": None,
#     "embeds": [
#         {
#             "title": "Terminids Campaign Record",
#             "description": f"\n\n<a:easyshine1:1349110651829747773> <:hd2bugs:1337190441170370693> Terminid Front Statistics <:hd2bugs:1337190441170370693> <a:easyshine3:1349110648528699422>\n" + 
#                         f"> Kills - {df[df["Enemy Type"] == "Terminids"]['Kills'].sum()}\n" +
#                         f"> Deaths - {df[df["Enemy Type"] == "Terminids"]['Deaths'].sum()}\n" +
#                         f"> Highest Kills in Mission - {df[df["Enemy Type"] == "Terminids"]['Kills'].max()}\n" +
#                         f"> Deployments - {len(df["Enemy Type" == "Terminids"])}\n" +
#                         f"> Major Order Deployments - {df[df["Enemy Type"] == "Terminids"]['Major Order'].astype(int).sum()}\n" +
#                         f"> DSS Deployments - {df[df["Enemy Type"] == "Terminids"]['DSS Active'].astype(int).sum()}\n" +

#                         f"\n<a:easyshine1:1349110651829747773>  <a:easyskullgold:1232018045791375360> Performance Statistics <a:easyskullgold:1232018045791375360> <a:easyshine3:1349110648528699422>\n" +                      
#                         f"> Rating - {Rating} | {int(Rating_Percentage)}%\n" +

#                         f"\n<a:easyshine1:1349110651829747773>  <:goldstar:1337818552094163034> Favourites <:goldstar:1337818552094163034> <a:easyshine3:1349110648528699422>\n" +     
#                         f"> Mission - {df['Mission Type'].mode()[0]} {MISSION_ICONS.get(df['Mission Type'].mode()[0], '')} (x{MissionCount})\n" +
#                         f"> Campaign - {df['Mission Category'].mode()[0]} {CAMPAIGN_ICONS.get(df['Mission Category'].mode()[0], '')} (x{CampaignCount})\n" +
#                         f"> Faction - {df['Enemy Type'].mode()[0]} {ENEMY_ICONS.get(df['Enemy Type'].mode()[0], '')} (x{FactionCount})\n" +
#                         f"> Difficulty - {df['Difficulty'].mode()[0]} {DIFFICULTY_ICONS.get(df['Difficulty'].mode()[0], '')} (x{DifficultyCount})\n" +
#                         f"> Planet - {df['Planet'].mode()[0]} {PLANET_ICONS.get(df['Planet'].mode()[0], '')} (x{PlanetCount})\n" +
#                         f"> Sector - {df['Sector'].mode()[0]} (x{SectorCount})\n",
#             "color": 7257043,
#             "author": {"name": "SEAF Battle Record"},
#             "footer": {"text": config['Discord']['UID'],"icon_url": "https://cdn.discordapp.com/attachments/1340508329977446484/1356025859319926784/5cwgI15.png?ex=67eb10fe&is=67e9bf7e&hm=ab6326a9da1e76125238bf3668acac8ad1e43b24947fc6d878d7b94c8a60ab28&"},
#             "image": {"url": f"{BIOME_BANNERS.get(df['Planet'].mode()[0], '')}"},
#             "thumbnail": {"url": "https://i.ibb.co/5g2b9NXb/Super-Earth-Icon.png"}
#         }
#     ],
#     "attachments": []
# }

# Send data to Discord
webhook_urls = WEBHOOK_URLS['TEST'] if DEBUG else WEBHOOK_URLS['PROD']
for webhook_url in webhook_urls:
    response = requests.post(webhook_url, json=embed_data)
    print("Data sent successfully." if response.status_code == 204 else f"Failed to send data. Status: {response.status_code}")
