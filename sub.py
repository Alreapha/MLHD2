import pandas as pd
import configparser
import requests
import json

#Constants
DEBUG = False

# Read configuration from config.config
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
rating_mapping = {"Outstanding Patriotism": 5, "Superior Valour": 4, "Honourable Duty":3, "Unremarkable Performance":2, "Dissapointing Service":1, "Disgraceful Conduct":0}
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
            "thumbnail": {"url": "https://i.ibb.co/5g2b9NXb/Super-Earth-Icon.png"}
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

# Planet type specific embeds with icons
planet_icons = {
    "Super Earth": {
        "emoji": config['PlanetIcons']['Human Homeworld']
    },
    "Cyberstan": {
        "emoji": config['PlanetIcons']['Automaton Homeworld']
    },
    "Malevelon Creek": {
        "emoji": config['PlanetIcons']['Malevelon Creek']
    },
    "Calypso": {
        "emoji": config['PlanetIcons']['Calypso']
    },
    "Diaspora X": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Enuliale": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Epsilon Phoencis VI": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Gemstone Bluffs": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Nabatea Secundus": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Navi VII": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Azur Secundus": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Erson Sands": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Nivel 43": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Zagon Prime": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Hellmire": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Omicron": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Oshaune": {
        "emoji": config['PlanetIcons']['Gloom']
    },
    "Fori Prime": {
        "emoji": config['PlanetIcons']['Gloom']
    }
}

# Group planets by enemy type
enemy_planets = {}
for planet in planets:
    planet_data = df[df["Planet"] == planet]
    if not planet_data.empty:
        enemy_type = planet_data["Enemy Type"].iloc[0]
        if enemy_type not in enemy_planets:
            enemy_planets[enemy_type] = []
        enemy_planets[enemy_type].append((planet, planet_data))

# Add enemy-specific embeds
for enemy_type, planet_list in enemy_planets.items():
    planets_description = ""
    for planet, planet_data in planet_list:
        last_date = planet_data["Time"].max() if "Time" in df.columns else "No date available"
        planets_description += f"{enemy_icons.get(enemy_type, {'emoji': ''})['emoji']} **{planet}**\n" + \
               f"> Deployments - {len(planet_data)}\n" + \
               f"> Major Order Deployments - {planet_data['Major Order'].astype(int).sum()}\n" + \
               f"> Kills - {planet_data['Kills'].sum()}\n" + \
               f"> Deaths - {planet_data['Deaths'].sum()}\n" + \
               f"> Last Deployment - {last_date}\n\n"

    embed_data["embeds"].append({
        "title": f"{enemy_type} Front",
        "description": planets_description,
        "color": enemy_icons.get(enemy_type, {"color": 7257043})["color"],
        "thumbnail": {"url": enemy_icons.get(enemy_type, {"url": ""})["url"]}
    })

# Send data to Discord
webhook_urls = WEBHOOK_URLS['TEST'] if DEBUG else WEBHOOK_URLS['PROD']
for webhook_url in webhook_urls:
    response = requests.post(webhook_url, json=embed_data)
    print("Data sent successfully." if response.status_code == 204 else f"Failed to send data. Status: {response.status_code}")
