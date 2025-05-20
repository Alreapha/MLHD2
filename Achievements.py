import pandas as pd
import configparser
import requests
import json

# Read configuration from config.config
config = configparser.ConfigParser()
config.read('config.config')

#Constants
DEBUG = config.getboolean('DEBUGGING', 'DEBUG', fallback=False)

# Read the Excel file
try:
    df = pd.read_excel('mission_log_test.xlsx') if DEBUG else pd.read_excel('mission_log.xlsx')
except FileNotFoundError:
    print("Error: Excel file not found. Please ensure the file exists in the correct location.")
    exit(1)


highest_streak = 0
profile_picture = ""
with open('streak_data.json', 'r') as f:
    streak_data = json.load(f)
    # Use "Helldiver" as the key
    highest_streak = streak_data.get("Helldiver", {}).get("highest_streak", 0)
    profile_picture = streak_data.get("Helldiver", {}).get("profile_picture_name", "")

#get total kills
total_kills = df['Kills'].sum()


# Aggregate data for each advancement query

# get total missions
total_missions = len(df)

#get total mission with major order active
total_missions_major_order = df[df['Major Order'] == 1].shape[0]

#get total mission with DSS active
total_missions_dss = df[df['DSS Active'] == 1].shape[0]

#get total terminid missions
total_terminid_missions = df[df['Enemy Type'] == 'Terminids'].shape[0]

#get total automaton missions
total_automaton_missions = df[df['Enemy Type'] == 'Automatons'].shape[0]

#get total illuminate missions
total_illuminate_missions = df[df['Enemy Type'] == 'Illuminates'].shape[0]

# get total terminid kills
total_terminid_kills = df[df['Enemy Type'] == 'Terminids']['Kills'].sum()

# get total automaton kills
total_automaton_kills = df[df['Enemy Type'] == 'Automatons']['Kills'].sum()

# get total illuminate kills
total_illuminate_kills = df[df['Enemy Type'] == 'Illuminates']['Kills'].sum()

# get if at least one mission was played on Malevelon Creek
malevelon_creek = df[df['Planet'] == 'Malevelon Creek'].shape[0] > 0

# get if at least on mission was rated Disgracful Conduct
disgraceful_conduct = df[df['Rating'] == 'Disgraceful Conduct'].shape[0] > 0

#get if at least one mission was played on Super Earth
super_earth = df[df['Planet'] == 'Super Earth'].shape[0] > 0

# get at least one mission was played on the Cyberstan
cyberstan = df[df['Planet'] == 'Cyberstan'].shape[0] > 0

# get if highest_streak is 30 or more
streak_30 = highest_streak >= 30


#assign bool values to variables
if total_missions >= 1000:
    CmdFavourite = True
else:
    CmdFavourite = False

if total_missions_major_order >= total_missions / 2:
    ReliableDiver = True
else:
    ReliableDiver = False

if total_missions_dss >= total_missions / 2:
    DSSDiver = True
else:
    DSSDiver = False

if total_terminid_missions >= 250:
    OutbreakPerfected = True
else:
    OutbreakPerfected = False

if total_automaton_missions >= 250:
    AutomatonPerfected = True
else:
    AutomatonPerfected = False

if total_illuminate_missions >= 250:
    IlluminatePerfected = True
else:
    IlluminatePerfected = False

if total_terminid_kills >= 100000:
    TerminidHunter = True
else:
    TerminidHunter = False

if total_automaton_kills >= 100000:
    AutomatonHunter = True
else:
    AutomatonHunter = False

if total_illuminate_kills >= 100000:
    IlluminateHunter = True
else:
    IlluminateHunter = False

if malevelon_creek:
    MalevelonCreek = True
else:
    MalevelonCreek = False

if disgraceful_conduct:
    DisgracefulConduct = True
else:
    DisgracefulConduct = False

if super_earth:
    SuperEarth = True
else:
    SuperEarth = False

if cyberstan:
    Cyberstan = True
else:
    Cyberstan = False

if streak_30:
    Streak30 = True
else:
    Streak30 = False

# Create a dictionary to store the achievements
achievements = {
    "CmdFavourite": CmdFavourite,
    "ReliableDiver": ReliableDiver,
    "DSSDiver": DSSDiver,
    "OutbreakPerfected": OutbreakPerfected,
    "AutomatonPerfected": AutomatonPerfected,
    "IlluminatePerfected": IlluminatePerfected,
    "TerminidHunter": TerminidHunter,
    "AutomatonHunter": AutomatonHunter,
    "IlluminateHunter": IlluminateHunter,
    "MalevelonCreek": MalevelonCreek,
    "DisgracefulConduct": DisgracefulConduct,
    "SuperEarth": SuperEarth,
    "Cyberstan": Cyberstan,
    "Streak30": Streak30
}

# Load Webhook URL from config
# Discord webhook configuration
WEBHOOK_URLS = {
    'PROD': config['Webhooks']['BAT'].split(','),
    'TEST': config['Webhooks']['TEST'].split(',')
}
ACTIVE_WEBHOOK = WEBHOOK_URLS['TEST'] if DEBUG else WEBHOOK_URLS['PROD']

#get title icons
TITLE_ICONS = {
    "CADET": config['TitleIcons']['CADET'],
    "SPACE CADET": config['TitleIcons']['SPACE CADET'], 
    "SERGEANT": config['TitleIcons']['SERGEANT'],
    "MASTER SERGEANT": config['TitleIcons']['MASTER SERGEANT'],
    "CHIEF": config['TitleIcons']['CHIEF'],
    "SPACE CHIEF PRIME": config['TitleIcons']['SPACE CHIEF PRIME'],
    "DEATH CAPTAIN": config['TitleIcons']['DEATH CAPTAIN'],
    "MARSHAL": config['TitleIcons']['MARSHAL'],
    "STAR MARSHAL": config['TitleIcons']['STAR MARSHAL'],
    "ADMIRAL": config['TitleIcons']['ADMIRAL'], 
    "SKULL ADMIRAL": config['TitleIcons']['SKULL ADMIRAL'],
    "FLEET ADMIRAL": config['TitleIcons']['FLEET ADMIRAL'],
    "ADMIRABLE ADMIRAL": config['TitleIcons']['ADMIRABLE ADMIRAL'],
    "COMMANDER": config['TitleIcons']['COMMANDER'],
    "GALACTIC COMMANDER": config['TitleIcons']['GALACTIC COMMANDER'],
    "HELL COMMANDER": config['TitleIcons']['HELL COMMANDER'],
    "GENERAL": config['TitleIcons']['GENERAL'],
    "5-STAR GENERAL": config['TitleIcons']['5-STAR GENERAL'],
    "10-STAR GENERAL": config['TitleIcons']['10-STAR GENERAL'],
    "PRIVATE": config['TitleIcons']['PRIVATE'],
    "SUPER PRIVATE": config['TitleIcons']['SUPER PRIVATE'],
    "SUPER CITIZEN": config['TitleIcons']['SUPER CITIZEN'],
    "VIPER COMMANDO": config['TitleIcons']['VIPER COMMANDO'],
    "FIRE SAFETY OFFICER": config['TitleIcons']['FIRE SAFETY OFFICER'],
    "EXPERT EXTERMINATOR": config['TitleIcons']['EXPERT EXTERMINATOR'],
    "FREE OF THOUGHT": config['TitleIcons']['FREE OF THOUGHT'],
    "SUPER PEDESTRIAN": config['TitleIcons']['SUPER PEDESTRIAN'],
    "ASSAULT INFANTRY": config['TitleIcons']['ASSAULT INFANTRY'],
    "SERVANT OF FREEDOM": config['TitleIcons']['SERVANT OF FREEDOM'],
    "SUPER SHERIFF": config['TitleIcons']['SUPER SHERIFF'],
    "DECORATED HERO": config['TitleIcons']['DECORATED HERO']
}

#generate message for advancements
if achievements["CmdFavourite"]:
    CmdFavourite_message = "Log 1000 Missions"
else:
    CmdFavourite_message = "HINT: You have the strength and the courage... to be free"

if achievements["CmdFavourite"]:
    CmdFavourite_title = "<a:EasyAwardBaftaMP2025:1363545915352289371> **HIGH COMMAND'S FAVOURITE**"
else:
    CmdFavourite_title = "<:achievement_style_3:1374174049726632067> **~~HIGH COMMAND'S FAVOURITE~~**"

if achievements["ReliableDiver"]:
    ReliableDiver_message = "More than 50% of your logged missions are involved in a Major Order :major_order: "
else:
    ReliableDiver_message = "HINT: You're one to obey orders"

if achievements["ReliableDiver"]:
    ReliableDiver_title = "<a:EasyAwardBaftaMusic2025:1359268029850058974> **RELIABLE DIVER**"
else:
    ReliableDiver_title = "<:achievement_style_1:1374174053254041640> **~~RELIABLE DIVER~~**"

if achievements["DSSDiver"]:
    DSSDiver_message = "More than 50% of your logged Missions are involved with the Democracy Space Station :dss: "
else:
    DSSDiver_message = "HINT: You like a good bit of support"

if achievements["DSSDiver"]:
    DSSDiver_title = "<a:EasyAwardBaftaMusic2025:1359268029850058974> **I <3 DSS**"
else:
    DSSDiver_title = "<:achievement_style_1:1374174053254041640> **~~I <3 DSS~~**"

if achievements["OutbreakPerfected"]:
    OutbreakPerfected_message = "Log 250 Terminid Missions"
else:
    OutbreakPerfected_message = "HINT: You're rather familiar with E-710"

if achievements["OutbreakPerfected"]:
    OutbreakPerfected_title = "<a:EasyMedal:1233854253077102653> **OUTBREAK PERFECTED**"
else:
    OutbreakPerfected_title = "<:achievement_style_2:1374174051551154267> **~~OUTBREAK PERFECTED~~**"

if achievements["AutomatonPerfected"]:
    AutomatonPerfected_message = "Log 250 Automaton Missions"
else:
    AutomatonPerfected_message = "HINT: You're rather familiar with losing access to your Stratagems"

if achievements["AutomatonPerfected"]:
    AutomatonPerfected_title = "<a:EasyMedal:1233854253077102653> **INCURSION DEVASTATED**"
else:
    AutomatonPerfected_title = "<:achievement_style_2:1374174051551154267> **~~INCURSION DEVASTATED~~**"

if achievements["IlluminatePerfected"]:
    IlluminatePerfected_message = "Log 250 Illuminates Missions"
else:
    IlluminatePerfected_message = "HINT: You're rather familiar with their autocratic intentions"

if achievements["IlluminatePerfected"]:
    IlluminatePerfected_title = "<a:EasyMedal:1233854253077102653> **INVASION ABOLISHED**"
else:
    IlluminatePerfected_title = "<:achievement_style_2:1374174051551154267> **~~INVASION ABOLISHED~~**"

if achievements["TerminidHunter"]:
    TerminidHunter_message = "Log 100,000 Kills against the Terminids"
else:
    TerminidHunter_message = "HINT: You douse yourself in E-710"

if achievements["TerminidHunter"]:
    TerminidHunter_title = "<a:EasyAwardBaftaMP2025:1363545915352289371> **BUG STOMPER**"
else:
    TerminidHunter_title = "<:achievement_style_3:1374174049726632067> **~~BUG STOMPER~~**"

if achievements["AutomatonHunter"]:
    AutomatonHunter_message = "Log 100,000 Kills against the Automatons"
else:
    AutomatonHunter_message = "HINT: You make things out of scrap metal in your spare time"

if achievements["AutomatonHunter"]:
    AutomatonHunter_title = "<a:EasyAwardBaftaMP2025:1363545915352289371> **CLANKER SCRAPPER**"
else:
    AutomatonHunter_title = "<:achievement_style_3:1374174049726632067> **~~CLANKER SCRAPPER~~**"

if achievements["IlluminateHunter"]:
    IlluminateHunter_message = "Log 100,000 Kills against the Illuminates"
else:
    IlluminateHunter_message = "HINT: You single handedly make an effort of wiping them out of the Second Galactic War"

if achievements["IlluminateHunter"]:
    IlluminateHunter_title = "<a:EasyAwardBaftaMP2025:1363545915352289371> **SQUID SEVERER**"
else:
    IlluminateHunter_title = "<:achievement_style_3:1374174049726632067> **~~SQUID SEVERER~~**"

if achievements["MalevelonCreek"]:
    MalevelonCreek_message = "Serve on Malevelon Creek"
else:
    MalevelonCreek_message = "HINT: You remember..."

if achievements["MalevelonCreek"]:
    MalevelonCreek_title = "<a:EasyAwardBaftaMusic2025:1359268029850058974> **NEVER FORGET**"
else:
    MalevelonCreek_title = "<:achievement_style_1:1374174053254041640> **~~NEVER FORGET~~**"

if achievements["DisgracefulConduct"]:
    DisgracefulConduct_message = "Get a Performance Rating of Disgraceful Conduct on a Mission"
else:
    DisgracefulConduct_message = "HINT: You... why?"

if achievements["DisgracefulConduct"]:
    DisgracefulConduct_title = "<a:EasyMedal:1233854253077102653> **you got this on purpose...**"
else:
    DisgracefulConduct_title = "<:achievement_style_2:1374174051551154267> **~~you got this on purpose...~~"

if achievements["SuperEarth"]:
    SuperEarth_message = "Serve on Super Earth"
else:
    SuperEarth_message = "HINT: You feel very welcome"

if achievements["SuperEarth"]:
    SuperEarth_title = "<a:EasyAwardBaftaMusic2025:1359268029850058974> **HOME SUPER HOME**"
else:
    SuperEarth_title = "<:achievement_style_1:1374174053254041640> **~~HOME SUPER HOME~~**"

if achievements["Cyberstan"]:
    Cyberstan_message = "Serve on an Enemy Homeworld"
else:
    Cyberstan_message = "HINT: You don't feel very welcome... like they have a choice"

if achievements["Cyberstan"]:
    Cyberstan_title = "<a:EasyAwardBaftaMusic2025:1359268029850058974> **ON THE ENEMY'S DOORSTEP**"
else:
    Cyberstan_title = "<:achievement_style_1:1374174053254041640> **~~ON THE ENEMY'S DOORSTEP~~**"

if achievements["Streak30"]:
    Streak30_message = "Reach a Streak of 30"
else:
    Streak30_message = "HINT: You'll need to take some annual leave after this... seriously... Democracy Applauds You!"

if achievements["Streak30"]:
    Streak30_title = "<a:EasyMedal:1233854253077102653> **INFLAMMABLE**"
else:
    Streak30_title = "<:achievement_style_2:1374174051551154267> **~~INFLAMMABLE~~**"

# generate embed message

helldiver_level = df['Level'].mode()[0]
helldiver_title = df['Title'].mode()[0]
latest_note = df['Note'].mode()[0] if not df['Note'].isnull().all() else "No notes available"

# Discord webhook configuration
WEBHOOK_URLS = {
    'PROD': config['Webhooks']['BAT'].split(','),
    'TEST': config['Webhooks']['TEST'].split(',')
}
ACTIVE_WEBHOOK = WEBHOOK_URLS['TEST'] if DEBUG else WEBHOOK_URLS['PROD']

# Create embed data
embed_data = {
    "content": None,
    "embeds": [
        {
            "title": "{helldiver_ses}\nHelldiver: {helldiver_name}",  # Empty title, will be set below
            "description": f"**Level {helldiver_level} | {helldiver_title} {TITLE_ICONS.get(df['Title'].mode()[0], '')}**\n\n\"{latest_note}\"\n\n<a:easyshine1:1349110651829747773> <a:easymedal:1233854253077102653> Achievements <a:easymedal:1233854253077102653> <a:easyshine3:1349110648528699422>\n" + 
                        f"> {CmdFavourite_title}\n" +
                        f"> *{CmdFavourite_message}*\n" +
                        f"> \n" +
                        f"> {ReliableDiver_title}\n" +
                        f"> *{ReliableDiver_message}*\n" +
                        f"> \n" +
                        f"> {DSSDiver_title}\n" +
                        f"> *{DSSDiver_message}*\n" +
                        f"> \n" +
                        f"> {OutbreakPerfected_title}\n" +
                        f"> *{OutbreakPerfected_message}*\n" +
                        f"> \n" +
                        f"> {AutomatonPerfected_title}\n" +
                        f"> *{AutomatonPerfected_message}*\n" +
                        f"> \n" +
                        f"> {IlluminatePerfected_title}\n" +
                        f"> *{IlluminatePerfected_message}*\n" +
                        f"> \n" +
                        f"> {TerminidHunter_title}\n" +
                        f"> *{TerminidHunter_message}*\n" +
                        f"> \n" +
                        f"> {AutomatonHunter_title}\n" +
                        f"> *{AutomatonHunter_message}*\n" +
                        f"> \n" +
                        f"> {IlluminateHunter_title}\n" +
                        f"> *{IlluminateHunter_message}*\n" +
                        f"> \n" +
                        f"> {MalevelonCreek_title}\n" +
                        f"> *{MalevelonCreek_message}*\n" +
                        f"> \n" +
                        f"> {DisgracefulConduct_title}\n" +
                        f"> *{DisgracefulConduct_message}*\n" +
                        f"> \n" +
                        f"> {SuperEarth_title}\n" +
                        f"> *{SuperEarth_message}*\n" +
                        f"> \n" +
                        f"> {Cyberstan_title}\n" +
                        f"> *{Cyberstan_message}*\n" +
                        f"> \n" +
                        f"> {Streak30_title}\n" +
                        f"> *{Streak30_message}*\n\n",
                        
            "color": 7257043,
            "author": {"name": "SEAF Battle Record"},
            "footer": {"text": config['Discord']['UID'],"icon_url": "https://cdn.discordapp.com/attachments/1340508329977446484/1356025859319926784/5cwgI15.png?ex=67eb10fe&is=67e9bf7e&hm=ab6326a9da1e76125238bf3668acac8ad1e43b24947fc6d878d7b94c8a60ab28&"},
            "image": {"url": f"https://cdn.discordapp.com/attachments/1340508329977446484/1374164186850000957/helldiversBanner.png?ex=682d0da0&is=682bbc20&hm=c80377ccc47f3e1b08661f1f48fadc8f8c171dbb9158087a9a96613a0ad366fb&"},
            "thumbnail": {"url": f"{profile_picture}"}
        }
    ],
    "attachments": []
}

# Send the embed message to Discord
for webhook_url in ACTIVE_WEBHOOK:
    try:
        response = requests.post(webhook_url, json=embed_data)
        if response.status_code == 204:
            print("Message sent successfully.")
        else:
            print(f"Failed to send message. Status code: {response.status_code}")
    except requests.exceptions.RequestException as e:
        print(f"Error sending message: {e}")