# MLHD2
 Helldivers 2 Mission Log Manager

Downloads
Helldivers 2 Operation Logger
https://github.com/Alreapha/MLHD2/tree/main

Python 3.10.6
https://www.python.org/downloads/release/python-3106/

pip 22.2.1 (Any Version Should Work)
https://pypi.org/project/pip/22.2.1/

Dependencies
Open Terminal
cd into HD2 Operation Logger directory (where the location of your download is for the program) Alternatively you can right click your file explorer window and select "Open in Terminal"
Run pip install -r .\requirements.txt
Double Click main.py to run HD2 Operation Logger

config.config
DISCORD_CLIENT_ID
You shouldn't ever have to edit this client ID, however it's here in case you do need it.

Excel Location
PROD is the main name of your spreadsheet where you can see all of your mission logs in one place on your device
TEST is not important and you shouldn't need to touch this unless you're exploring the insides of the program yourself or are guided as a tester

Webhooks
PROD is the main webhook that will link to our server, if you wish you can create your own webhook link and have it also upload to your own server by:
https://discord.com/our-link,https://discord.com/your-link
TEST is not important and you shouldn't need to touch this unless you're exploring the insides of the program yourself or are guided as a tester

EnemyIcons, DifficultyIcons, Stars
These will only be important if you add your own Webhook link as the emojis in the embed may not function correctly in your own server

SystemColors
You do not need to touch these at all, these are references only

If you know how, you can make a .bat file for easier use instead of running the main.py, and you can then turn this into a desktop shortcut to treat it like an .exe
Due to current restrictions, we're unable to make this an exe ourselves at this time


