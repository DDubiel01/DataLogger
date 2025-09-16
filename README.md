## DataLogger
Your one stop shop for logging snipes

## Purpose
This was created to make logging and tracking snipes easier. The code allows loggers to input pairs of names, makes sure they are valid, and writes them to a Google spreadsheet. Why not just log them in a spreadsheet directly? SniperCode prevents you from making errors by only allowing whitelisted names and by automatically logging dates. It's also easier to launch and you don't have to worry about keeping track of spreadsheet versions or ensuring loggers are keeping the same format.

## Features
- **White List Names**
    - DataLogger uses a list of acceptable names stored on Sheets to ensure loggers are keeping consistent format and mitigating mistakes
- **Google Sheets Access**
    - Sheets is used as a database for storing snipes. Users don't have to worry about accessing and viewing databases and can easily get to the raw data to correct mistakes
- **Analytics**
    - Stat breakdowns for the group as a whole and for individual players. Users can select which stats they want generated and view the results in a spreadsheet
- **Multiple Themes**
    - DataLogger can run in any Themed Tkinter theme. The default is set to Aqua

## Installation
1. Install the requirements.txt file  
`pip install -r path/to/requirements.txt`
1. Drop the required OAuth credentials into the gui/config folder with the name client_secret.json
1. Done!

1. hello from vscode