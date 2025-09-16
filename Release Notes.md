**3.2.1**
- Added error handling for proginfo write failure
- Added prompt for sheet name, user, and analytics output location if none is provided
- Added support to handle disconnected wifi while the program is running. If the program fails to write data to Google it will now output to a spreadsheet instead of crashing and not saving anything.