# Daily-Logic-Reports

This program runs through Google App Scripts. First, it takes the downloaded Logic Reports, combines them into one google spreadsheet, copies the template spreadsheet, and inserts the data from each sheet into the template.

## Usage
This program was designed to autommate some tasks I performed at my old job. Each day we scrape 15 Activity Reports [(program link)](https://github.com/watson-clara/Logics-Database-Report-Scrapper/blob/2eb5ddb20dde461fc19deb508a700a29f0d55833/activity_report.py) and 15 SO Reports with [(program link)](https://github.com/watson-clara/Logics-Database-Report-Scrapper/blob/2eb5ddb20dde461fc19deb508a700a29f0d55833/activity_report.py). Previously the company was doing this manually, but I figured out a way to completely automate this by connecting Google Drive to my Desktop. After all the reports are loaded onto my google drive I use this program, Daily-Logic-Reports, through Google App Scripts to put the data into the correct spreadsheet.
