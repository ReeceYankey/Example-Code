# Overview
CanvasWebScraper.py launches a chrome emulator where you just have to log in to canvas and it will automatically grab existing grades. It then stores these into csv files, updates the Galipatia Academic Success Database.xlsx accordingly, and saves it into a file called updated.xlsx. If you run UpdateFromCSV.py directly, it uses the existing csv files from a previous run to update the spreadsheet.

# Setup
1. pip install selenium
2. pip install pandas
3. pip install beautifulsoup4
4. download the [chrome driver](https://chromedriver.chromium.org/downloads) for your version of chrome
5. change line 8 in CanvasWebScraper.py to the path of the chromedriver.exe
