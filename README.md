# spreadsheet-updater

Scrapes data on stocks and funds and fills it into a Google Spreadsheet.

The user is able to change which quotes/funds are targeted by editing a "config" sheet within their spreadsheet. This makes it very easy for them to add/remove quotes/funds that they want data for.

Web scraping was done using Selenium. Need to download Chromedriver and move to "/usr/local/bin". Chromedriver version needs to match client's Google Chrome version.