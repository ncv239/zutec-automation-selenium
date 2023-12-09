# zutec-automation-selenium
Script to automate check sheet generation in [Zutec](https://www.zutec.com/) document management platform for routine tasks of shift engineers. Uses selenium.

This particular script will
- login to Zutec using credentials provided in file `credentials.py`
- read data from excel table (see sample `test_data.xlsx`)
- generate a proof-drill checksheetfor each data-row in excel sheet

#Usage
- Download the chrome driver for selenium in the script directory
- Install python libraries from file `req.txt`
- Update excel data and credentials. Make sure to put dates as of type `str`/`TEXT` in excel workbook.
- Run the file `add_proof_drill.py`
- Selenium will open a browser, maximise the window to ensure css selectors work properly

