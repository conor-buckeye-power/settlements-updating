# Documentation

### Files that need to be downloaded prior to running (list in-progress)
- Load without Losses Report

##### As of 7/9/2018: The program will only ask for Month and Year one time, at the beginning of the program.

### Get MISO LMPs
- Select a month and year for the LMP retriever to run on
- The program will iterate from 1 to 31 checking to see if cell **A4** is empty or not
- If cell **A4** *is* empty then the program will begin pulling LMPs for that date

### FTR LMP Updater
- Select a month and year for the LMP retriever to run on
- The program will open FTR LMP files from Kevin's folder and add them to the FTR LMPs workbook
- If cell **A201** is empty then the program will start on that sheet adding LMPs until the program runs out of sheets to use
- *Note*: This program uses data from L:\Resource Scheduling\Kevin\LMP Data\DataMiner

### Load without Losses
- The load without losses report will need to be manually updated and saved at this location L:\Resource Scheduling\Conor\Data\load_with_and_without_losses
- The program will ask for a month and year to update
- From here that specified months data will be pulled from the report and the Load workbooks will be opened
- The new data will replace the current existing data from cells **B38:Y68**
