# NeoSpecter
A Python script that inputs Excel data into the NeoSerra CRM.

## Libraries Used
1. Openpyxl to drive Excel.
2. Selenium webdriver to fill-in web forms.
3. Pandas may be used if Openpyxl is unable to perform as required.
4. tkinter for GUI.
5. Script bundled into .exe for easy use on Windows OS.

## User Experience
1. User opens NeoSpector.exe.
2. GUI loads. Shows user a _folder picker_ widget.
3. User clicks on folder picker and selects folder containing excel files with client award data.
4. User clicks SELECT FOLDER BUTTON.
5. User clicks LOAD EXCEL FILES BUTTON on GUI. Progress bar appears.
6. Excel files are loaded. Success or failure status is displayed in pop-up.
7. If successful, User clicks RUN SPECTOR BUTTON to activate Selenium.
8. A warning message shown telling the user to not use their mouse or keyboard while Spector is running.
9. User clicks OK BUTTON. Pop-up closes.
10. Progress bar is shown.
11. Selenium activates. Web browser is opened and navigates to NEO SERRA CRM LOGIN.
12. LOG IN credintials are entered.
13. Selenium navigates to AWARDS PAGE.
14. clicks new award.
15. Copies data from the Excel sheets and enters it into the Neo Serra Webforms. Typing simulation is a nice add-on.
16. Webform is saved.
17. Selenium repeats steps 13-16 until all of the excel data has been entered.
18. Spector shows dialogue box showing that the operation is successful.
19. User is returned to home screen.

## Process list
### With PANDAS
1. Client excel data and client awards data is collected into PANDAS. THIS IS ALL PUBLIC INFORMATION. 
2. PANDAS array is used as the data source for the webdriver.

## Misc. Features
1. *Config file*
A simple configuration file/module will be used to set the script parameters. These parameters will be set using the tkinter GUI or by manually changing the file.

2. *Interactive instructions*
A manual to help orient new users in how to use the program.

3. *Ghost Icon used as mascot*

## ERRORS
NeoSpecter is rigid in that it relies on the Excel data formats and the NeoSerra webforms to remain the same
over time. If the formating is changed with the excel datafile exports from EVA or NeoSerra changes the XPATH of their awards webform, the program *_will_* break.