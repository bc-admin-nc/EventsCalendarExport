# GETTING STARTED

This Repository contains a Typescript project which is setup to transpile to App Script and run within the context of a Google Sheet.  This code does the following:
1. reads entries from a Google calendar for a given date range (user provided)
2. create rows in the associated Google sheet (starting at row 2) which should be importable into eDirectory.
    - in order to generate these rows, the script does the following:
        - parses that calendar event data - which is often incomplete or inconsistent with known band and venues
        - uses fuzzing matching to match artists with known values in a Google Sheet
        - uses fuzzing matching to match venuse with known values in a Google Sheet
        - outputs the data in a format that is uploadable to eDirectory

# Project setup
## Recommened skills
- basic Typescript/Javascript skills
- basic knowledge of source control, secifically git
- basic knowledge of Google Sheets / Google Calendar
- the ability to use Google/ChatGPT

## Prerequsite installation
You will need the following development tools installed.  If you don't know what these are you'll need to familiarize yourself with these basic tools before proceeding.  You can google it or ask chatGPT.
- **git** - way to interact with a Git repository (a version control system) using text commands in your terminal, allowing you to perform actions like committing changes, managing branches, and pulling updates directly from the command line instead of using a graphical user interface
- **npm** - a registry of open-source JavaScript packages and a tool for managing dependencies.
- **clasp**: CLASP (Command Line Apps Script Project) is a command-line tool that helps developers manage Google Apps Script (GAS) projects more efficiently. It allows you to develop scripts locally, use version control (like Git), and deploy them without needing to use the online Apps Script editor.
- **tsc**: stands for TypeScript Compiler. It is a command-line tool that comes with TypeScript and is used to compile TypeScript code into JavaScript code.
- **VS Code** - a free, lightweight, and customizable source code editor developed by Microsoft that allows users to write and edit code across various programming languages

## Download and setup the development enviroment
``` 
mkdir EventsCalendarExport
cd EventsCalendarExport
git clone https://github.com/bc-admin-nc/EventsCalendarExport.git
```

## Download pre-requesite 3rd-party package
Execite the following command within the EventsCalendarExport folder (or wherever the package.json file resides)
```npm install```
This will install all of the third-party depencencies required by the project.  This project is NOT run locally, it is compiled and uploaded (via `clasp push`) to a Google Sheet and executed from there.

## Development 
### Overview
Currently the entire project is comprised of 2 file.  
- **DateForm.html** - used to collect user input, currently just the start and end dates for the calendar query
- **ExportCalendarToSheets.ts** - typescript file that queries the Google Calendar and other sheets and outputs a new Google Sheet which can me imported into eDirector.

### Logging into Google using clasp
You make the changes locally and push the changes to the Google cloud account.  Currently all scripts and assets (sheets, calendars, etc) are under the `localmusic@bandsandclubsofthetriangle.com` account.  In order to access this account you'll need to log into that account via the following command:
`clasp login`

- Navigate to the URL presented on the command line (if a browser doesn't open automatically), 
- Make sure `localmusic@bandsandclubsofthetriangle.com` selected. 
- Click Continue on the "Sign in to clasp – The Apps Script CLI" dialog.  
- Click 'Select all" and Continue on the "clasp – The Apps Script CLI wants access to your Google Account" dialog

Now your development environment is setup and you have permission to access the required Google assets.

### Associating a Google Sheet
This code executes as a container-bound script where the Google Sheet is the container/default context for the script.  You must first create a Google Sheet, or look up an exsiting Google Sheet in order to get it's __sheetId__.  The id can be found in URL.  It's the part after `https://docs.google.com/spreadsheets/d/` and before the next `/`.  It is a series of 44 letters, numbers, and symbols.

In order to associated this project with a Google Sheet, you will to update the `scriptId` value in `.clasp.json` the Google __sheetId__.  Once this value is updated, then you can make changes to the script in VS Code and push changes to the Google Sheet. 

Note: The code you see is the project is Typescript.  Typescript transpiled in to JavaScript which then gets transpiled into Google AppScript by clasp as part of the push process.  So the code you see in the online AppScript editor will vary slightly from what you see in the source project.  This is expected.

### Pushing changes to the Google Sheet
You will use VS Code to make changes to the .html or typescript code as needed.  Once the changes have been made locally to the typescript code, you can push those change the associated Google Sheet, 

To do this, use the following command:
```clasp push```
This will fail if you are not logged into Google using clasp

### Updating the git repository
Although the code is transpiled and copied into the associated Google Sheet, the original Typesript source code project lives in github.  Once you are ready to make your changes permanent, remember to do a `git commit ` and `git push` so that the source code on the github repository is updated with your local changes.

### To run the script from the Google Sheet
1. Open the Google Sheet:
```https://docs.google.com/spreadsheets/d/1a8PYEsNdVhQFG0-al3SKDRLsX5YyQq3_nWhu4rJiURw/edit?gid=0#gid=0```
2. Click Extentions > App Script - You should see the online App Script editor and the files that got uploaded when you did `clasp push`.  
**NOTE:**  I recommend you always make your changes to the Typescript file in the source project and use `clasp push` to modify the App Script bound to the sheet.  Avoid making permanent changes directly to the App Script.


### To debug the script from the Google Sheet
There are several test functions writen into the code that will help debug issues.  You can write more in Typescript, compile them  and run them from the online AppScript editory using the Debug function and selecting the specific test function to execute.  All test functions are prefixed with `test_`.
