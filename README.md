## Instructions:

### First create a file with a streaming formula using the OJS add-in
1. Use node version 18.18.0 (`nvm install 18.18.0 && nvm use 18.18.0`)
2. Install node modules (`npm install`)
3. Build app (`npm run build`)
4. Run node server by opening a terminal in project root and running `npm run dev-server`
5. Then run addin in excel by opening another terminal in project root: `npm run start:desktop`
6. PC Excel will open automatically
7. Type in Cell A1 `=TESTRTD(1,2)`, it would be streaming an incrementing value (ex. `OJS: 6`)
8. Save this file under `streaming-ojs`
9. Close the file

### Next we clear the OJS add-in from PC Excel
1. Go to the terminal running the node server and kill it
2. [Clear PC Excel Web add-in cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache#manually-clear-the-cache-in-excel-word-and-powerpoint) by deleting this folder %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\

### Next we install the COM add-in
1. Navigate to `the test-rtd-server` folder
2. Install the COM add-in by running `Install.bat` as admin
3. Open a new Excel file
4. Type in Cell A1 `=TESTRTD(1,2)`, check that it is streaming this value ex. `COM: #` (ex: `COM: 6`)
    - note that if the result it `#VALUE`:
      - make sure that you ran `Install.bat` as admin
      - close Excel and try opening it again
5. Close without saving

### Last we reproduce the bug by opening `streaming-ojs`
1. Open the previously saved excel file `streaming-ojs`
2. Notice that A1 which contains `=TESTRTD(1,2)` is not streaming and is stuck at with the cached web value `OJS: #` 
3. In A2 type `=TESTRTD(1,2)` and notice that it is streaming `COM: #`
4. Save and close excel file
5. Open it once again and notice that A2 resumes streaming but A1 is still stuck
6. Lastly enter A1, and notice it starts streaming using the COM formula: `COM: #`

#### Notes:
- source code for sample COM add-in is inside `test-rtd-server` folder
- the OJS code was generated using [yeoman generator in MS docs](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview)


---
## OLD INSTRUCTIONS
COM add-in

- source code for sample COM add-in is in zip file inside `test-rtd-server`

1. Navigate to `the test-rtd-server` folder
2. Run `Install.bat` as admin (This is the COM/Excel add-in)
3. Run Excel
4. Type in `=TESTRTD(1,2)`, it should stream an incrementing value (ex. COM: 7)
5. Press Alt + F, then T
6. Navigate to Addins, you will see the two add-ins for the COM/Excel add-in

OJS add-in

- We will also be installing the OJS add-in alongside the COM add-in to replicate the issue
- The manifest file includes an Equivalent add-in tag for the COM add-in used above

1. nvm install 18.18.0
2. npm install
3. npm run build (Make sure you are in project root)
4. Open a terminal in project root: run "npm run dev-server"
5. Open another terminal in project root: "npm run start:desktop"
6. Excel should launch automatically
7. Type in `=TESTRTD(1,2)`, it would be streaming an incrementing value (ex. OJS: 6)
8. Press Alt + F, then T
9. Navigate to Addins, you will see that only one add-in from the COM addin is visible
10. OJS is overriding COM addin's `=TESTRTD` formula

When relaunching the add-in to avoid the error message

1. Ensure that the server on the first terminal is still running, do not close this terminal.
2. Head on to: %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
3. Delete the Resources folder.
4. Open another terminal in project root: "npm run start:desktop"
5. "Waiting for add-ins runtime" message should not appear.
