# DFP Export Processor

This project contains a script to process a DFP Export (saved as a Google Drive spreadsheet) in order to:
- add filters on input columns
- add "Total" & "Average" cells on output columns (based on filters choices)
- generate a sheet per specified dimension containing one chart per output type

## Installation

```bash
git clone https://github.com/lgiraudel/dfp-export-processor.git
npm install
```

## Usage

In a spreadsheet, go to Tools > Script Editor

![Script Editor](./img/script-editor.png?raw=true)

In the new window, go to Resources > Cloud platform project:

![Cloud platform project](./img/cloud-platform-project.png?raw=true)

You'll be asked to set a project name. This is not important, choose anything and save.
In the new modal, paste "928077908796" in input field and click "Set Project".

![Set project ID](./img/project-id.png?raw=true)

You'll be asked to confirm. Modal should then be updated with a success message:

![Project ID success](./img/project-id-success.png?raw=true)

Close the modal and go to File > Project properties.

![Project properties](./img/project-properties.png?raw=true)

In the modal, copy the "Script ID" value:

![Script ID](./img/script-id.png?raw=true)

Edit the .clasp.json file and paste the script ID in "scriptId" parameter. You can close the "Script editor" tab, you won't need it anymore.

___

If this is the first time you execute the script, you'll have to configure several thing to allow CLASP (Google's command line library to communicate with Google App Script projects) to execute the script.

First, run:
```
npx clasp login
```
It will open a page in your browser asking for some autorizations. Allowing it will generate a .clasprc.json file in your home directory.

Then execute:
```
npx clasp open --creds
```
It will open a page where you should see a list of Oath 2.0 client IDs. Download the "DFP Export Processor".

![OAuth credential](./img/oauth.png?raw=true)

Move the file to the dfp-export-processor project and rename it "creds.json":

```bash
mv ~/Downloads/client_secret_928077908796-khee3hv0egg56q1951u1df6ro2i7jn7c.apps.googleusercontent.com.json creds.json
```

Then use following command:
```
npx clasp login --creds creds.json
```
It will create a local `.clasprc.json` file.

Finally, you'll have to enable Google Apps Script API in your Google Account by going to https://script.google.com/u/1/home/usersettings

![Enable Google Apps Script API](./img/enable-api.png?raw=true)

![Google Apps Script API enabled](./img/api-enabled.png?raw=true)

___

Now you can execute the script:
```bash
npx clasp push && npx clasp run main -p '["Key-values"]'
```

When you'll be asked if you want to push and overwrite, choose "y".

## TODO

[ ] wrap script with a npm command, passing it the script ID to generate the `.clasp.json` file.