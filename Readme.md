# DFP Export Process

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

In the modal, paste "928077908796" in input field and click "Set Project".

![Set project ID](./img/project-id.png?raw=true)

You'll be asked to confirm. Modal should then be updated with a success message:

![Project ID success](./img/project-id-success.png?raw=true)

Close the modal and go to File > Project properties. You'll be asked to set a project name. This is not important, choose anything.

![Project properties](./img/project-properties.png?raw=true)

In the modal, copy the "Script ID" value:

![Script ID](./img/script-id.png?raw=true)

Edit the .clasp.json file and paste the script ID in "scriptId" parameter. You can close the "Script editor" tab, you won't need it anymore.

Now you can execute the script:
```bash
clasp push && clasp run main -p '["Key-values"]'
```

When you'll be asked if you want to push and overwrite, choose "y".
