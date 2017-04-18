# RCTC Transform
This is a proof-of-concept for converting the Reportable Conditions Trigger Codes table (RCTC) from Excel to JSON and XML formats.  JSON or XML files are simpler and easier for computer programs to interact with.  This transformation helps to accommodate automatic incorporation of trigger code data into EHR systems and other health IT products.

## Requirements
This program is written in [Node.js](https://nodejs.org/en/).  Make sure it is installed before you start working with `rctc_transform`

## Installation
Install the required node modules:
- Navigate to the `rctc_transform` directory
- `npm install`

## Usage
This application is a command-line utility that takes an Excel RCTC file as its only argument, for example:
- `node index.js {full path to file}`
- You should see the following:
  - `The xml file was saved: {full path to new xml file}`
  - `The json file was saved: {full path to new json file}`
