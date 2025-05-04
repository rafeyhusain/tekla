# Overview:
This project contains two `Tekla Structures 2025` macros:

## ExportToExcel
Exports member information (ID, Name, Section, Start/End points, Comment) to an Excel file `Desktop\Model.xlsx`.
See the [Members](#Members) section for details on the exported data format. 

Note, you can run [Macro2](docs/xlsx/Macro2.xlsm) inside `Desktop\Model.xlsx` to generate `Connections` sheet.

## ImportFromExcel
Reads connections from an Excel file `Desktop\Model.xlsx` and creates connections between members in Tekla based on `Connections` sheet.
See the [Connections Sheet](#Connections) section for details on the exported data format.

## Install Software

- Install [Tekla Structures 2025 - Free Trial](https://download.tekla.com/tekla-structures/free-trial)
- Install [.NET Framework 4.8](https://dotnet.microsoft.com/en-us/download/dotnet-framework/thank-you/net48-web-installer)

## Steps to Use Macros

- Open desrired model e.g. `YourModelName` in `Tekla Structures 2025`
- Create 2 macros `ExportToExcel` and `ImportFromExcel` in Tekla Structures 2025
- Either download this git repo or its zip from GitHub
- Paste [`modeling`](docs/modeling) folder under `docs/modeling` in this repo to `C:\TeklaStructuresModels\<YourModelName>\macros\modeling`
- Or paste it under `C:\Users\YourUsername\AppData\Local\Trimble\Tekla Structures\version\Macros\` to be used globally across all models

## Run macros
- Go to `Edit` > `Components` > `Applications & Components` > `Search for your macro name`
- Double click `ExportToExcel` or `ImportFromExcel` in `Tekla Structures 2025` to run macros

## Model Excel File

This file is an output of [ExportToExcel](#ExportToExcel) Tekla macro.

### Members
| ID    | Name   | Section             | Start_X | Start_Y | Start_Z | End_X | End_Y | End_Z | Comment |
|-------|--------|---------------------|---------|---------|---------|-------|-------|-------|---------|
| 28771 | BEAM   | WI300-15-20*300     | 14400   | 24000   | 7200    | 7200  | 24000 | 7200  |         |
| 28605 | COLUMN | WI300-15-20*300     | 7200    | 24000   | -500    | 7200  | 24000 | 7200  |         |
| 26930 | COLUMN | WI300-15-20*300     | 14400   | 24000   | -500    | 14400 | 24000 | 7200  | Splice  |


### Connections
This file is an output of when you run Excel macro [Macro2](docs/xlsx/Macro2.xlsm).

| Preliminary ID | Pre. Name | Pre. Section        | Pre. Comment | Secondary ID | Sec. Name | Sec. Section        | Sec. Comment | Mid-Connection | Connection Type |
|----------------|-----------|---------------------|--------------|---------------|-----------|---------------------|---------------|----------------|-----------------|
| 28605          | COLUMN    | WI300-15-20*300     |              | 28771         | BEAM      | WI300-15-20*300     |               | 119            |                 |
| 26930          | COLUMN    | WI300-15-20*300     |              | 28771         | BEAM      | WI300-15-20*300     |               | 119            |                 |


Only rows with a non-empty `Selected` field (e.g. `Yes`, `yes` or `X`) will be processed.  

### Column Descriptions

- **ID**: Unique identifier of the Tekla part (used for connection mapping later).
- **Name**: Part name (`.Name` property).
- **Profile**: Cross-section or shape (e.g., W200X42).
- **StartX, StartY, StartZ**: Start point coordinates of the part.
- **EndX, EndY, EndZ**: End point coordinates of the part.
- **Comment**: Any custom comment from a UDA (optional).
