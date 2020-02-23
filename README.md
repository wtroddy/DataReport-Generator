# DataReport-Generator
## Description:
This PS module will generate "pretty formatted" xlsx's of an input csv dataset.

Given a set of input parameters, an .xlsx file that will be generated that is ready for end users to receive with metadata in addition to the input csv data. 

## Example Use:
Data reports are generated using the function "DataReport" and can be invoked with different methods 1) using a reference .tsv file with relevant arguments or 2) by manually passing an array of arguments.

### 1) Reference File Use
Using a reference file is the default behavior of the function. The function works by taking in a input file with up to 7 variables. To support strings with commas this should be a tsv file. 


## Input-Output:
Use of the script will create an formatted excel document with the following characteristics and values. In an attempt to create additional flexibility for the scripts use, the variable position trumps the label of the variable. The result is that labels are custom to each user. 

Custom variables in the following cells:
|Variable Position|Spreadsheet Position|Formating|Notes|
|-----------------|--------------------|---------|-----|
|Reference Variable #1|G2, Name of Output File|||
|Reference Variable #2|A1, Name of Output File|Bolded, Font Size 13|note: if this is not provided by the user, then a value will be derived from the value from variable #6|
|Reference Variable #3|A2|Bolded||
|Reference Variable #4|A4|||
|Reference Variable #5|A6|||
|Reference Variable #6|G3|Bolded, Right Aligned|Name of CSV File Loaded - important: this should be a path to a .csv file and is used to identify data that is added to the report and must be included. e.g.: .\path\to\my\cool\data.csv|
|Reference Variable #7|G4 (Bolded, Right Aligned)
|Reference Label #1|F2 (Bolded, Right Aligned)
|Reference Label #2|||Not Output|
|Reference Label #3|||Not Output|
|Reference Label #4|A3|Bolded||
|Reference Label #5|A5|Bolded||
|Reference Label #6|F3|Bolded, Right Aligned||
|Reference Label #7|F4|Bolded, Right Aligned||

Data from the file referenced in variable #6 is loaded into the spreadsheet starting in row 10 with the input variable names bolded, a background cell color of gray added, and the cells locked to allow scrolling while keeping the variable names shown.

The output .xlsx file will be named as a combinated of either a) variable #1 and variable #2 or b) variable #1 and variable #6 and saved in the current directory in a new folder named "output". If no output folder exists than the script will generate one.

## Setup: 
### Environment Setup:
To load PowerShell modules, the .psm1 file needs to be saved in a folder that is in the PSModulePath. 

To see current PowerShell Module Paths run:
```
$env:PSModulePath.split(";")
```

To add a new path for the current session only, the generic form is:
```
# update PSModulePath for current PS session
$env:PSModulePath = $env:PSModulePath + ";C:\My\Path\To\PowerShell_Modules\"
```

To verify that the new modules are available, this command will list all available modules within each PSModulePath.
```
# list available modules 
Get-Module -ListAvailable
```
### Running Scripts:
Note: in order to run scripts, the Execution Policy needs to be set to the appropriate level. 
For example:
```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
will set the execution policy to allow for remotely signed scripts for the current user. 

Additional details are available via [Microsoft's website](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7) . 

