# DataReport-Generator
## Description:
This PS module will generate "pretty formatted" xlsx's of an input csv dataset.

Given a set of input parameters, an .xlsx file that will be generated that is ready for end users to receive with metadata in addition to the input csv data. 

## Contents:
[Example Use](##Example Use)

## Example Use:
Data reports are generated using the function "DataReport" and can be invoked with different methods 1) using a reference .tsv file with relevant arguments or 2) by manually passing an array of arguments.

### 1) Reference File Use
Using a reference file is the default behavior of the function. The function works by taking in a input file with up to 7 variables. To support strings with commas this should be a tsv file. This functionality supports creating multiple sheets per workbook and multiple workbook with a single function. Multiple sheets are added to workbook when they share a common value for the first variable (e.g. the ID) of the input file. 

The general syntax for using a reference file is:
```
DataReport ".\input_ref_file.txt"
```

However, it should noted that this can be called more explicitly. Such as:
```
DataReport -input_data_file ".\input_ref_file.txt"
```

The input_data_file is anticipated to include 7 columns of reference data, as defined in the Input-Output and Example Files sections below.

### 2) Command Line Argument Use (Manual Mode)
It is also possible to pass arguments for a single sheet to be generated using the same function. 

The general syntax for using the "manual mode" is:
```
DataReport -input_mode_manual:$true -input_data:@("ID","Title","Subtitle","Date","Description","Input_Path","Code_Path") -input_labels:@("00000", "My Data", "Raw Data Name", "YYYY-MM-DD", "details on the data", "path\to\my\cool\data.csv", "path\to\my\cool\code.sql)
```

It should be noted that a more dynamic use would be to define a variable with the array and that this may be called as a manual input parameter.
```
$label_array = @("ID","Title","Subtitle","Date","Description","Input_Path","Code_Path")
$data_array = @("00000", "My Data", "Raw Data Name", "YYYY-MM-DD", "details on the data", "path\to\my\cool\data.csv", "path\to\my\cool\code.sql)
DataReport -input_mode_manual:$true -input_data:$data_array -input_labels:$label_array
```

### Example Files
#### Example Input Data File 
An example input data file may be formatted as:

|RequestID|Title|SubTitle|Date|Description|Input_Path|Code_Path|
|---------|-----|--------|----|-----------|----------|---------|
|00000|My Data|Raw Data Name|YYYY-MM-DD|Details on the data|\path\to\my\cool\data.csv |\path\to\my\cool\code.sql|
|00000|My Data|Statistics|YYYY-MM-DD|Details on the data|\path\to\my\cool\data.csv |\path\to\my\cool\code.R|


#### Example Output Data File
The example input will generate a pretty formated xlsx workbook with two sheets. An example of the first row is shown with this relative spacing:

|My Data|    |    |    |    |    |    |    |
|-------|----|----|----|----|----|----|----|
|Raw Data Name| | | | |Input_Path|\path\to\my\cool\data.csv||
|Date| | | | |Code_Path|\path\to\my\cool\code.sql||
|YYYY-MM-DD| | | | | | ||
|Description| | | | | | ||
|Detail on the data| | | | | | ||
| | | | | | | ||
| | | | | | | ||
|CSV-V1|CSV-V2|CSV-V3|CSV-V4|CSV-V5|CSV-V6|CSV-V7|...|
|Data|Data|Data|Data|Data|Data|Data|...|

## Input-Output:
Use of the script will create an formatted excel document with the following characteristics and values. In an attempt to create additional flexibility for the scripts use, the variable position trumps the label of the variable. The result is that labels are custom to each user. 

Custom variables in the following cells:
|Variable Position|Spreadsheet Position|Formating|Notes|
|-----------------|--------------------|---------|-----|
|Reference Variable #1|G2, Name of Output File|||
|Reference Variable #2|A1, Name of Output File|Bolded, Font Size 13|note: if undefined then a value will be derived from the value from variable #6|
|Reference Variable #3|A2|Bolded||
|Reference Variable #4|A4|||
|Reference Variable #5|A6|||
|Reference Variable #6|G3|Bolded, Right Aligned|Name of CSV File Loaded^|
|Reference Variable #7|G4|Bolded, Right Aligned||
|Reference Label #1|F2|Bolded, Right Aligned||
|Reference Label #2|||Not Output|
|Reference Label #3|||Not Output|
|Reference Label #4|A3|Bolded||
|Reference Label #5|A5|Bolded||
|Reference Label #6|F3|Bolded, Right Aligned||
|Reference Label #7|F4|Bolded, Right Aligned||

^important: this should be a path to a .csv file and is used to identify data that is added to the report and must be included. e.g.: .\path\to\my\cool\data.csv

Data from the file referenced in variable #6 is loaded into the spreadsheet starting in row 10 with the input variable names bolded, a background cell color of gray added, and the cells locked to allow scrolling while keeping the variable names shown.

The output .xlsx file will be named as a combinated of either a) variable #1 and variable #2 or b) variable #1 and variable #6 and saved in the current directory in a new folder named "output". If no output folder exists than the script will generate one.

## Setup: 
### Environment Setup:
To load PowerShell modules, the .psm1 file needs to be saved in a folder that is in the PSModulePath. 

To see current PowerShell Module Paths run:
```
# get list of current PS Module Paths
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

Modules can be loaded using the command:
```
# import the modules 
Import-Module -Name DataReport-Generator
```

### Running Scripts:
Note: in order to run scripts, the Execution Policy needs to be set to the appropriate level. 
For example:
```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
will set the execution policy to allow for remotely signed scripts for the current user. 

Additional details are available via [Microsoft's website](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7) . 

