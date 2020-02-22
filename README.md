# DataReport-Generator
## Description:
This PS module will generate "pretty formatted" xlsx's of an input csv dataset.

Given a set of input parameters, an .xlsx file that will be generated that is ready for end users to receive with metadata in addition to the input csv data. 

## Example Use:


## Setup: 
### Environment Setup 
To load PowerShell modules, the .psm1 file needs to be saved in a folder that is in the PSModulePath. 

To see current PowerShell Module Paths run:
```
$env:PSModulePath.split(";")
```

To add a new path for the current session only, the generic form is:
```
$env:PSModulePath = $env:PSModulePath + ";C:\My\Path\To\PowerShell_Modules\"
```