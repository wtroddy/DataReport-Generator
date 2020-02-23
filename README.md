# DataReport-Generator
## Description:
This PS module will generate "pretty formatted" xlsx's of an input csv dataset.

Given a set of input parameters, an .xlsx file that will be generated that is ready for end users to receive with metadata in addition to the input csv data. 

## Example Use:


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

