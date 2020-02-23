﻿<#
  .Synopsis
  Generate a report with manually defined parameters
  
  .Description
  some description here

  .Parameter input_path
  describe
     
#>

function DataReport{
    # define input parameters 
    param(
    $input_data_file, 
    [switch]$input_mode_manual,
    $input_labels,
    $input_data
    )

    ### directory management 
    # set root dir to the pwd
    $root_dir = pwd

    # join strings
    $output_dirname = Join-Path $root_dir "\output"

    # check if the output directory exists, if fnot create it
    if (!(Test-Path $output_dirname)) {
        New-Item -ItemType Directory -Force -Path $output_dirname
    }

    ### conditionally load data based on the input mode 
    # for tsv input 
    if ($input_mode_manual -eq $false) { 
        # if the input mode is not manual, load in the csv
        $indata = Import-Csv -Delimiter "`t" -path $input_data_file -Header var1, var2, var3, var4, var5, var6, var7
        
        # split the first values into the header
        $input_labels = $indata[0]
        
        # update the indata object to just be data 
        $indata = $indata[1..($indata.Count-1)]
    } elseif ($input_mode_manual -eq $true) {
            # relabel input data in hash 
            $indata = @{var1 = $input_data[0]
                        var2 = $input_data[1] 
                        var3 = $input_data[2]
                        var4 = $input_data[3]
                        var5 = $input_data[4] 
                        var6 = $input_data[5]
                        var7 = $input_data[6]
                        }

            # relabel input labels in hash 
            $input_labels = @{var1 = $input_labels[0]
                              var2 = $input_labels[1] 
                              var3 = $input_labels[2]
                              var4 = $input_labels[3]
                              var5 = $input_labels[4] 
                              var6 = $input_labels[5]
                              var7 = $input_labels[6]
                             }
    }

    ### prep for processing by requests input 
    # get the ids from the request input 
    $request_ids = $indata.var1 | Sort-Object | Get-Unique

    ### loop for each ID in the input
    ForEach($id in $request_ids){
        # create new excel object 
        $excel = New-Object -ComObject excel.application

        # add a workbook 
        $workbook = $excel.Workbooks.Add()

        # get the current request
        $CurReq = $indata | Where-Object {$_.var1 -eq $id}

        # get the length of the object
        $obj_details = $CurReq | Measure-Object

        $number_entries = $obj_details.Count

            # current entry counter
            $cur_entry = 1

            ForEach($CurData in $CurReq){

                ### Title
                if ($CurData.var1) {
                    $RequestTitle = $CurData.var2
                } else {
                    $RequestTitle = ($CurData.var2.Split("\")[-1].Split("."))
                }

                # Output Filename
                $output_fname = "\" + $CurData.var1 + "_" + ($CurData.var2 -replace "\s+", "") + ".xlsx"

                # create a variable referencing the first sheet of the wb
                $ws= $workbook.Worksheets.Item(1)

                # change the name of the sheet
                # get the value
                $new_sheetname = $CurData.var3    ### this is a janky solution but it works
                # set the value 
                $ws.Name = "$new_sheetname"

                # add data 
                $i = 10
                Import-Csv $CurData.var6 | ForEach-Object {
                    $j = 1
                    foreach ($prop in $_.PSObject.Properties) {
                    if ($i -eq 10) {
                        $ws.Cells.Item($i, $j).Value = $prop.Name

                        # formatting 
                        $ws.Cells.Item($i,$j).Interior.ColorIndex = 15
                        $ws.Cells.Item($i,$j++).Font.Bold=$True

                    } else {
                        $ws.Cells.Item($i, $j++).Value = $prop.Value
                    }
                    }
                    $i++
                }

                # find the filled cells and autofit them 
                $usedRange = $ws.UsedRange						
                $usedRange.EntireColumn.AutoFit() | Out-Null

                # fill cells with header information 
                $ws.Cells.Item(1,1) = $CurData.var2
                $ws.Cells.Item(2,1) = $CurData.var3
                $ws.Cells.Item(3,1) = $input_labels.var4
                $ws.Cells.Item(4,1) = $CurData.var4
                $ws.Cells.Item(5,1) = $input_labels.var5
                $ws.Cells.Item(6,1) = $CurData.var5

                # internal use data 
                $ws.Cells.Item(1,6) = "Internal Use"
                $ws.Cells.Item(2,6) = $input_labels.var1
                $ws.Cells.Item(2,7) = $CurData.var1
                $ws.Cells.Item(3,6) = $input_labels.var6
                $ws.Cells.Item(3,7) = $CurData.var6
                $ws.Cells.Item(4,6) = $input_labels.var7
                $ws.Cells.Item(4,7) = $CurData.var7

                # format header data
                $ws.Range("A1:A3").Font.Bold=$True              # BOLD - range A1:A3
                $ws.Range("A5:A5").Font.Bold=$True              # BOLD - range A5
                $ws.Range("F1:F5").Font.Bold=$True              # BOLD - range F1:F5
                $ws.Cells.item(1,1).Font.Size=13                # SIZE - Title = 13
                $ws.Range("A3:A5").HorizontalAlignment = -4131  # LEFT ALIGN - range A3:A5
                $ws.Range("B3:B5").HorizontalAlignment = -4131  # LEFT ALIGN - range B3:B5
                $ws.Range("F2:F5").HorizontalAlignment = -4152  # RIGHT ALIGN - range F2:F5
                $ws.Range("G2:G5").HorizontalAlignment = -4131  # LEFT ALIGN - range G3:G5
                $ws.Range("F1:G1").MergeCells = $true           # MERGE - range F1:G1
                $ws.Range("F1:G1").HorizontalAlignment = -4108  # CENTER - range F1:G1

                #freeze the top rows
                $ws.Rows.Item("11:11").Select()
                $ws.Application.ActiveWindow.FreezePanes = $true
            
                # remove gridlines for a clean background 
                $excel.ActiveWindow.DisplayGridlines = $false
            
                # add a new sheet for the next round 
                if (!($cur_entry -eq $number_entries)) {
                    $ws = $workbook.Worksheets.add() | Out-Null
                    $cur_entry++
                }
            
            }
        # Save and Quit 
        $workbook.SaveAs($output_dirname+$output_fname) | Out-Null
        $excel.Quit() | Out-Null

    }

}