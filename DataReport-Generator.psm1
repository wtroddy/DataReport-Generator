<#
  .Synopsis
  Generate a report with manually defined parameters
  
  .Description
  This PS module will generate "pretty formatted" xlsx's of an input csv dataset.
  Given a set of input parameters, an .xlsx file that will be generated that is ready for end users to receive with metadata in addition to the input csv data. 

  .Parameter input_data_file
  The path to an input tab separated file for batch creation of pretty formated xlsx files.

  .Parameter input_mode_manual
  True/False flag to manually define input parameters instead of using a batch file. The default value of $false will load a file.

  .Parameter input_labels
  Input array of labels for input_data. See README for additional details on use and optional arguments.
  Recommended values: $label_array = @("ID","Title","Subtitle","Date","Description","Input_Path","Code_Path", "Directory_Path").  

  .Parameter input_data
  Input array of data/metadata to fill into xlsx. See README for additional details on use and optional arguments.
  Recommended values: $data_array = @("00000", "My Data", "Raw Data Name", "YYYY-MM-DD", "details on the data", "path\to\my\cool\data.csv", "path\to\my\cool\code.sql, "\path\to\my\pretty\data")

  .Parameter write_checksum
  Optional argument to include the MD5 checksum in output file. Default value is $true. 
     
#>

function DataReport{
    # define input parameters 
    param(
    $input_data_file, 
    [switch]$input_mode_manual,
    $input_labels,
    $input_data,
    $write_checksum
    )

    # parameter management
    if ($write_checksum -eq $null) {
        $write_checksum = $true
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
                        var8 = $input_data[7]
                        }

            # relabel input labels in hash 
            $input_labels = @{var1 = $input_labels[0]
                              var2 = $input_labels[1] 
                              var3 = $input_labels[2]
                              var4 = $input_labels[3]
                              var5 = $input_labels[4] 
                              var6 = $input_labels[5]
                              var7 = $input_labels[6]
                              var8 = $input_labels[7]
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
                if ($CurData.var4) {
                    $RequestTitle = $CurData.var4
                } else {
                    $RequestTitle = ($CurData.var2.Split("\")[-1].Split("."))
                }

                ### Output Management
                # Output Folder
                if ($CurData.var8) {
                    $output_dirname = $CurData.var8
                } else {
                    $output_dirname = pwd
                }

                # Output Filename
                $output_fname = "\" + $CurData.var1 + "_" + ($RequestTitle -replace "\s+", "") + ".xlsx"

                # create a variable referencing the first sheet of the wb
                $ws= $workbook.Worksheets.Item(1)

                # change the name of the sheet
                # get the value
                $new_sheetname = $CurData.var5    ### this is a janky solution but it works
				
                # set the value 
                $ws.Name = "$new_sheetname"

                ### add MD5 checksums 
                $CSV_File_Hash = Get-FileHash $CurData.var2 -Algorithm MD5
                
                
                ### load the csv and loop through data 
                $i = 11
                Import-Csv $CurData.var2 | ForEach-Object {
                    $j = 1
                    foreach ($prop in $_.PSObject.Properties) {
                        if ($i -eq 11) {
                            # add header
                            $ws.Cells.Item($i-1, $j).Value = $prop.Name
                            
                            # header formatting 
                            $ws.Cells.Item($i-1,$j).Interior.ColorIndex = 15
                            $ws.Cells.Item($i-1,$j).Font.Bold=$True

                            # add first row of data 
                            $ws.Cells.Item($i, $j++).Value = $prop.Value
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
                $ws.Cells.Item(1,1) = $CurData.var4
                $ws.Cells.Item(2,1) = $CurData.var5
                $ws.Cells.Item(3,1) = $input_labels.var6
                $ws.Cells.Item(4,1) = $CurData.var6
                $ws.Cells.Item(5,1) = $input_labels.var7
                $ws.Cells.Item(6,1) = $CurData.var7

                # internal use data 
                $ws.Cells.Item(1,6) = "Internal Use"
                $ws.Cells.Item(2,6) = $input_labels.var1
                $ws.Cells.Item(2,7) = $CurData.var1
                $ws.Cells.Item(3,6) = $input_labels.var2
                $ws.Cells.Item(3,7) = $CurData.var2
                $ws.Cells.Item(4,6) = $input_labels.var3
                $ws.Cells.Item(4,7) = $CurData.var3
                # check if the checksum flag is true, if yes add this datta
                if ($write_checksum -eq $true) {
                    $ws.Cells.Item(5,6) = "Input MD5 Checksum"
                    $ws.Cells.Item(5,7) = $CSV_File_Hash.Hash 
                }

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