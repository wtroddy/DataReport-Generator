<#
  .Synopsis
  Generate a report with manually defined parameters
  
  .Description
  some description here

  .Parameter input_path
  describe
     
#>

function DataReport-Manual{
    # define input parameters 
    param( 
    [string[]] $input_path,
    $RequestID, 
    $RequestTitle, 
    $RequestSubTitle, 
    $RequestDescription, 
    $RequestExtractDataDate, 
    $root_dir
    )

    ### set default and input paramaters
    # directory and file management 

    # set root dir if there wasn't an argument passed in
    if (!($root_dir)) {
        $root_dir = pwd
    }

    # join strings
    $output_dirname = Join-Path $root_dir "\output"

    # check if the output directory exists, if fnot create it
    if (!(Test-Path $output_dirname)) {
        New-Item -ItemType Directory -Force -Path $output_dirname   
    }

    ### set variables if they don't exist 
    # Request ID
    if (!($RequestID)) {
        $RequestID = "No_RequestID"
    }

    # Title
    if (!($RequestTitle)) {
        # if there isn't a title provided just use the filename 
        $RequestTitle = ($input_path.Split("\")[-1].Split(".")[0])
    }

    # Output Filename
    if (!($output_fname)) {
        $output_fname = "\" + $RequestID + "_" + ($RequestTitle -replace "\s+", "") + ".xlsx"
    }

    # create new excel object 
    $excel = New-Object -ComObject excel.application

    # show the excel object 
    $excel.visible = $True

    # add a workbook 
    $workbook = $excel.Workbooks.Add()

    # remove gridlines 
    $excel.ActiveWindow.DisplayGridlines = $false

    # create a variable referencing the first sheet of the wb
    $ws= $workbook.Worksheets.Item(1)

    # change the name of the sheet 
    $ws.Name = "$RequestTitle"

    # fill cells with header information 
    $ws.Cells.Item(1,1) = $RequestTitle
    $ws.Cells.Item(2,1) = $RequestSubTitle
    $ws.Cells.Item(3,1) = "Request ID: "
    $ws.Cells.Item(3,2) = $RequestID
    $ws.Cells.Item(4,1) = "Date of Data Extract: "
    $ws.Cells.Item(4,2) = $RequestExtractDataDate
    $ws.Cells.Item(5,1) = "Request Description: "
    $ws.Cells.Item(5,2) = $RequestDescription

    # internal use data 
    $ws.Cells.Item(1,6) = "Internal Use"
    $ws.Cells.Item(2,6) = "Raw Data: "
    $ws.Cells.Item(2,7) = $input_path
    $ws.Cells.Item(3,6) = "Code: "
    $ws.Cells.Item(3,7) = "Need variable for this"

    # format header data
    $ws.Range("A1:A5").Font.Bold=$True              # BOLD - range A1:A5
    $ws.Range("F1:F5").Font.Bold=$True               # BOLD - range F1:F5
    $ws.Cells.item(1,1).Font.Size=13                # SIZE - Title = 13
    $ws.Range("A3:A5").HorizontalAlignment = -4152  # RIGHT ALIGN - range A3:A5
    $ws.Range("B3:B5").HorizontalAlignment = -4131  # LEFT ALIGN - range B3:B5
    $ws.Range("F1:F5").HorizontalAlignment = -4152  # RIGHT ALIGN - range A3:A5

    # add data 
    $i = 10
    Import-Csv $input_path | ForEach-Object {
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


    #freeze the top rows
    $ws.Rows.Item("11:11").Select()
    $ws.Application.ActiveWindow.FreezePanes = $true


    # Save and Quit 
    $workbook.SaveAs($output_dirname+$output_fname)
    $excel.Quit()

}


function DataReport{
    # define input parameters 
    param( 
    $input_data
    )

    # set root dir to the pwd
    $root_dir = pwd

    # join strings
    $output_dirname = Join-Path $root_dir "\output"

    # check if the output directory exists, if fnot create it
    if (!(Test-Path $output_dirname)) {
        New-Item -ItemType Directory -Force -Path $output_dirname
    }

    # add data 
    $indata = Import-Csv -Delimiter "`t" -path $input_data

    $request_ids = $indata.RequestID | Sort-Object | Get-Unique

    ForEach($id in $request_ids){
        # create new excel object 
        $excel = New-Object -ComObject excel.application

        # show the excel object 
        #$excel.visible = $True

        # add a workbook 
        $workbook = $excel.Workbooks.Add()

        # get the current request
        $CurReq = $indata | Where-Object {$_.RequestID -eq $id}

        # get the length of the object
        $obj_details = $CurReq | Measure-Object

        $number_entries = $obj_details.Count

            # current entry counter
            $cur_entry = 1

            ForEach($CurData in $CurReq){

                # Title
                if ($CurReq[0].RequestTitle) {
                    $RequestTitle = $CurReq[0].RequestTitle
                } else {
                    $RequestTitle = ($CurReq[0].input_path.Split("\")[-1].Split(".")[0])
                }

                # Output Filename
                $output_fname = "\" + $CurReq[0].RequestID + "_" + ($CurData[0].RequestTitle -replace "\s+", "") + ".xlsx"

                # create a variable referencing the first sheet of the wb
                $ws= $workbook.Worksheets.Item(1)

                # change the name of the sheet
                $ws.Name = $CurData.RequestSubTitle

                # add data 
                $i = 10
                Import-Csv $CurData.input_path | ForEach-Object {
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
                $ws.Cells.Item(1,1) = $CurData.RequestTitle
                $ws.Cells.Item(2,1) = $CurData.RequestSubTitle
                $ws.Cells.Item(3,1) = "Date of Data Extract: "
                $ws.Cells.Item(4,1) = $CurData.RequestExtractDate
                $ws.Cells.Item(5,1) = "Request Description: "
                $ws.Cells.Item(6,1) = $CurData.RequestDescription

                # internal use data 
                $ws.Cells.Item(1,6) = "Internal Use"
                $ws.Cells.Item(2,6) = "Request ID: "
                $ws.Cells.Item(2,7) = $CurData.RequestID
                $ws.Cells.Item(3,6) = "Data: "
                $ws.Cells.Item(3,7) = $CurData.input_path
                $ws.Cells.Item(4,6) = "Code: "
                $ws.Cells.Item(4,7) = $CurData.code_path

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
                
                #autotfi
            
                # remove gridlines for a clean background 
                $excel.ActiveWindow.DisplayGridlines = $false
            
                # add a new sheet for the next round 
                if (!($cur_entry -eq $number_entries)) {
                    $ws = $workbook.Worksheets.add() | Out-Null
                    $cur_entry++
                }
            
            }
        # Save and Quit 
        $workbook.SaveAs($output_dirname+$output_fname)
        $excel.Quit()

    }

}