
Sub ExtractEDIData()
 
    'Change File Locations to match where your data is
    ' List of workbooks we will loop through
    TulsaWorkbook = "R:\Operations\Pre-Production\Planning\Planning_Erica\EDI\Templates\Tulsa.xlsx"
    SagWorkbook = "R:\Operations\Pre-Production\Planning\Planning_Erica\EDI\Templates\Sag.xlsx"
    MxLine1Workbook = "R:\Operations\Pre-Production\Planning\Planning_Erica\EDI\Templates\MxLine1.xlsx"
    MxLine2Workbook = "R:\Operations\Pre-Production\Planning\Planning_Erica\EDI\Templates\MxLine2.xlsx"
    RollinsWorkbook = "R:\Operations\Pre-Production\Planning\Planning_Erica\EDI\Templates\Rollins.xlsx"
    
    ' Create the sheets we need in this workbook
    ThisWorkbook.Worksheets.Add.Name = "extractData"
    ThisWorkbook.Worksheets.Add.Name = "temp"
    ThisWorkbook.Worksheets.Add.Name = "Consolidated"
    
    ' Perform the data extraction from the workbooks
    ExtractWorkbookData (TulsaWorkbook)
    ExtractWorkbookData (SagWorkbook)
    ExtractWorkbookData (MxLine1Workbook)
    ExtractWorkbookData (MxLine2Workbook)
    ExtractWorkbookData (RollinsWorkbook)
    
    ' Generate the consolidated output
    FormatConsolidated
End Sub

Sub StartTulsaDataExtract()

    'Change File Locations to match where your data is
    TulsaWorkbook = "R:\Operations\Pre-Production\Planning\Planning_Erica\EDI\Templates\Tulsa.xlsx"
    
    ThisWorkbook.Worksheets.Add.Name = "extractData"
    ThisWorkbook.Worksheets.Add.Name = "temp"
    ThisWorkbook.Worksheets.Add.Name = "Consolidated"
    
    ExtractWorkbookData (TulsaWorkbook)

    FormatConsolidated
End Sub

Sub StartSagDataExtract()

    'Change File Locations to match where your data is
    SagWorkbook = "R:\Operations\Pre-Production\Planning\Planning_Erica\EDI\Templates\Sag.xlsx"
    
    ThisWorkbook.Worksheets.Add.Name = "extractData"
    ThisWorkbook.Worksheets.Add.Name = "temp"
    ThisWorkbook.Worksheets.Add.Name = "Consolidated"
    
    ExtractWorkbookData (SagWorkbook)

    FormatConsolidated
End Sub

Sub StartMxLine1DataExtract()

    'Change File Locations to match where your data is
    MxLine1Workbook = "R:\Operations\Pre-Production\Planning\Planning_Erica\EDI\Templates\MxLine1.xlsx"
    
    ThisWorkbook.Worksheets.Add.Name = "extractData"
    ThisWorkbook.Worksheets.Add.Name = "temp"
    ThisWorkbook.Worksheets.Add.Name = "Consolidated"
    
    ExtractWorkbookData (MxLine1Workbook)

    FormatConsolidated
End Sub

Sub StartMxLine2DataExtract()

    'Change File Locations to match where your data is
    MxLine2Workbook = "R:\Operations\Pre-Production\Planning\Planning_Erica\EDI\Templates\MxLine2.xlsx"
    
    ThisWorkbook.Worksheets.Add.Name = "extractData"
    ThisWorkbook.Worksheets.Add.Name = "temp"
    ThisWorkbook.Worksheets.Add.Name = "Consolidated"
    
    ExtractWorkbookData (MxLine2Workbook)

    FormatConsolidated
End Sub

Sub StartRollinsDataExtract()

    'Change File Locations to match where your data is
    RollinsWorkbook = "R:\Operations\Pre-Production\Planning\Planning_Erica\EDI\Templates\Rollins.xlsx"
    
    ThisWorkbook.Worksheets.Add.Name = "extractData"
    ThisWorkbook.Worksheets.Add.Name = "temp"
    ThisWorkbook.Worksheets.Add.Name = "Consolidated"
    
    ExtractWorkbookData (RollinsWorkbook)

    FormatConsolidated
End Sub

Sub ExtractWorkbookData(workbookFilePath As String)

    ' Open the external workbook
    Dim wb As Workbook: Set wb = Workbooks.Open(workbookFilePath)
    
    ' Define each sheet/tab that we will access in the workbook
    Dim sheet648BSource As Worksheet: Set sheet648BSource = wb.Worksheets("648B")
    Dim sheet677BSource As Worksheet: Set sheet677BSource = wb.Worksheets("677B")
    Dim sheet680ASource As Worksheet: Set sheet680ASource = wb.Worksheets("680A")
    Dim sheet744ASource As Worksheet: Set sheet744ASource = wb.Worksheets("744A")
    
    ' Set the sheet/tab in this workbook where we will place the extracted data
    Dim extractDestination As Worksheet: Set extractDestination = ThisWorkbook.Worksheets("extractData")
    
    ' Define the variable for the last row in each sheet/tab
    Dim lastRow648B As Long
    Dim lastRow677B As Long
    Dim lastRow680A As Long
    Dim lastRow744A As Long
    
    ' Set the last row for each sheet/tab. This will be used to loop through the data
    If WorksheetFunction.CountA(sheet648BSource.UsedRange) > 0 Then
        lastRow648B = sheet648BSource.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
    If WorksheetFunction.CountA(sheet677BSource.UsedRange) > 0 Then
        lastRow677B = sheet677BSource.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
    If WorksheetFunction.CountA(sheet680ASource.UsedRange) > 0 Then
        lastRow680A = sheet680ASource.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
    If WorksheetFunction.CountA(sheet744ASource.UsedRange) > 0 Then
        lastRow744A = sheet744ASource.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
    
    ' Define variables to be used in the data extraction. These get reused during the loops.
    ' Quantity does not need set since it is pulled directly
    Dim partNumber As String
    Dim dateString As String
    
    ' Find which row we need to start inserting data in the extractData sheet
    Dim extractRow As Long: extractRow = extractDestination.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    ' Loop to write data into extractData sheet
    With extractDestination
        .Cells(1, 1) = "Part Number"
        .Cells(1, 2) = "Quantity"
        .Cells(1, 3) = "Date"
        
        ' Loop for the 648B sheet/tab. This will loop from row 1 until the last row of the sheet/tab
        For currentRow648B = 1 To lastRow648B
            ' If first column of the row is empty, skip it
            If sheet648BSource.Cells(currentRow648B, 1).Value = "" Then
                
            ' If first column of the row is BP, then it is the part number for the following rows until another BP is found
            ElseIf sheet648BSource.Cells(currentRow648B, 1).Value = "BP" Then
                partNumber = sheet648BSource.Cells(currentRow648B, 2)
                
            ' Logic to process the part quantity information
            Else
                ' Write the part number in the extractData sheet/tab
                .Cells(extractRow, 1) = partNumber
                
                ' Write the quantity in the extractData sheet/tab
                .Cells(extractRow, 2) = sheet648BSource.Cells(currentRow648B, 1)
                
                ' Write the date in the extractData sheet/tab. It might need formatted though
                ' If the date length is 6 digits, append 20 to the front then format as mm/dd/yyyy
                If Len(sheet648BSource.Cells(currentRow648B, 3)) = 6 Then
                    dateString = "20" & sheet648BSource.Cells(currentRow648B, 3)
                    .Cells(extractRow, 3).Value = DateSerial(Left(dateString, 4), Mid(dateString, 5, 2), Right(dateString, 2))
                    .Cells(extractRow, 3).NumberFormat = "mm/dd/yyyy"
                
                ' Assume date is 8 digits and can be directly formatted to mm/dd/yyyy
                Else
                    .Cells(extractRow, 3).Value = DateSerial(Left(sheet648BSource.Cells(currentRow648B, 3).Value, 4), Mid(sheet648BSource.Cells(currentRow648B, 3).Value, 5, 2), Right(sheet648BSource.Cells(currentRow648B, 3).Value, 2))
                    .Cells(extractRow, 3).NumberFormat = "mm/dd/yyyy"
                End If
                
                'Increment the row we are writing data to in extractData
                extractRow = extractRow + 1
            End If
        ' Increment the row we are reading in 648B sheet/tab
        Next currentRow648B
        
        ' Loop for the 677B sheet/tab. This will loop from row 1 until the last row of the sheet/tab
        For currentRow677B = 1 To lastRow677B
            ' If first column of the row is empty, skip it
            If sheet677BSource.Cells(currentRow677B, 1).Value = "" Then
                
            ' If first column of the row is BP, then it is the part number for the following rows until another BP is found
            ElseIf sheet677BSource.Cells(currentRow677B, 1).Value = "BP" Then
                partNumber = sheet677BSource.Cells(currentRow677B, 2)
            
            ' Logic to process the part quantity information
            Else
                ' Write the part number in the extractData sheet/tab
                .Cells(extractRow, 1) = partNumber
                
                ' Write the quantity in the extractData sheet/tab
                .Cells(extractRow, 2) = sheet677BSource.Cells(currentRow677B, 1)
                
                ' Write the date in the extractData sheet/tab. It might need formatted though
                ' If the date length is 6 digits, append 20 to the front then format as mm/dd/yyyy
                If Len(sheet677BSource.Cells(currentRow677B, 3)) = 6 Then
                    dateString = "20" & sheet677BSource.Cells(currentRow677B, 3)
                    .Cells(extractRow, 3).Value = DateSerial(Left(dateString, 4), Mid(dateString, 5, 2), Right(dateString, 2))
                    .Cells(extractRow, 3).NumberFormat = "mm/dd/yyyy"
                
                ' Assume date is 8 digits and can be directly formatted to mm/dd/yyyy
                Else
                    .Cells(extractRow, 3).Value = DateSerial(Left(sheet677BSource.Cells(currentRow677B, 3).Value, 4), Mid(sheet677BSource.Cells(currentRow677B, 3).Value, 5, 2), Right(sheet677BSource.Cells(currentRow677B, 3).Value, 2))
                    .Cells(extractRow, 3).NumberFormat = "mm/dd/yyyy"
                End If
                
                 'Increment the row we are writing data to in extractData
                extractRow = extractRow + 1
            End If
        ' Increment the row we are reading in 677B sheet/tab
        Next currentRow677B
        
        ' Loop for the 680A sheet/tab. This will loop from row 1 until the last row of the sheet/tab
        For currentRow680A = 1 To lastRow680A
            ' If first column of the row is empty, skip it
            If sheet680ASource.Cells(currentRow680A, 1).Value = "" Then
                
            ' If first column of the row is BP, then it is the part number for the following rows until another BP is found
            ElseIf sheet680ASource.Cells(currentRow680A, 1).Value = "BP" Then
                partNumber = sheet680ASource.Cells(currentRow680A, 2)
            
            ' Logic to process the part quantity information
            Else
                ' Write the part number in the extractData sheet/tab
                .Cells(extractRow, 1) = partNumber
                
                ' Write the quantity in the extractData sheet/tab
                .Cells(extractRow, 2) = sheet680ASource.Cells(currentRow680A, 1)
                
                ' Write the date in the extractData sheet/tab. It might need formatted though
                ' If the date length is 6 digits, append 20 to the front then format as mm/dd/yyyy
                If Len(sheet680ASource.Cells(currentRow680A, 3)) = 6 Then
                    dateString = "20" & sheet680ASource.Cells(currentRow680A, 3)
                    .Cells(extractRow, 3).Value = DateSerial(Left(dateString, 4), Mid(dateString, 5, 2), Right(dateString, 2))
                    .Cells(extractRow, 3).NumberFormat = "mm/dd/yyyy"
                
                ' Assume date is 8 digits and can be directly formatted to mm/dd/yyyy
                Else
                    .Cells(extractRow, 3).Value = DateSerial(Left(sheet680ASource.Cells(currentRow680A, 3).Value, 4), Mid(sheet680ASource.Cells(currentRow680A, 3).Value, 5, 2), Right(sheet680ASource.Cells(currentRow680A, 3).Value, 2))
                    .Cells(extractRow, 3).NumberFormat = "mm/dd/yyyy"
                End If
                
                'Increment the row we are writing data to in extractData
                extractRow = extractRow + 1
            End If
            
        ' Increment the row we are reading in 680A sheet/tab
        Next currentRow680A
        
        ' Loop for the 744A sheet/tab. This will loop from row 1 until the last row of the sheet/tab
        For currentRow744A = 1 To lastRow744A
            ' If first column of the row is empty, skip it
            If sheet744ASource.Cells(currentRow744A, 1).Value = "" Then
                
            ' If first column of the row is BP, then it is the part number for the following rows until another BP is found
            ElseIf sheet744ASource.Cells(currentRow744A, 1).Value = "BP" Then
                partNumber = sheet744ASource.Cells(currentRow744A, 2)
            
            ' Logic to process the part quantity information
            Else
                ' Write the part number in the extractData sheet/tab
                .Cells(extractRow, 1) = partNumber
                
                ' Write the quantity in the extractData sheet/tab
                .Cells(extractRow, 2) = sheet744ASource.Cells(currentRow744A, 1)
                
                ' Write the date in the extractData sheet/tab. It might need formatted though
                ' If the date length is 6 digits, append 20 to the front then format as mm/dd/yyyy
                If Len(sheet744ASource.Cells(currentRow744A, 3)) = 6 Then
                    dateString = "20" & sheet744ASource.Cells(currentRow744A, 3)
                    .Cells(extractRow, 3).Value = DateSerial(Left(dateString, 4), Mid(dateString, 5, 2), Right(dateString, 2))
                    .Cells(extractRow, 3).NumberFormat = "mm/dd/yyyy"
                
                ' Assume date is 8 digits and can be directly formatted to mm/dd/yyyy
                Else
                    .Cells(extractRow, 3).Value = DateSerial(Left(sheet744ASource.Cells(currentRow744A, 3).Value, 4), Mid(sheet744ASource.Cells(currentRow744A, 3).Value, 5, 2), Right(sheet744ASource.Cells(currentRow744A, 3).Value, 2))
                    .Cells(extractRow, 3).NumberFormat = "mm/dd/yyyy"
                End If
                
                'Increment the row we are writing data to in extractData
                extractRow = extractRow + 1
            End If
            
        ' Increment the row we are reading in 680A sheet/tab
        Next currentRow744A
    
    End With
    
    ' Close the external workbook
    wb.Close
End Sub

Sub FormatConsolidated()
    ' Set the sheets/tabs we will use
    Dim extractDataSource As Worksheet: Set extractDataSource = Worksheets("extractData")
    Dim tempSource As Worksheet: Set tempSource = Worksheets("temp")
    Dim consolidatedDestination As Worksheet: Set consolidatedDestination = Worksheets("Consolidated")
    
    ' Set the last row in the extractData sheet/tab so we know how far to loop
    Dim lastRowExtractData As Long: lastRowExtractData = extractDataSource.UsedRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    'Copy unique Part numbers from extractData sheet/tab to the temp sheet/tab
    extractDataSource.Range(extractDataSource.Cells(1, 1), extractDataSource.Cells(lastRowExtractData, 1)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("temp").Range("A:A"), Unique:=True
    tempSource.Range(tempSource.Cells(1, 1), tempSource.Cells(Rows.Count, 1).End(xlUp)).Sort Key1:=tempSource.Range("A1"), Order1:=xlAscending, Header:=xlYes
    
    'Copy unique dates from extractData sheet/tab to the temp sheet/tab
    extractDataSource.Range(extractDataSource.Cells(1, 3), extractDataSource.Cells(lastRowExtractData, 3)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("temp").Range("B:B"), Unique:=True
    tempSource.Range(tempSource.Cells(1, 2), tempSource.Cells(Rows.Count, 2).End(xlUp)).Sort Key1:=tempSource.Range("B1"), Order1:=xlAscending, Header:=xlYes
    
    ' set the last rows for unique part numbers and unique dates in the temp sheet/tab. This will help us determine the number of rows/columns in the consolidated sheet/tab
    Dim lastRowUniquePart As Long: lastRowUniquePart = tempSource.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lastRowUniqueDate As Long: lastRowUniqueDate = tempSource.Cells(Rows.Count, 2).End(xlUp).Row
    
    ' Loop through unique dates and add them to Consolidated sheet/tab across the columns.
    ' Loop starts at row 2 since there is a header row on the temp sheet/tab
    ' Can use same loop variable for temp row and consolidated column since both start a 2 and increase by 1 per loop
    For currentRowUniqueDate = 2 To lastRowUniqueDate
        consolidatedDestination.Cells(1, currentRowUniqueDate) = tempSource.Cells(currentRowUniqueDate, 2)
        consolidatedDestination.Cells(1, currentRowUniqueDate).NumberFormat = "mm/dd/yyyy"
    Next currentRowUniqueDate
    
    ' Loop through unique Part numbers and add them to Consolidated sheet/tab down the rows.
    ' Loop starts at row 2 since there is a header row on the temp sheet/tab
    ' Can use same loop variable for temp row and consolidated row since both start a 2 and increase by 1 per loop
    For currentRowUniquePart = 2 To lastRowUniquePart
        consolidatedDestination.Cells(currentRowUniquePart, 1) = tempSource.Cells(currentRowUniquePart, 1)
    Next currentRowUniquePart
    
    ' Loop through the extractData sheet/tab and populate quantities in the consolidated sheet/tab
    For currentRowExtractData = 2 To lastRowExtractData
        ' Find the row for the given Part Number in the Consolidated sheet/tab
        outputRow = consolidatedDestination.UsedRange.Find(extractDataSource.Cells(currentRowExtractData, 1), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        
        ' Find the column for the given Date in the Consolidated sheet/tab
        outputColumn = consolidatedDestination.UsedRange.Find(extractDataSource.Cells(currentRowExtractData, 3), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        
        ' Insert the quantity into the Consolidated sheet/tab at the row and column that we found
        consolidatedDestination.Cells(outputRow, outputColumn) = consolidatedDestination.Cells(outputRow, outputColumn) + extractDataSource.Cells(currentRowExtractData, 2)
    Next currentRowExtractData
    
End Sub

Sub colorConsolidatedInventory()
    ' Set the sheets/tabs we will use
    Dim inventorySheet As Worksheet: Set inventorySheet = Worksheets("Inventory")
    Dim consolidatedDestination As Worksheet: Set consolidatedDestination = Worksheets("Consolidated")
    
    ' Set the last row in the Inventory sheet/tab so we know how far to loop. Also we need to set the last row and column of the Consolidated sheet/tab so we know where to process
    Dim lastRowInventory As Long: lastRowInventory = inventorySheet.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lastRowConsolidated As Long: lastRowConsolidated = consolidatedDestination.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lastColumnConsolidated As Long: lastColumnConsolidated = consolidatedDestination.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Define a variable to keep track of our rolling inventory count
    Dim inventoryCount As Long
    
    ' Nested loop. We will loop through each row in the consolidated. Inside of processing each row we will loop through each column from left to right. Start at row 2 since there is a header row
    For currentRowConsolidated = 2 To lastRowConsolidated
    
        ' Search the inventory sheet/tab for the part number. set this variable equal to that part's quantity
        inventoryCount = inventorySheet.Cells(inventorySheet.UsedRange.Find(consolidatedDestination.Cells(currentRowConsolidated, 1), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row, 2)
        
        ' For the given row (part), loop through each column (date) from left to right. Start in column 2 since part numbers are in column 1
        For currentColumnConsolidated = 2 To lastColumnConsolidated
        
        ' If the cell value equals 0, clear the cell
        If consolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Value = 0 Then
            consolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Clear
        End If
        
        ' Only apply formatting if the cell has a number in it
        If consolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Value > 0 Then
            ' Subtract the current cell consolidated count from the remaining inventory. This keeps track of remaining inventory
            inventoryCount = inventoryCount - consolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated)
            If inventoryCount >= 0 Then
                ' Color green if there is enough inventory to cover this
                consolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Font.Color = RGB(34, 139, 34)
            Else
                ' Color red and set pattern if there is not enough inventory to cover this
                consolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Font.Color = vbRed
                consolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Interior.Pattern = xlPatternGray8
            End If
        End If
        ' Go to Next Column in Consolidated
        Next currentColumnConsolidated
    ' Go to Next Row in Consolidated
    Next currentRowConsolidated
End Sub

Sub GenerateInventory()
    
    ' Create the sheets we need in this workbook
    ThisWorkbook.Worksheets.Add.Name = "Inventory"
    ThisWorkbook.Worksheets.Add.Name = "New Parts"
    
    ' Open the Inventory workbook
    Dim wb As Workbook: Set wb = Workbooks.Open("R:\Operations\Pre-Production\Planning\Planning_Erica\EDI\Templates\Master Inventory.xlsx")
    
    ' Set the sheets/tabs we will use
    Dim inventorySheet As Worksheet: Set inventorySheet = wb.Worksheets("Master Inventory")
    Dim consolidatedInventorySheet As Worksheet: Set consolidatedInventorySheet = ThisWorkbook.Worksheets("Inventory")
    Dim newPartsSheet As Worksheet: Set newPartsSheet = ThisWorkbook.Worksheets("New Parts")
    Dim consolidatedSource As Worksheet: Set consolidatedSource = ThisWorkbook.Worksheets("Consolidated")
    
    ' Need to know last row of Total Consolidated sheet/tab to know what part numbers to search for. Need to know last row of Inventory tab for search funtion
    Dim lastRowTotalConsolidatedSource As Long: lastRowTotalConsolidatedSource = consolidatedSource.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lastRowInventorySource As Long: lastRowInventorySource = inventorySheet.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lastRowConsolidatedInventorySource As Long: lastRowConsolidatedInventorySource = consolidatedInventorySheet.Cells(Rows.Count, 1).End(xlUp).Row
    Dim currentRowNewParts As Long: currentRowNewParts = 1
    
    ' Variable to assign found part number rows as we loop
    Dim foundPart As Range
    
    ' Loop through Total Consolidated sheet/tab part numbers. start in row 2 to avoid header row
    For currentRowTotalConsolidated = 2 To lastRowTotalConsolidatedSource
        'Find row of Part Number in Inventory Tab. Returns Nothing if part number is not on the inventory
        Set foundPart = inventorySheet.UsedRange.Find(consolidatedSource.Cells(currentRowTotalConsolidated, 1), SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        
        ' If Part Number is not found we need to add it to the consolidated inventory sheet and new parts sheet
        If foundPart Is Nothing Then
            ' Write part number to inventory sheet/tab
            consolidatedInventorySheet.Cells(lastRowConsolidatedInventorySource, 1) = consolidatedSource.Cells(currentRowTotalConsolidated, 1)
            
            ' set part quantity to 0 in inventory sheet/tab
            consolidatedInventorySheet.Cells(lastRowConsolidatedInventorySource, 2) = 0
            
            ' add part number to new parts sheet/tab
            newPartsSheet.Cells(currentRowNewParts, 1) = consolidatedSource.Cells(currentRowTotalConsolidated, 1)
            
            ' Increment the variables since we have written to them
            lastRowConsolidatedInventorySource = lastRowConsolidatedInventorySource + 1
            currentRowNewParts = currentRowNewParts + 1
        ' If Part number is found, add it to the consolidated inventory sheet
        Else
            ' Write part number to inventory sheet/tab
            consolidatedInventorySheet.Cells(lastRowConsolidatedInventorySource, 1) = consolidatedSource.Cells(currentRowTotalConsolidated, 1)
            
            ' set part quantity to 0 in inventory sheet/tab
            consolidatedInventorySheet.Cells(lastRowConsolidatedInventorySource, 2) = inventorySheet.Cells(foundPart.Row, 2).Value
            
            ' Increment the variables since we have written to them
            lastRowConsolidatedInventorySource = lastRowConsolidatedInventorySource + 1
        End If
    Next currentRowTotalConsolidated
    
    wb.Close

End Sub
