Sub TotalConsolidatedDataExtract()

    ' Change File Location to match where your data is
    ' List of workbooks we will loop through
    Dim CustomerDemandWorkbook As String
    CustomerDemandWorkbook = "C:\HUB\Documents\personal\EC\Customer Demand.xlsx"
    
    ' Create the sheets we need in this workbook
    ThisWorkbook.Worksheets.Add.Name = "extractData"
    ThisWorkbook.Worksheets.Add.Name = "temp"
    ThisWorkbook.Worksheets.Add.Name = "Total Consolidated"
    
    ' Perform the data extraction from the Customer Demand Workbook
    ExtractCustomerDemandData CustomerDemandWorkbook
    
    FormatTotalConsolidated
    
End Sub

Sub ExtractCustomerDemandData(workbookFilePath As String)

    ' Define the tab names
    Dim mxLine1 As String: mxLine1 = "MxLine1"
    Dim mxLine2 As String: mxLine2 = "MxLine2"
    Dim Rollins As String: Rollins = "Rollins"
    Dim SAG As String: SAG = "SAG"
    Dim Service As String: Service = "Service"
    Dim Tulsa As String: Tulsa = "Tulsa"
    Dim OptimasOthers As String: OptimasOthers = "Optimas & Others"
    
    ' Extract mxline1
    ExtractFromTemplatedTab workbookFilePath, mxLine1
    ExtractFromTemplatedTab workbookFilePath, mxLine2
    ExtractFromTemplatedTab workbookFilePath, Rollins
    ExtractFromTemplatedTab workbookFilePath, SAG
    ExtractFromTemplatedTab workbookFilePath, Service
    ExtractFromTemplatedTab workbookFilePath, Tulsa
    ExtractFromOptimasOthersTab workbookFilePath, OptimasOthers
    
End Sub

Sub ExtractFromTemplatedTab(workbookFilePath As String, tabName As String)

    ' Open the external workbook
    Dim wb As Workbook: Set wb = Workbooks.Open(workbookFilePath)
    
    ' Open the sheet/tab that we are extracting from
    Dim templatedSourceTab As Worksheet: Set templatedSourceTab = wb.Worksheets(tabName)
    
    ' Open the sheet/tab we are loading into
    Dim extractDataTab As Worksheet: Set extractDataTab = ThisWorkbook.Worksheets("extractData")
    
    ' Define variables for the last row and column of the templatedSource
    Dim lastRowTemplatedSourceTab As Long
    Dim lastColumnTemplatedSourceTab As Long
    
    ' Set last row and column for templated source tab. Only do this if the sheet/tab is not blank
    If WorksheetFunction.CountA(templatedSourceTab.UsedRange) > 0 Then
        lastRowTemplatedSourceTab = templatedSourceTab.Cells(Rows.Count, 1).End(xlUp).Row
        lastColumnTemplatedSourceTab = templatedSourceTab.Cells(1, Columns.Count).End(xlToLeft).Column
    End If
    
    ' Set the last row in the extractData sheet/tab so we know how far to loop
    Dim lastRowExtractDataTab As Long: lastRowExtractDataTab = extractDataTab.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    ' Create a variable that will eventually be an array to store the column numbers for columns that have a date. this will be useful for looping later
    Dim dateColumnArray
    
    ' Loop through the templated source tab finding all of the columns with a date in row 1. add those column numbers to the array. start in column 2 since part numbers are in column 1
    For currentTemplatedSourceTabColumn = 2 To lastColumnTemplatedSourceTab
        'Check to see if the cell value is a date
        If IsDate(templatedSourceTab.Cells(1, currentTemplatedSourceTabColumn)) Then
            ' If array has been created, redim to add a value
            If IsArray(dateColumnArray) Then
                ReDim Preserve dateColumnArray(UBound(dateColumnArray) + 1)
            ' Else create the array if this is the first entry
            Else
                ReDim dateColumnArray(0) As Long
            End If
            '  Add the column number to the array
            dateColumnArray(UBound(dateColumnArray)) = currentTemplatedSourceTabColumn
        End If
    Next currentTemplatedSourceTabColumn
    
    ' extractDataTab format: partNumber, quantity, date
    ' Populate the details on the extractData sheet/tab
    For currentTemplatedSourceTabRow = 2 To lastRowTemplatedSourceTab
    
        'Loop through the date columns for the given row to extract the data
        For Each dateColumn In dateColumnArray
            ' Only add data to extractData sheet/tab if there is a quantity
            If templatedSourceTab.Cells(currentTemplatedSourceTabRow, dateColumn) > 0 Then
                ' Populate Part Number
                extractDataTab.Cells(lastRowExtractDataTab, 1) = templatedSourceTab.Cells(currentTemplatedSourceTabRow, 1)
                
                'Populate Quantity
                extractDataTab.Cells(lastRowExtractDataTab, 2) = templatedSourceTab.Cells(currentTemplatedSourceTabRow, dateColumn)
                
                'Populate Date and format it to mm/dd/yyyy
                extractDataTab.Cells(lastRowExtractDataTab, 3) = templatedSourceTab.Cells(1, dateColumn)
                extractDataTab.Cells(lastRowExtractDataTab, 3).NumberFormat = "mm/dd/yyyy"
                
                ' Increment lastRowExtractDataTab since we populated the row we are in
                lastRowExtractDataTab = lastRowExtractDataTab + 1
            End If
        
        Next dateColumn
        
    Next currentTemplatedSourceTabRow
    
    wb.Close
    
End Sub

Sub ExtractFromOptimasOthersTab(workbookFilePath As String, tabName As String)

    ' Open the external workbook
    Dim wb As Workbook: Set wb = Workbooks.Open(workbookFilePath)
    
    ' Open the sheet/tab that we are extracting from
    Dim optimasSourceTab As Worksheet: Set optimasSourceTab = wb.Worksheets(tabName)
    
    ' Open the sheet/tab we are loading into
    Dim extractDataTab As Worksheet: Set extractDataTab = ThisWorkbook.Worksheets("extractData")
    
    ' Define variables for the last row and column of the templatedSource
    Dim lastRowOptimasSourceTab As Long
    Dim lastColumnOptimasSourceTab As Long
    
    ' Set last row and column for templated source tab. Only do this if the sheet/tab is not blank
    If WorksheetFunction.CountA(optimasSourceTab.UsedRange) > 0 Then
        lastRowOptimasSourceTab = optimasSourceTab.Cells(Rows.Count, 1).End(xlUp).Row
        lastColumnOptimasSourceTab = optimasSourceTab.Cells(1, Columns.Count).End(xlToLeft).Column
    End If
    
    ' Set the last row in the extractData sheet/tab so we know how far to loop
    Dim lastRowExtractDataTab As Long: lastRowExtractDataTab = extractDataTab.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    ' Create a variable that will eventually be an array to store the column numbers for columns that have a date. this will be useful for looping later
    Dim dateColumnArray
    
    ' Loop through the templated source tab finding all of the columns with a date in row 1. add those column numbers to the array. start in column 2 since part numbers are in column 1
    For currentOptimasSourceTabColumn = 2 To lastColumnOptimasSourceTab
        'Check to see if the cell value is a date
        If IsDate(optimasSourceTab.Cells(1, currentOptimasSourceTabColumn)) Then
            ' If array has been created, redim to add a value
            If IsArray(dateColumnArray) Then
                ReDim Preserve dateColumnArray(UBound(dateColumnArray) + 1)
            ' Else create the array if this is the first entry
            Else
                ReDim dateColumnArray(0) As Long
            End If
            '  Add the column number to the array
            dateColumnArray(UBound(dateColumnArray)) = currentOptimasSourceTabColumn
        End If
    Next currentOptimasSourceTabColumn
    
    ' extractDataTab format: partNumber, quantity, date
    ' Populate the details on the extractData sheet/tab
    For currentOptimasSourceTabRow = 2 To lastRowOptimasSourceTab
    
        ' If Part Number is found for the row, then extract data
        If optimasSourceTab.Cells(currentOptimasSourceTabRow, 2).Value <> "" Then
            'Loop through the date columns for the given row to extract the data
            For Each dateColumn In dateColumnArray
                ' Only add data to extractData sheet/tab if there is a quantity
                If optimasSourceTab.Cells(currentOptimasSourceTabRow, dateColumn) > 0 Then
                    ' Populate Part Number
                    extractDataTab.Cells(lastRowExtractDataTab, 1) = optimasSourceTab.Cells(currentOptimasSourceTabRow, 2)
                    
                    'Populate Quantity
                    extractDataTab.Cells(lastRowExtractDataTab, 2) = optimasSourceTab.Cells(currentOptimasSourceTabRow, dateColumn)
                    
                    'Populate Date and format it to mm/dd/yyyy
                    extractDataTab.Cells(lastRowExtractDataTab, 3) = optimasSourceTab.Cells(1, dateColumn)
                    extractDataTab.Cells(lastRowExtractDataTab, 3).NumberFormat = "mm/dd/yyyy"
                    
                    ' Increment lastRowExtractDataTab since we populated the row we are in
                    lastRowExtractDataTab = lastRowExtractDataTab + 1
                End If
            Next dateColumn
        End If
        
    Next currentOptimasSourceTabRow
    
    wb.Close
    
End Sub

Sub FormatTotalConsolidated()
    ' Set the sheets/tabs we will use
    Dim extractDataSource As Worksheet: Set extractDataSource = Worksheets("extractData")
    Dim tempSource As Worksheet: Set tempSource = Worksheets("temp")
    Dim consolidatedDestination As Worksheet: Set consolidatedDestination = Worksheets("Total Consolidated")
    
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
        outputRow = consolidatedDestination.Columns(1).Find(extractDataSource.Cells(currentRowExtractData, 1), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        
        ' Find the column for the given Date in the Consolidated sheet/tab
        outputColumn = consolidatedDestination.Rows(1).Find(extractDataSource.Cells(currentRowExtractData, 3), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        
        ' Insert the quantity into the Consolidated sheet/tab at the row and column that we found
        consolidatedDestination.Cells(outputRow, outputColumn) = consolidatedDestination.Cells(outputRow, outputColumn) + extractDataSource.Cells(currentRowExtractData, 2)
    Next currentRowExtractData
    
End Sub

Sub colorTotalConsolidatedInventory()
    ' Set the sheets/tabs we will use
    Dim inventorySheet As Worksheet: Set inventorySheet = Worksheets("Inventory")
    Dim totalConsolidatedDestination As Worksheet: Set totalConsolidatedDestination = Worksheets("Total Consolidated")
    
    ' Set the last row in the Inventory sheet/tab so we know how far to loop. Also we need to set the last row and column of the Consolidated sheet/tab so we know where to process
    Dim lastRowInventory As Long: lastRowInventory = inventorySheet.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lastRowConsolidated As Long: lastRowConsolidated = totalConsolidatedDestination.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lastColumnConsolidated As Long: lastColumnConsolidated = totalConsolidatedDestination.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Define a variable to keep track of our rolling inventory count
    Dim inventoryCount As Long
    
    ' Nested loop. We will loop through each row in the consolidated. Inside of processing each row we will loop through each column from left to right. Start at row 2 since there is a header row
    For currentRowConsolidated = 2 To lastRowConsolidated
    
        ' Search the inventory sheet/tab for the part number. set this variable equal to that part's quantity
        inventoryCount = inventorySheet.Cells(inventorySheet.UsedRange.Find(totalConsolidatedDestination.Cells(currentRowConsolidated, 1), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row, 2)
        
        ' For the given row (part), loop through each column (date) from left to right. Start in column 2 since part numbers are in column 1
        For currentColumnConsolidated = 2 To lastColumnConsolidated
        
        ' If the cell value equals 0, clear the cell
        If totalConsolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Value = 0 Then
            totalConsolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Clear
        End If
        
        ' Only apply formatting if the cell has a number in it
        If totalConsolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Value > 0 Then
            ' Subtract the current cell consolidated count from the remaining inventory. This keeps track of remaining inventory
            inventoryCount = inventoryCount - totalConsolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated)
            If inventoryCount >= 0 Then
                ' Color green if there is enough inventory to cover this
                totalConsolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Font.Color = RGB(34, 139, 34)
            Else
                ' Color red and set pattern if there is not enough inventory to cover this
                totalConsolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Font.Color = vbRed
                totalConsolidatedDestination.Cells(currentRowConsolidated, currentColumnConsolidated).Interior.Pattern = xlPatternGray8
            End If
        End If
        ' Go to Next Column in Consolidated
        Next currentColumnConsolidated
    ' Go to Next Row in Consolidated
    Next currentRowConsolidated
End Sub

Sub TotalConsolidatedNewParts()
    
    ' Create the sheets we need in this workbook
    ThisWorkbook.Worksheets.Add.Name = "New Parts"
    
    ' Set the sheets/tabs we will use
    Dim inventorySheet As Worksheet: Set inventorySheet = Worksheets("Inventory")
    Dim newPartsSheet As Worksheet: Set newPartsSheet = Worksheets("New Parts")
    Dim totalConsolidatedSource As Worksheet: Set totalConsolidatedSource = Worksheets("Total Consolidated")
    
    ' Need to know last row of Total Consolidated sheet/tab to know what part numbers to search for. Need to know last row of Inventory tab for search funtion
    Dim lastRowTotalConsolidatedSource As Long: lastRowTotalConsolidatedSource = totalConsolidatedSource.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lastRowInventorySource As Long: lastRowInventorySource = inventorySheet.Cells(Rows.Count, 1).End(xlUp).Row
    Dim currentRowNewParts As Long: currentRowNewParts = 1
    
    ' Variable to assign found part number rows as we loop
    Dim foundPart As Range
    
    ' Loop through Total Consolidated sheet/tab part numbers. start in row 2 to avoid header row
    For currentRowTotalConsolidated = 2 To lastRowTotalConsolidatedSource
        'Find row of Part Number in Inventory Tab. Returns Nothing if part number is not on the inventory
        Set foundPart = inventorySheet.UsedRange.Find(totalConsolidatedSource.Cells(currentRowTotalConsolidated, 1), SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        
        ' If Part Number is not found we need to add it to the inventory sheet and new parts sheet
        If foundPart Is Nothing Then
            ' Write part number to inventory sheet/tab
            inventorySheet.Cells(lastRowInventorySource + 1, 1) = totalConsolidatedSource.Cells(currentRowTotalConsolidated, 1)
            
            ' set part quantity to 0 in inventory sheet/tab
            inventorySheet.Cells(lastRowInventorySource + 1, 2) = 0
            
            ' add part number to new parts sheet/tab
            newPartsSheet.Cells(currentRowNewParts, 1) = totalConsolidatedSource.Cells(currentRowTotalConsolidated, 1)
            
            ' Increment the variables since we have written to them
            lastRowInventorySource = lastRowInventorySource + 1
            currentRowNewParts = currentRowNewParts + 1
        End If
    Next currentRowTotalConsolidated

End Sub
