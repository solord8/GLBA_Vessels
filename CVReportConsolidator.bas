Attribute VB_Name = "CVReportConsolidator"
'Daniel Solorzano-Jones
'solorzanodani@hotmail.com
'Last updated: 8/7/2025

Sub ConsolidateReports()
    On Error GoTo ErrorHandler
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim consolidatedWs As Worksheet
    Dim lastRow As Long
    Dim lastDataRow As Long
    Dim metadataValue As Variant ' To hold the metadata value
    Dim metadataColumn As Integer ' Column to add metadata in consolidated sheet

    ' SET THE FOLDER PATH WHERE THE REPORTS ARE LOCATED
    folderPath = "Q:\Administration A\A24 Committees\Backcountry\Charter Vessel\Reports\Charter2025\" ' THE FOLDER PATH MUST END IN "\"
    fileName = Dir(folderPath & "*.xlsx")
    
    ' Set the worksheet where consolidated data will be stored
    Set consolidatedWs = ThisWorkbook.Sheets("Raw Data") ' Change to your target sheet
    vesselName = consolidatedWs.Cells(1, consolidatedWs.Columns.Count).End(xlToLeft).Column + 1 ' Next column for metadata
    
    ' Define headers
    consolidatedWs.Cells(1, 1).Value = "Activity"
    consolidatedWs.Cells(1, 2).Value = "Date"
    consolidatedWs.Cells(1, 3).Value = "Start Time"
    consolidatedWs.Cells(1, 4).Value = "End Time"
    consolidatedWs.Cells(1, 5).Value = "Passengers"
    consolidatedWs.Cells(1, 6).Value = "Crew"
    consolidatedWs.Cells(1, 7).Value = "Kayaks"
    consolidatedWs.Cells(1, 8).Value = "Location"
    consolidatedWs.Cells(1, 9).Value = "Detail"
    consolidatedWs.Cells(1, 10).Value = "Comments"
    metadataColumn = 11 ' Change if a metadata column is needed or adjust accordingly
    
    ' Add header for the metadata in the consolidated worksheet
    consolidatedWs.Cells(1, vesselName).Value = "Vessel Name" ' Change header name if needed
    
    ' Loop through each Excel file in the folder
    Do While fileName <> ""
        Debug.Print "Processing file: " & fileName ' Debugging line
        Set wb = Workbooks.Open(folderPath & fileName)
        
        ' Check if the target sheet exists
        On Error Resume Next
        Set ws = wb.Sheets("GLBA Off-Vessel Report")
        On Error GoTo ErrorHandler
        
        If Not ws Is Nothing Then
            ' Extract the Vessel Name from a specific cell (D2)
            metadataValue = ws.Range("D2").Value ' Change the cell reference as needed
            
            ' Find the last row of data in the current worksheet
            lastDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' Copy raw data (starting from row 6)
            If lastDataRow >= 6 Then ' Ensure there is data to copy
                Dim i As Long
                Dim dataRange As Range

                ' Define the range to iterate through (columns A to J)
                Set dataRange = ws.Range("A6:J" & lastDataRow) ' Change columns as needed

                lastRow = consolidatedWs.Cells(consolidatedWs.Rows.Count, 1).End(xlUp).Row + 1
                
                ' Copy the data values without formatting to the consolidated worksheet
                For i = 1 To dataRange.Rows.Count
                    Dim j As Long
                    For j = 1 To dataRange.Columns.Count
                        consolidatedWs.Cells(lastRow + i - 1, j).Value = dataRange.Cells(i, j).Value
                    Next j
                Next i

                ' Fill the metadata value in the new column for all copied rows
                consolidatedWs.Range(consolidatedWs.Cells(lastRow, metadataColumn), _
                                      consolidatedWs.Cells(lastRow + dataRange.Rows.Count - 1, metadataColumn)).Value = metadataValue
            End If
            
            ' Format columns
            consolidatedWs.Columns(2).NumberFormat = "mm/dd/yy" ' Date format (adjust as needed)
            consolidatedWs.Columns(3).NumberFormat = "hh:mm" ' Start Time format
            consolidatedWs.Columns(4).NumberFormat = "hh:mm" ' End Time format
        Else
            Debug.Print "SHEET 'GLBA Off-Vessel Report' NOT FOUND IN FILE: " & fileName
        End If
        
        ' Close the current workbook without saving
        wb.Close False
        fileName = Dir
    Loop

    MsgBox "Consolidation complete!"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: (No GLBA Off-Vessel Report Found) " & Err.Description
    Resume Next
End Sub

