Attribute VB_Name = "TVDataProcessor"
'Daniel Solorzano-Jones
'solorzanodani@hotmail.com
'Last updated: 8/7/2025

Sub ProcessRawData()
    Dim rawDataWs As Worksheet
    Dim processedDataWs As Worksheet
    Dim lastRow As Long
    Dim finalRow As Long
    Dim i As Long
    
    ' Set the worksheets
    Set rawDataWs = ThisWorkbook.Sheets("Raw Data")
    Set processedDataWs = ThisWorkbook.Sheets("Processed Data")

    ' Clear the Final Table sheet before new data
    processedDataWs.Cells.Clear

    ' Define headers for the Final Table as specified
    processedDataWs.Cells(1, 1).Value = "Vessel"
    processedDataWs.Cells(1, 2).Value = "Type"
    processedDataWs.Cells(1, 3).Value = "Activity"
    processedDataWs.Cells(1, 4).Value = "GROUPS"
    processedDataWs.Cells(1, 5).Value = "PAX"
    processedDataWs.Cells(1, 6).Value = "CREW"
    processedDataWs.Cells(1, 7).Value = "TOTAL PEOPLE"
    processedDataWs.Cells(1, 8).Value = "Location of Activity"
    processedDataWs.Cells(1, 9).Value = "Location Detail"
    processedDataWs.Cells(1, 10).Value = "Wilderness"
    processedDataWs.Cells(1, 11).Value = "Date"
    processedDataWs.Cells(1, 12).Value = "START TIME COR"
    processedDataWs.Cells(1, 13).Value = "END TIME COR"
    processedDataWs.Cells(1, 14).Value = "LOCATION STANDARDIZED"
    processedDataWs.Cells(1, 15).Value = "ACTIVITY STANDARDIZED"
    processedDataWs.Cells(1, 16).Value = "YEAR"
    processedDataWs.Cells(1, 17).Value = "Comments"
    
    finalRow = 2 ' Start writing data from the second row
    
    ' List of valid locations
    validLocations = Array("Other", "Bartlett Cove", "Bear Track", "Dundas", "Excursion", _
                           "Fern Harbor", "Geikie", "Gloomy", "Hugh Miller", _
                           "Jaw Point", "Johns Hopkins", "Lamplugh", "Reid", _
                           "Russel Cut", "Sandy", "Tidal")
                           
    'List of valid activities
    validActivities = Array("Kayak", "Skiff", "Hike")
    
    'List of Wilderness activities
    wildernessActivities = Array("Hike", "Skiff/Hike", "Kayak/Hike")
    
    'List of Wilderness Waters
    wildernessWaters = Array("Rendu", "Hugh Miller Inlet", "Adams", "Beardslee", "Scidmore")
                           
    ' Get the last row of data in Raw Data
    lastRow = rawDataWs.Cells(rawDataWs.Rows.Count, 1).End(xlUp).Row

    ' Loop through RawData and transform data
    For i = 2 To lastRow ' Assuming headers are in row 1
        Dim vesselName As Variant
        Dim typeOfActivity As String
        Dim passengers As Long
        Dim crew As Long
        Dim activityDate As Variant
        Dim location As String
        Dim detail As String
        Dim locationStandard As String
        Dim comments As String
        
        ' Read values from RawData (adjust column numbers as necessary)
        vesselName = rawDataWs.Cells(i, "K").Value ' Adjust column for Vessel Name as necessary
        typeOfActivity = rawDataWs.Cells(i, "A").Value ' Type of Activity
        passengers = 0 ' Passengers
        crew = 0 ' Crew
        
        ' Check if the value is numeric before assigning
        If IsNumeric(rawDataWs.Cells(i, "E").Value) Then
            passengers = CLng(rawDataWs.Cells(i, "E").Value) ' Passengers
        End If

        If IsNumeric(rawDataWs.Cells(i, "F").Value) Then
            crew = CLng(rawDataWs.Cells(i, "F").Value) ' Crew
        End If

        activityDate = rawDataWs.Cells(i, "B").Value ' Date (assumed in column B)
        location = rawDataWs.Cells(i, "H").Value  ' Location
        detail = rawDataWs.Cells(i, "I").Value ' Detailed Location
        comments = rawDataWs.Cells(i, "J").Value  ' Comments
        wilderness = "" ' Wilderness
        
       
        ' Determine the Location Standard based on the condition
        If location = "Other" Then
            If InStr(detail, " G") > 0 Then
                locationStandard = Trim(Left(detail, InStr(detail, " G") - 1)) ' Take text before the G of "Glacier"
            ElseIf InStr(detail, ",") > 0 Then
                locationStandard = Trim(Left(detail, InStr(detail, ",") - 1)) ' Take text before the comma
            Else
                locationStandard = Trim(detail) ' Just take the detail if no special cases apply
            End If
        Else
            If InStr(location, " G") > 0 Then
                locationStandard = Trim(Left(location, InStr(location, " G") - 1)) ' Take text before the G of "Glacier"
            ElseIf InStr(location, ",") > 0 Then
                locationStandard = Trim(Left(location, InStr(location, ",") - 1)) ' Take text before the comma
            Else
                locationStandard = Trim(location) ' Otherwise, use the Location as is, trim any whitespace
            End If
        End If
        
        ' Determine if activity is in Wilderness
        If InStr(locationStandard, "Bartlett Cove") > 0 Then
            wilderness = "No"
        ElseIf InStr(typeOfActivity, "Hike") > 0 Then
            wilderness = "Yes"
        Else
           ' Check if locationStandard contains any string from wildernessWaters
            Dim j As Long
            wilderness = "No" ' Default to No
            For j = LBound(wildernessWaters) To UBound(wildernessWaters)
                If InStr(locationStandard, wildernessWaters(j)) > 0 Then
                    wilderness = "Yes"
                    Exit For
                End If
            Next j
        End If
              
        
        ' Check if activityDate is a valid date
        If Not IsDate(activityDate) Then
            ' Highlight row for review if date is invalid
            processedDataWs.Cells(finalRow, 1).Value = vesselName
            processedDataWs.Cells(finalRow, 2).Value = "TV" ' Type
            processedDataWs.Cells(finalRow, 3).Value = typeOfActivity
            processedDataWs.Cells(finalRow, 4).Value = "" ' Groups (empty column)
            processedDataWs.Cells(finalRow, 5).Value = passengers
            processedDataWs.Cells(finalRow, 6).Value = crew
            processedDataWs.Cells(finalRow, 7).Value = passengers + crew ' Total People
            processedDataWs.Cells(finalRow, 8).Value = location ' Location
            processedDataWs.Cells(finalRow, 9).Value = detail ' Detail
            processedDataWs.Cells(finalRow, 10).Value = "" ' Wilderness (empty column)
            processedDataWs.Cells(finalRow, 11).Value = "INVALID DATE" ' Highlight invalid date
            processedDataWs.Cells(finalRow, 12).Value = rawDataWs.Cells(i, "C").Value ' Start Time
            processedDataWs.Cells(finalRow, 13).Value = rawDataWs.Cells(i, "D").Value ' End Time
            processedDataWs.Cells(finalRow, 14).Value = locationStandard ' Location Standard
            processedDataWs.Cells(finalRow, 15).Value = typeOfActivity ' Activity
            processedDataWs.Cells(finalRow, 16).Value = "" ' Year (empty)
            processedDataWs.Cells(finalRow, 17).Value = comments ' Comments
            
            ' Format the row with a highlight color for review
            processedDataWs.Rows(finalRow).Interior.Color = RGB(255, 200, 200) ' Light red
            
            finalRow = finalRow + 1 ' Move to next row in Final Table for invalid dates
            GoTo NextRecord ' Skip to next record
        End If
        
        ' Write to Final Table
        processedDataWs.Cells(finalRow, 1).Value = vesselName ' Vessel Name
        processedDataWs.Cells(finalRow, 2).Value = "TV" ' Type
        processedDataWs.Cells(finalRow, 3).Value = typeOfActivity ' Type of Activity
        processedDataWs.Cells(finalRow, 4).Value = "" ' Groups (empty column)
        processedDataWs.Cells(finalRow, 5).Value = passengers ' Passengers
        processedDataWs.Cells(finalRow, 6).Value = crew ' Crew
        processedDataWs.Cells(finalRow, 7).Value = passengers + crew ' Total People
        processedDataWs.Cells(finalRow, 8).Value = location ' Location
        processedDataWs.Cells(finalRow, 9).Value = detail ' Detail
        processedDataWs.Cells(finalRow, 10).Value = wilderness ' Wilderness (empty column)
        processedDataWs.Cells(finalRow, 11).Value = rawDataWs.Cells(i, "B").Value ' Date
        processedDataWs.Cells(finalRow, 12).Value = rawDataWs.Cells(i, "C").Value ' Start Time
        processedDataWs.Cells(finalRow, 13).Value = rawDataWs.Cells(i, "D").Value ' End Time
        processedDataWs.Cells(finalRow, 14).Value = locationStandard ' Location Standard
        processedDataWs.Cells(finalRow, 15).Value = typeOfActivity ' Activity (assuming this is in Column A)
        processedDataWs.Cells(finalRow, 16).Value = Year(rawDataWs.Cells(i, "B").Value) ' Extract Year from Date
        processedDataWs.Cells(finalRow, 17).Value = comments ' Comments
        
        ' Highlight rows with non-standard Activities
        If IsError(Application.Match(typeOfActivity, validActivities, 0)) Then
            processedDataWs.Rows(finalRow).Interior.Color = RGB(255, 255, 0) ' Yellow highlight
        End If
         
        ' Highlight rows with invalid Location Standard
        If IsError(Application.Match(locationStandard, validLocations, 0)) Then
            processedDataWs.Rows(finalRow).Interior.Color = RGB(255, 165, 0) ' Orange highlight
        End If
              
        
        ' Format Date and Time Columns
        processedDataWs.Cells(finalRow, 11).NumberFormat = "mm/dd/yyyy" ' Format Date
        processedDataWs.Cells(finalRow, 12).NumberFormat = "hh:mm:ss" ' Format Start Time
        processedDataWs.Cells(finalRow, 13).NumberFormat = "hh:mm:ss" ' Format End Time

        finalRow = finalRow + 1 ' Move to next row in Final Data
        
NextRecord:
    Next i
    
    MsgBox "Data transformation to 'Final Table' complete!"
End Sub
