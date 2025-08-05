Attribute VB_Name = "CVDataProcessor"
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
    processedDataWs.Cells(1, 3).Value = "Type of Activity"
    processedDataWs.Cells(1, 4).Value = "Groups"
    processedDataWs.Cells(1, 5).Value = "Passengers"
    processedDataWs.Cells(1, 6).Value = "Crew"
    processedDataWs.Cells(1, 7).Value = "Total People"
    processedDataWs.Cells(1, 8).Value = "Location"
    processedDataWs.Cells(1, 9).Value = "Detail"
    processedDataWs.Cells(1, 10).Value = "Wilderness"
    processedDataWs.Cells(1, 11).Value = "Date"
    processedDataWs.Cells(1, 12).Value = "Start Time"
    processedDataWs.Cells(1, 13).Value = "End Time"
    processedDataWs.Cells(1, 14).Value = "Location Standard"
    processedDataWs.Cells(1, 15).Value = "Activity Standardized"
    processedDataWs.Cells(1, 16).Value = "Year"
    processedDataWs.Cells(1, 17).Value = "Comments"
    
    finalRow = 2 ' Start writing data from the second row
    
    ' List of valid locations
    validLocations = Array("Bartlett Cove", "Bear Track", "Blue Mouse Cove", _
                           "Dundas", "East Arm", "Excursion", "Fern Harbor", _
                           "Fingers", "Geikie", "Gloomy", "Hugh Miller", "Johns Hopkins", _
                           "Lamplugh", "Lower Bay", "McBride", "Other", _
                           "Queen", "Reid", "Russel Cut", "Sandy", "Tarr", _
                           "Tidal", "Upper Muir")
                           
    'List of valid activities
    validActivities = Array("Kayak", "Skiff", "Hike")
                           
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
        passengers = rawDataWs.Cells(i, "E").Value ' Passengers
        crew = rawDataWs.Cells(i, "F").Value ' Crew
        activityDate = rawDataWs.Cells(i, "B").Value ' Date (assumed in column B)
        location = rawDataWs.Cells(i, "H").Value  ' Location
        detail = rawDataWs.Cells(i, "I").Value ' Detailed Location
        comments = rawDataWs.Cells(i, "J").Value  ' Comments
        
        
       
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
        processedDataWs.Cells(finalRow, 10).Value = "" ' Wilderness (empty column)
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


