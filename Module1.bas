Attribute VB_Name = "Module1"
Public Sub populatepenman() 'sub button for populating reference et with penman monteith calculations
    ActiveWorkbook.Sheets("Reference_ET_PM").Select 'select reference et with penman monteith calculations worksheet
    Range("a8:zz9999").ClearContents 'clear contents in reference et with penman monteith calculations sheets from a8 to zz9999
    Application.Wait (Now + TimeValue("0:00:01"))
    Application.ScreenUpdating = False
    Sheets("Reference_ET_PM").Range("6:6").ClearComments
    Sheets("temp3").Visible = True 'makes sheet1 for temp visible
    'get the range for input data
    y1 = Range("c1").Value 'starting year for data collection
    y2 = Range("c2").Value 'ending year for data collection
    'call for a function to populate the dates in range input
    Call createDate(y1, y2)
    'loop through each weather station column for AZMET
    ithCol = 1
    colCount = Range("5:5").Cells.SpecialCells(xlCellTypeConstants).Count 'counts the amount of entries in row 5
    Range("B6").Select 'selects station number for pulling data
    Dim azmetRefenceETPen As Collection 'create collections variable type to hold data
    While ActiveCell.Value <> "" 'while cell is active
    'select station number to process
         stationNum = ActiveCell.Value
    ' Set Excel status bar text
        progressSymbols = 20
        progressComp = Round(progressSymbols * ithCol / colCount, 0)
        Application.StatusBar = "Processing " & stationNum & " | " & Format(ithCol, "000") & " of " & _
            Format(colCount, "000") & " | " & String(progressComp, ChrW(&H25CF)) & _
            String(progressSymbols - progressComp, ChrW(&H25CC)) & " " & _
            Format(100 * ithCol / colCount, "00") & "% Complete"
        
        'populate sheet1 tab with azmet precip weather data
        Set azmetRefenceETPen = azmetDataPen(stationNum, y1, y2)
        'call for function to populate the dates
        'move to the next column for next weather station to pull data
        Call getAzmetDataPen(azmetRefenceETPen)
        Cells(6, ActiveCell.Column + 1).Select
        ithCol = ithCol + 1 ' update while loop
    Wend
    Application.StatusBar = "Done!" 'display message of completion after printing on status bar
    Sheets("temp3").Visible = False 'hides temp3 sheet holding values from data pull
    Application.ScreenUpdating = True
    Call delete_Connections 'deletes all connections
End Sub

Public Sub populateprecip() 'sub button for populating precipitation tab
    ActiveWorkbook.Sheets("Precipitation").Select 'select precipitation worksheet
    Range("a8:zz9999").ClearContents 'clear contents in precipitation sheets from a8 to zz9999
    Application.Wait (Now + TimeValue("0:00:01"))
    Application.ScreenUpdating = False
    Sheets("precipitation").Range("6:6").ClearComments
    Sheets("temp").Visible = True 'makes sheet1 for temp visible
    'get the range for input data
    y1 = Range("c1").Value 'starting year for data collection
    y2 = Range("c2").Value 'ending year for data collection
    'call for a function to populate the dates in range input
    Call createDate(y1, y2)
    'loop through each weather station column for AZMET
    ithCol = 1
    colCount = Range("5:5").Cells.SpecialCells(xlCellTypeConstants).Count 'counts the amount of entries in row 5
    Range("B6").Select 'selects station number for pulling data
    Dim azmetPrecipArray As Collection 'create collections variable type to hold data
    While ActiveCell.Value <> "" 'while cell is active
    'select station number to process
         stationNum = ActiveCell.Value
    ' Set Excel status bar text
        progressSymbols = 20
        progressComp = Round(progressSymbols * ithCol / colCount, 0)
        Application.StatusBar = "Processing " & stationNum & " | " & Format(ithCol, "000") & " of " & _
            Format(colCount, "000") & " | " & String(progressComp, ChrW(&H25CF)) & _
            String(progressSymbols - progressComp, ChrW(&H25CC)) & " " & _
            Format(100 * ithCol / colCount, "00") & "% Complete"
        
        'populate sheet1 tab with azmet precip weather data
        Set azmetPrecipArray = azmetData(stationNum, y1, y2)
        'call for function to populate the dates
        'move to the next column for next weather station to pull data
        Call getAzmetData(azmetPrecipArray)
        Cells(6, ActiveCell.Column + 1).Select
        ithCol = ithCol + 1 ' update while loop
    Wend
    Application.StatusBar = "Done!"
    Call delete_Connections
    Application.ScreenUpdating = True
End Sub
Public Sub populateRefET() 'sub button for populating precipitation tab
    ActiveWorkbook.Sheets("reference_et").Select 'select precipitation worksheet
    Range("a8:zz9999").ClearContents 'clear contents in precipitation sheets from a8 to zz9999
    Application.Wait (Now + TimeValue("0:00:01"))
    Application.ScreenUpdating = False
    Sheets("reference_et").Range("6:6").ClearComments
    Sheets("temp2").Visible = True 'makes sheet1 for temp visible

    'get the range for input data
    y1 = Range("c1").Value 'starting year for data collection
    y2 = Range("c2").Value 'ending year for data collection
    'call for a function to populate the dates in range input
    Call createDate(y1, y2)
    'loop through each weather station column for AZMET
    ithCol = 1
    colCount = Range("5:5").Cells.SpecialCells(xlCellTypeConstants).Count 'counts the amount of entries in row 5
    Range("B6").Select 'selects station number for pulling data
    Dim azmetRefETArray As Collection 'create collections variable type to hold data
    While ActiveCell.Value <> "" 'while cell is active
    'select station number to process
         stationNum = ActiveCell.Value
    ' Set Excel status bar text
        progressSymbols = 20
        progressComp = Round(progressSymbols * ithCol / colCount, 0)
        Application.StatusBar = "Processing " & stationNum & " | " & Format(ithCol, "000") & " of " & _
            Format(colCount, "000") & " | " & String(progressComp, ChrW(&H25CF)) & _
            String(progressSymbols - progressComp, ChrW(&H25CC)) & " " & _
            Format(100 * ithCol / colCount, "00") & "% Complete"
        
        'populate sheet1 tab with azmet precip weather data
        Set azmetRefETArray = azmetDataR(stationNum, y1, y2)
        'call for function to populate the dates
        'move to the next column for next weather station to pull data
        Call getAzmetDataR(azmetRefETArray)
        Cells(6, ActiveCell.Column + 1).Select
        ithCol = ithCol + 1 ' update while loop
    Wend
    Application.StatusBar = "Done!"
    Sheets("temp2").Visible = False
    Call delete_Connections
    Application.ScreenUpdating = True
End Sub

'function to grab precip data from azmet websites
Public Function azmetData(stationNum, y1, y2) As Collection
    'build array for numbers
    Dim collData As Collection
    Set collData = New Collection
    'Call delete_Connections
    'clear contents for temp tab
    ActiveWorkbook.Sheets("temp").Select
    Range("a1:zz9999").ClearContents
    Range("a1").Select
    test = y1
    While (test <= y2)
        'URL Pattern
        'https://cals.arizona.edu/azmet/data/2017rd.txt
        urlString = "text;https://cals.arizona.edu/azmet/data/" & stationNum & test & "rd.txt"
        'ActiveSheet.QueryTables.Item(1).Delete
        ' Get USGS Data
        With ActiveSheet.QueryTables.Add(Connection:=urlString, Destination:=Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .Refresh
        End With
        Range("a1").Select
            While ActiveCell.Value <> ""
            collData.Add ActiveCell.Value & "," & ActiveCell.Offset(0, 1) & "," & ActiveCell.Offset(0, 11) 'adds one row into the data array
            ActiveCell.Offset(1, 0).Select 'jumps to the next row
            Wend
    test = test + 1
    Wend
    Set azmetData = collData 'set function return
       
End Function
'function to grab data from azmet websites
Public Function azmetDataR(stationNum, y1, y2) As Collection
    'build array for numbers
    Dim collDataR As Collection
    Set collDataR = New Collection
    'Call delete_Connections
    'clear contents for temp tab
    ActiveWorkbook.Sheets("temp2").Select
    Range("a1:zz9999").ClearContents
    test = y1
    While (test <= y2)
        'URL Pattern
        'https://cals.arizona.edu/azmet/data/2017rd.txt
        urlString = "text;https://cals.arizona.edu/azmet/data/" & stationNum & test & "rd.txt"
        'ActiveSheet.QueryTables.Item(1).Delete
        ' Get USGS Data
        With ActiveSheet.QueryTables.Add(Connection:=urlString, Destination:=Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .Refresh
        End With
        Range("a1").Select
            While ActiveCell.Value <> ""
            collDataR.Add ActiveCell.Value & "," & ActiveCell.Offset(0, 1) & "," & ActiveCell.Offset(0, 25) 'adds one row into the data array
            ActiveCell.Offset(1, 0).Select 'jumps to the next row
            Wend
    test = test + 1
    Wend
    Set azmetDataR = collDataR 'set function return
       
End Function
Public Function azmetDataPen(stationNum, y1, y2) As Collection
    'build array for numbers
    Dim collDataR As Collection
    Set collDataR = New Collection
    'Call delete_Connections
    'clear contents for temp tab
    ActiveWorkbook.Sheets("temp3").Select
    Range("a1:zz9999").ClearContents
    test = y1
    While (test <= y2)
        'URL Pattern
        'https://cals.arizona.edu/azmet/data/2017rd.txt
        urlString = "text;https://cals.arizona.edu/azmet/data/" & stationNum & test & "rd.txt"
        'ActiveSheet.QueryTables.Item(1).Delete
        ' Get USGS Data
        With ActiveSheet.QueryTables.Add(Connection:=urlString, Destination:=Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .Refresh
        End With
        Range("a1").Select
            While ActiveCell.Value <> ""
            collDataR.Add ActiveCell.Value & "," & ActiveCell.Offset(0, 1) & "," & ActiveCell.Offset(0, 26) 'adds one row into the data array
            ActiveCell.Offset(1, 0).Select 'jumps to the next row
            Wend
    test = test + 1
    Wend
    Set azmetDataPen = collDataR 'set function return
       
End Function

'place data into into worksheet
Public Sub getAzmetData(dataArray)
    ActiveWorkbook.Sheets("precipitation").Select
    Dim findResult As Range
    
    For Each dataItem In dataArray
        vals = Split(dataItem, ",")
        t = DateSerial(vals(0), 1, vals(1))
        v = vals(2)
        v = ((v * 0.001) / 0.0254)
        v = Round(v, 2)
        On Error Resume Next
        findString = t
        Set findResult = ActiveSheet.Range("A8:A9999").Find(What:=findString, LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
            If (Not findResult Is Nothing) Then
                    ' Data found
                    ithrow = findResult.Row
                    ithCol = ActiveCell.Column
                    Cells(ithrow, ithCol).Value = v
            Else
            'no data found
            End If
    Next dataItem
   
   
End Sub
'place data into into worksheet
Public Sub getAzmetDataR(dataArray)
    ActiveWorkbook.Sheets("Reference_ET").Select
    Dim findResult As Range
    
    For Each dataItem In dataArray
        vals = Split(dataItem, ",")
        t = DateSerial(vals(0), 1, vals(1))
        v = vals(2)
        v = ((v * 0.001) / 0.0254)
        v = Round(v, 2)
        On Error Resume Next
        findString = t
        Set findResult = ActiveSheet.Range("A8:A9999").Find(What:=findString, LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
            If (Not findResult Is Nothing) Then
                    ' Data found
                    ithrow = findResult.Row
                    ithCol = ActiveCell.Column
                    Cells(ithrow, ithCol).Value = v
            Else
            'no data found
            End If
    Next dataItem
   
   
End Sub
Public Sub getAzmetDataPen(dataArray)
    ActiveWorkbook.Sheets("Reference_ET_PM").Select
    Dim findResult As Range
    
    For Each dataItem In dataArray
        vals = Split(dataItem, ",")
        t = DateSerial(vals(0), 1, vals(1))
        v = vals(2)
        On Error Resume Next
        findString = t
        Set findResult = ActiveSheet.Range("A8:A9999").Find(What:=findString, LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
            If (Not findResult Is Nothing) Then
                    ' Data found
                    ithrow = findResult.Row
                    ithCol = ActiveCell.Column
                    Cells(ithrow, ithCol).Value = v
            Else
            'no data found
            End If
    Next dataItem
   
   
End Sub
'create a function to take in day of the year and year and print a date
Public Function createDate(y1, y2)
    Range("a8").Select
    year1 = DateSerial(y1, 1, 1)
    year2 = DateSerial(y2, 12, 31)
    rowcounter = 0
    While year1 <= year2
        ActiveCell.Offset(rowcounter, 0).Value = year1
        year1 = DateAdd("d", 1, year1)
        rowcounter = rowcounter + 1
    Wend
    
End Function

Sub delete_Connections()
    Do While ActiveWorkbook.Connections.Count > 0
        ActiveWorkbook.Connections.Item(ActiveWorkbook.Connections.Count).Delete
    Loop
End Sub
