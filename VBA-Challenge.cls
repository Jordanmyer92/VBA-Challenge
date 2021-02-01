VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub stocks()

Dim readName As String
Dim nextName As String
Dim totalSV As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim read As Double
Dim change As Double
Dim i As Long
Dim lastRow As Double
Dim finalRow As Double
Dim ws As Worksheet

For Each ws In Worksheets
    

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    groupNo = 1
    totalSV = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        readName = ws.Cells(i, 1).Value
        nextName = ws.Cells(i + 1, 1).Value
        
        If nextName = readName Then
            totalSV = totalSV + ws.Cells(i, 7).Value
        Else
            
            totalSV = totalSV + ws.Cells(i, 7).Value
            ws.Cells(groupNo + 1, 9).Value = readName
            ws.Cells(groupNo + 1, 12).Value = totalSV
           
            groupNo = groupNo + 1
            totalSV = 0
            
            
            
            
    End If
        
        
Next i

    ws.Cells(1, 10).Value = "Yearly Change"
    groupNo = 1
    yearlyChange = 0
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    
    For i = 2 To lastRow
        read = ws.Cells(i, 3).Value
        change = ws.Cells(i + 1, 3).Value
        If change = read Then
            yearlyChange = yearlyChange + ws.Cells(i, 3).Value
            
            
        Else
            
            yearlyChange = yearlyChange + ws.Cells(i, 3).Value
            ws.Cells(groupNo + 1, 10).Value = read
           
            groupNo = groupNo + 1
            yearlyChange = 0
            
        End If
            

Next i

    groupNo = 1
    yearlyChange = 0
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    
    For i = 2 To lastRow
        read = ws.Cells(i, 6).Value
        change = ws.Cells(i, 10).Value
        If change = read Then
            yearlyChange = yearlyChange + ws.Cells(2, 7).Value
       
        Else
            
            yearlyChange = change - read
            ws.Cells(groupNo + 1, 6).Value = read
            ws.Cells(groupNo + 1, 10).Value = yearlyChange
            
            
            
            groupNo = groupNo + 1
            yearlyChange = 0
            
        End If
        
     
        
Next i

    lastRow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10) < 0 Then
             ws.Cells(i, 10).Interior.ColorIndex = 3
             
        End If
        
Next i
    
    groupNo = 1
    percentChange = 0
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    ws.Cells(1, 11) = "Percent Change"
        For i = 2 To lastRow
        read = ws.Cells(i, 3).Value
        change = ws.Cells(i, 11).Value
        If nextName = readName Then
            percentChange = percentChange + ws.Cells(i, 3).Value
        Else
            
            percentChange = percentChange + ws.Cells(i, 3).Value
            ws.Cells(groupNo + 1, 11).Value = read
          
            groupNo = groupNo + 1
             percentChange = 0
    
        End If
        
Next i
    
    groupNo = 1
    percentChange = 0
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    
    For i = 2 To lastRow
        On Error Resume Next
        read = ws.Cells(i, 6).Value
        change = ws.Cells(i, 11).Value
        If change = read Then
            percentChange = percentChange + ws.Cells(i, 11).Value
        'ElseIf (change - read) / change = 0 Then
            'ws.Cells(i, 11).Value = 0
            
        
        
        Else
            
            percentChange = (change - read) / change
            ws.Cells(groupNo + 1, 6).Value = read
            ws.Cells(groupNo + 1, 11).Value = percentChange
           
            groupNo = groupNo + 1
            totalSV = 0
            
            ws.Cells(i, 11).NumberFormat = "0.00%"
        End If
        
        
    Next i
    


Next ws

End Sub

Private Sub Workbook_Open()

End Sub
