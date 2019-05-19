Sub Stock_Summary_Loop():

'-----------Set Loop for All Sheets-----------------

' Loop through all sheets
    
    For Each ws In ActiveWorkbook.Worksheets
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

'------Nest For-Loop in For-Each WS Loop------

' Create a variable for holding the ticker

    Dim Ticker As String

' Create and set variable for holding total vol. per ticker

    Dim Volume_Total As Double
    Vol_total = 0

' Keep track of the location for each ticker on summary sheet

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
' Loop through all ticker sheets

    For i = 2 To LastRow

' Check if we are still within the ticker, if not...

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ' Set ticker name
        Ticker = ws.Cells(i, 1).Value
        
    ' Add to the Volume Total
        
       Vol_total = Vol_total + ws.Cells(i, 7).Value
       
    'Print ticker name on summary sheet
        
       ws.Range("J" & Summary_Table_Row).Value = Ticker
       
    'Print volume total on summary sheet
    
        ws.Range("K" & Summary_Table_Row).Value = Vol_total
        
    'Add one to summary sheet row
    
        Summary_Table_Row = Summary_Table_Row + 1
        
    'Reset volume total
        
        Vol_total = 0
        
'If cell immediately follow a row is same ticker..

    Else

    'Add to the volume total
    
        Vol_total = Vol_total + ws.Cells(i, 7).Value
        
    End If
    
   
    Next i
    
Next ws
        
       

End Sub


