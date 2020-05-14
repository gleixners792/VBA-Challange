Attribute VB_Name = "TheVBAofWallStreet"
Sub Stck_Mrkt_Anlyst()

' Declare Variables

Dim SheetName As String
Dim LastRow As Double
Dim RowCnt As Double

Dim Ticker As String
Dim TickerCnt As Long
Dim FirstTicker As Boolean
Dim OpenPrice As Currency
Dim ClosePrice As Currency
Dim Volume As Double

Dim TickCol As Integer
Dim YrlyCol As Integer
Dim PrcntCol As Integer
Dim VolCol As Integer

' Define Columns for Writing Result Data.

TickCol = 9
YrlyCol = 10
PrcntCol = 11
VolCol = 12

' Go through Each Sheet in the Workbook

For Each ws In Worksheets

SheetName = ws.Name
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'MsgBox ("Sheet Name--> " & SheetName & " Number of Rows--> " & LastRow)

' Set-up Column Headings for Each New Sheet

ws.Cells(1, TickCol).Value = "Ticker"
ws.Cells(1, YrlyCol).Value = "Yearly Change"
ws.Cells(1, PrcntCol).Value = "Percent Change"
ws.Cells(1, VolCol).Value = "Total Stock Volume"

' Reset Variables for Each Sheet

TickerCnt = 0
FirstTicker = True
Volume = 0
OpenPrice = 0
ClosePrice = 0
Ticker = " "

' Go through all Row Data and Tabulate Results in New Columns

For RowCnt = 2 To LastRow

' Check to see if we need to post data into the spreadsheet

If ws.Cells(RowCnt, 1).Value <> Ticker Then

    If Not FirstTicker Then
    
    ws.Cells(TickerCnt + 1, TickCol).Value = Ticker
    ws.Cells(TickerCnt + 1, YrlyCol).Value = ClosePrice - OpenPrice
    
'Check to make sure we are not dividing by zero.
    If (OpenPrice = 0) Then
        ws.Cells(TickerCnt + 1, PrcntCol).Value = 1
    Else
        ws.Cells(TickerCnt + 1, PrcntCol).Value = (ClosePrice - OpenPrice) / OpenPrice
    End If
    
    ws.Cells(TickerCnt + 1, VolCol).Value = Volume
    ws.Cells(TickerCnt + 1, PrcntCol).NumberFormat = "0.00%"

'Shade Cells based upon Stock Gain or Loss
    
    If (ClosePrice - OpenPrice) >= 0 Then
    
    ws.Cells(TickerCnt + 1, YrlyCol).Interior.ColorIndex = 10
    
    Else
    
    ws.Cells(TickerCnt + 1, YrlyCol).Interior.ColorIndex = 3
    
    End If
    
End If

' Reset Data on New Ticker in the Sheet

FirstTicker = False
TickerCnt = TickerCnt + 1
Volume = ws.Cells(RowCnt, 7).Value
OpenPrice = ws.Cells(RowCnt, 3).Value
ClosePrice = ws.Cells(RowCnt, 6).Value
Ticker = ws.Cells(RowCnt, 1).Value

Else

' Continue to Collect Data on the Same Ticker
Volume = ws.Cells(RowCnt, 7).Value + Volume
ClosePrice = ws.Cells(RowCnt, 6).Value


End If


Next RowCnt

' Post the last ticker data in the spreadsheet

    ws.Cells(TickerCnt + 1, TickCol).Value = Ticker
    ws.Cells(TickerCnt + 1, YrlyCol).Value = ClosePrice - OpenPrice
    
'Check to make sure we are not dividing by zero.
    If (OpenPrice = 0) Then
        ws.Cells(TickerCnt + 1, PrcntCol).Value = 1
    Else
        ws.Cells(TickerCnt + 1, PrcntCol).Value = (ClosePrice - OpenPrice) / OpenPrice
    End If
    
    ws.Cells(TickerCnt + 1, VolCol).Value = Volume
    ws.Cells(TickerCnt + 1, PrcntCol).NumberFormat = "0.00%"

'Shade Cells based upon Stock Gain or Loss
    
    If (ClosePrice - OpenPrice) >= 0 Then
    
    ws.Cells(TickerCnt + 1, YrlyCol).Interior.ColorIndex = 10
    
    Else
    
    ws.Cells(TickerCnt + 1, YrlyCol).Interior.ColorIndex = 3
    
    End If

'MsgBox ("Made it to the end--> " & RowCnt)

Next ws

End Sub
