VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Market()

' Loop through each sheet in workbook
          For Each ws In Worksheets
          
    
    ' Label Columns and Define Variables
   
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1") = "Total Stock Value"
   
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
   
    Dim Ticker_Symbol As String
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    Dim Stock_Volume As LongLong
    Stock_Value = 0
    Dim Annual_Open As Double
    Dim Annual_Close As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As String
    
    
    
    
    Annual_Open = ws.Cells(2, 3).Value

    ' Loop through each row
        For r = 2 To LastRow
           

        ' Check if we are still within the same ticker symbol
                If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
                                
                    Ticker_Symbol = ws.Cells(r, 1).Value
                    
        ' Initializing first open and close stock value
                    Annual_Close = ws.Cells(r, 6).Value
                    
       ' Calculate Yearly Change
    
                    Yearly_Change = Annual_Close - Annual_Open
                    
        ' Calculate Percent Change
                    If Annual_Open = 0 Then
                        Percent_Change = 0
                        
                        Else
                        
                            Percent_Change = FormatPercent((((Annual_Close - Annual_Open) / Annual_Open)), 2, vbUseDefault, vbUseDefault, vbUseDefault)
                    
                    End If
                    
                    
        ' Add to Total Stock Volume
                    Stock_Volume = Stock_Volume + ws.Cells(r, 7).Value

        ' Add the ticker symbol to the Summary Table
                    ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
                    
        ' Add Stock Total to Summary Table
                    ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                    
        ' Add Yearly Change to Summary Table
                    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                    
        ' Add Percent Change to Summary Table
                    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                    
                        
        
        ' Add one row to the Summary Table Row
                    Summary_Table_Row = Summary_Table_Row + 1
                    
        ' Reset Total Stock Volume
                    Stock_Volume = 0

         
                    Annual_Open = ws.Cells(r + 1, 3).Value
            Else
            
                Stock_Volume = Stock_Volume + ws.Cells(r, 7).Value
                Ticker_Symbol = ws.Cells(r, 1).Value
            
                   
        End If
    
    ' If positive change, green, if negative, red
        
        If ws.Cells(r, 11).Value < 0 Then
                        
                            ws.Cells(r, 11).Interior.ColorIndex = 3
                            
                        ElseIf ws.Cells(r, 11).Value > 0 Then
                        
                            ws.Cells(r, 11).Interior.ColorIndex = 4
                            
                        Else
                        
                              
                            ws.Cells(r, 11).Interior.ColorIndex = 2
                
                        
                            
                        End If
                            
                    
       
            
   
    ' Loop
    
    Next r
   
Next ws


   
End Sub


