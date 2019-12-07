VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock()

Dim Ticker_Symbol As String

Dim Ticker_Vol As Double
Ticker_Symbol = 0

Dim Yearly_change As Double

Dim Percent_change As Double

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

For i = 2 To 705714

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker_Symbol = Cells(i, 1).Value
        
        Ticker_Vol = Ticker_Vol + Cells(i, 7).Value
        
        Range("K" & Summary_Table_Row).Value = Yearly_change
        
        Range("L" & Summary_Table_Row).Value = Percent_change
        
        Range("J" & Summary_Table_Row).Value = Ticker_Symbol
        
        Range("M" & Summary_Table_Row).Value = Ticker_Vol

        Summary_Table_Row = Summary_Table_Row + 1
        
        Ticker_Vol = 0
        
        Yearly_change = 0
        
        Percent_change = 0
        
        
Else
    
    Yearly_change = Yearly_change + Cells(i + 1, 3).Value - Cells(i + 261, 6).Value

    Ticker_Vol = Ticker_Vol + Cells(i, 7).Value
    

    End If
    
    If Cells(i, 11).Value > 0 Then
    
        Cells(i, 11).Interior.ColorIndex = 4
    
    ElseIf Cells(i, 11).Value < 0 Then
     
        Cells(i, 11).Interior.ColorIndex = 3
    
    End If
    
Next

End Sub

