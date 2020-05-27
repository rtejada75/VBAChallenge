Attribute VB_Name = "Module1"
Sub Year_Stock()

  ' Set an initial variable for holding the ticker name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the Open and Close Values
  Dim Open_Value As Double
  Dim Close_Value As Double
  Dim Open_ValueCounter As Long
  Dim Divide_Zero As Double

  
  Open_ValueCounter = 0
  
  ' Set an initial variable for holding the Total Volume per Stock
  Dim TickerVolume_Total As Double
  
  TickerVolume_Total = 0

  ' Keep track of the location for each Ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
   ' Keep track of Greatest % Increase/ Decrease / Total Volume
  Dim Ticker_Greatest As String
  Dim Greatest_Increase As Double
  Dim Greatest_Decrease As Double
  Dim Greatest_StockVolume As Double
   
   Greatest_Increase = 0
   Greatest_Decrease = 0
   Greatest_StockVolume = 0

  ' Loop through all Ticker Dates
  For i = 2 To 900000

    ' Check if we are still within the same Ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = Cells(i, 1).Value
      
      ' Set the Close_Value
      Close_Value = Cells(i, 6).Value

      ' Add to the Volume Total
      TickerVolume_Total = TickerVolume_Total + Cells(i, 7).Value
      
      
      
      Open_Value = Cells(i - Open_ValueCounter, 3).Value

      ' Print the Ticker Name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Ticker Volume Total to the Summary Table
      Range("L" & Summary_Table_Row).Value = TickerVolume_Total

      ' Print the Yearly Change to the Summary Table
      Range("J" & Summary_Table_Row).Value = Close_Value - Open_Value
      
      ' Color change in Ticker price...Red Negative; Green Positive
      If Cells(Summary_Table_Row, 10).Value < 0 Then
      Range("J" & Summary_Table_Row).Interior.Color = vbRed
      
      Else
      Range("J" & Summary_Table_Row).Interior.Color = vbGreen
      
      End If
      
      ' Print the Yearly % Change to the Summary Table
       If Open_Value = 0 Then
       
       Divide_Zero = 0
       
       Range("K" & Summary_Table_Row).Value = Divide_Zero
       
       Else
       
       Divide_Zero = (Close_Value / Open_Value - 1)
       
       End If
       
       Range("K" & Summary_Table_Row).Value = Divide_Zero
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Volume Total
      TickerVolume_Total = 0
      
      ' Reset the Open Value Counter
      Open_ValueCounter = 0

    ' If the cell immediately following a row is the same Ticker...
    
    Else

      ' Add to the Ticker Volume Total
      TickerVolume_Total = TickerVolume_Total + Cells(i, 7).Value
      
      ' Add to Open Value Counter
      Open_ValueCounter = Open_ValueCounter + 1
      
      ' Set the Close_Value
      Close_Value = Cells(i, 6).Value

    End If

  Next i
 
 
 For J = 2 To 80000

' Determine Greatest Ticker Increase
    If Greatest_Increase > Cells(J, 11).Value Then
    
    Greatest_Increase = Greatest_Increase
    
    Else

    Greatest_Increase = Cells(J, 11).Value
    
    Ticker_Greatest = Cells(J, 9).Value

' Print the Ticker and Greatest Increase
      Range("Q2").Value = Ticker_Greatest
      Range("R2").Value = Greatest_Increase

End If

Next J



For Z = 2 To 80000

' Determine Greatest Ticker Decrease
    If Greatest_Decrease < Cells(Z, 11).Value Then
    
    Greatest_Decrease = Greatest_Decrease
    
    Else

    Greatest_Decrease = Cells(Z, 11).Value
    
    Ticker_Greatest = Cells(Z, 9).Value

End If

Next Z

' Print the Ticker and Greatest Decrease
      Range("Q3").Value = Ticker_Greatest
      Range("R3").Value = Greatest_Decrease
      
      
      
For Y = 2 To 80000

' Determine Greatest Stock Increase
    If Greatest_StockVolume > Cells(Y, 12).Value Then
    
    Greatest_StockVolume = Greatest_StockVolume
    
    Else

    Greatest_StockVolume = Cells(Y, 12).Value
    
    Ticker_Greatest = Cells(Y, 9).Value

End If

Next Y

' Print the Yearly % Change to the Summary Table
      Range("Q4").Value = Ticker_Greatest
      Range("R4").Value = Greatest_StockVolume

End Sub



