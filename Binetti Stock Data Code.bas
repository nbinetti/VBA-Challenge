Attribute VB_Name = "Module1"
Sub Stock_Data()
    ' Set CurrentWs as a worksheet object variable
    Dim CurrentWs As Worksheet
    Dim Need_Summary_Table_Header As Boolean
    Dim Command As Boolean
    
    Need_Summary_Table_Header = False
    'Set Header
    Command = True
    
    ' Loop through all worksheets in active workbook
    For Each CurrentWs In Worksheets
    
        ' Set ticker name variable
        Dim Ticker As String
        Ticker = " "
        
        ' Set variable for holding ticker name total
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        
        ' Set new variables for Moderate Solution Part
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Change_Price As Double
        Change_Price = 0
        Dim Change_Percent As Double
        Change_Percent = 0
        ' Set new variables for Hard Solution Part
        Dim Max_Ticker_Name As String
        Max_Ticker_Name = " "
        Dim Min_Ticker_Name As String
        Min_Ticker_Name = " "
        Dim Max_Percent As Double
        Max_Percent = 0
        Dim Min_Percent As Double
        Min_Percent = 0
        Dim Max_Volume_Ticker As String
        Max_Volume_Ticker = " "
        Dim Max_Volume As Double
        Max_Volume = 0
         
        ' Keep track of each ticker name
        ' in summary table for current worksheet
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        ' Set initial row count for the current worksheet
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        ' Results for all worksheets except the first one
        If Need_Summary_Table_Header Then
            ' Set Titles for the Summary Table for current worksheet
            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"
            ' Set titles for bonus Summary Table on the right for current worksheet
            CurrentWs.Range("O2").Value = "Greatest % Increase"
            CurrentWs.Range("O3").Value = "Greatest % Decrease"
            CurrentWs.Range("O4").Value = "Greatest Total Volume"
            CurrentWs.Range("P1").Value = "Ticker"
            CurrentWs.Range("Q1").Value = "Value"
        Else
            'first worksheet, reset flag for the rest
            Need_Summary_Table_Header = True
        End If
        
        ' Set value of oen price for the first Ticker of CurrentWs
        ' ticker open price for remainder
        Open_Price = CurrentWs.Cells(2, 3).Value
        
        ' Loop from the beginning of the current worksheet(Row2) till its last row
        For i = 2 To Lastrow
        
      
            ' Check next to see if ticker has changed
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
                ' Set the ticker name
                Ticker_Name = CurrentWs.Cells(i, 1).Value
                
                ' Calculate Change_Price and Change_Percent
                Close_Price = CurrentWs.Cells(i, 6).Value
                Change_Price = Close_Price - Open_Price
                ' Check Division by 0 condition
                If Open_Price <> 0 Then
                    Change_Percent = (Change_Price / Open_Price) * 100
                Else
                    ' In case it does not return anything
                    MsgBox ("For " & Ticker_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet.")
                End If
                
                ' Add to ticker total volume
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
              
                
                ' Print the Ticker Names and Price Change in Column I and J
                CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
                CurrentWs.Range("J" & Summary_Table_Row).Value = Change_Price
        
                If (Change_Price > 0) Then
                    'highlight green
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Change_Price <= 0) Then
                    'highlight red
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 ' Print the Percent Change in Column K
                CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Change_Percent) & "%")
                ' Print total Volume in Column L
                CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                ' Add 1 to the summary table row count
                Summary_Table_Row = Summary_Table_Row + 1
                ' Reset Change_Price and Close Price variables to zero
                Change_Price = 0
                Close_Price = 0
                ' identify next Ticker's Open_Price
                Open_Price = CurrentWs.Cells(i + 1, 3).Value
              
                
                ' New Summary Table
                If (Change_Percent > Max_Percent) Then
                    Max_Percent = Change_Percent
                    Max_Ticker_Name = Ticker_Name
                ElseIf (Change_Percent < Min_Percent) Then
                    Min_Percent = Change_Percent
                    Min_Ticker_Name = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > Max_Volume) Then
                    Max_Volume = Total_Ticker_Volume
                    Max_Volume_Ticker = Ticker_Name
                End If
                
                ' Reset variables to zero
                Change_Percent = 0
                Total_Ticker_Volume = 0
                
            
            ' If next cell after a row is still the same ticker, just add to TTV
            Else
                ' find the TTV
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
            End If
            ' test MsgBox (CurrentWs.Rows(i).Cells(2, 1))
      
        Next i

            If Not Command Then
            
                CurrentWs.Range("Q2").Value = (CStr(Max_Percent) & "%")
                CurrentWs.Range("Q3").Value = (CStr(Min_Percent) & "%")
                CurrentWs.Range("P2").Value = Max_Ticker_Name
                CurrentWs.Range("P3").Value = Min_Ticker_Name
                CurrentWs.Range("Q4").Value = Max_Volume
                CurrentWs.Range("P4").Value = Max_Volume_Ticker
                
            Else
                Command = False
            End If
        
     Next CurrentWs
End Sub
