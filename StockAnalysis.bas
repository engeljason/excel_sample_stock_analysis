Attribute VB_Name = "StockAnalysis"
Sub StockAnalysis()
    ' for each sheet
        ' for each row
        ' filter by date update: not needed (already filtered into sheets by date)
            ' aggregate by ticker symbol
                ' total changes
                    ' format by color
                ' calculate percent change
                ' total volume
    
    ' columns
    ' (1) ticker | (2) date | (3) open | (4) high | (5) low | (6) close | (7) vol
   
    If Not Worksheets(1).Name = "Results" Then
        Worksheets.Add(Before:=Sheets(1)).Name = "Results"
    Else
        Worksheets("Results").Cells.Clear
    End If
    
    Dim resultsheet As Worksheet
    Set resultsheet = Worksheets("Results")
    
    ' result sheet row/column for current stock sheet
    Dim rsr As Integer
    Dim rsc As Integer
    
    For Each ws In Worksheets:
        If ws.Name = "Results" Then
            ' do nothing
            ' may as well initialize result block here
            rsr = 3
            rsc = 1
        Else:
            ' setup new result block in resultsheet
            resultsheet.Cells(2, rsc).Value = "Ticker"
            resultsheet.Columns(rsc).AutoFit
            resultsheet.Cells(2, rsc + 1).Value = "Yearly Change"
            resultsheet.Columns(rsc + 1).AutoFit
            resultsheet.Cells(2, rsc + 2).Value = "Percent Change"
            resultsheet.Columns(rsc + 2).AutoFit
            resultsheet.Cells(2, rsc + 3).Value = "Total Volume"
            resultsheet.Columns(rsc + 3).AutoFit
            ' add titlebar after autofit so it doesn't wreck formatting
            titlebar = ws.Name + " Stock Summary"
            resultsheet.Cells(1, rsc).Value = titlebar
            
            ' bonus section
            resultsheet.Cells(1, rsc + 5).Value = ws.Name + "'s Greatest"
            resultsheet.Cells(2, rsc + 5).Value = "Percent Increase"
            resultsheet.Cells(3, rsc + 5).Value = "Percent Decrease"
            resultsheet.Cells(4, rsc + 5).Value = "Total Volume"
            resultsheet.Columns(rsc + 5).AutoFit
            resultsheet.Cells(1, rsc + 6).Value = "Ticker"
            resultsheet.Columns(rsc + 6).AutoFit
            resultsheet.Cells(1, rsc + 7).Value = "Value"
            resultsheet.Columns(rsc + 7).AutoFit
            
            ' find range of data
            Dim lastrow As Long
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            ''' gather data
            ' ---------------------------
            '''
            Dim stockdict As Scripting.Dictionary
            Set stockdict = New Scripting.Dictionary
    
            For i = 2 To lastrow:
                    
                Dim irow As Range
                Set irow = ws.Rows(i)
                
                Dim ticker As Variant
                ticker = irow.Cells(1).Value2
                
                Dim mystock As Stock
                
                If Not stockdict.Exists(ticker) Then
                    Set mystock = New Stock
                    mystock.Name = ticker
                    Call stockdict.Add(ticker, mystock)
                End If
                
                Set mystock = stockdict(ticker)
                
                Call mystock.Update(irow.Cells(3).Value2, irow.Cells(6).Value2, irow.Cells(2).Value2, irow.Cells(7).Value2)
                
                ' Debug.Print (mystock.ToString())
                
            Next i
            ''' end gather data
            ' ---------------------------
            '''
            
            ' award winning elements
            Dim stkMostGain As Stock
            Dim stkMostLoss As Stock
            Dim stkMostVolume As Stock
            Dim first As Boolean
            first = True
            ' update data in result sheet
            For Each skey In stockdict.Keys
                Dim kStock As Stock
                Set kStock = stockdict(skey)
                
                ' pull out values once
                Dim tchange As Double
                tchange = kStock.TotalChange
                Dim pchange As Double
                pchange = kStock.PercentChange
                Dim tvolume As Double
                tvolume = kStock.Volume
                
                ' find bonus winners
                If first Then
                    Set stkMostGain = kStock
                    Set stkMostLoss = kStock
                    Set stkMostVolume = kStock
                    first = False
                Else
                    If tchange > stkMostGain.TotalChange Then
                        Set stkMostGain = kStock
                    End If
                    If tchange < stkMostLoss.TotalChange Then
                        Set stkMostLoss = kStock
                    End If
                    If tvolume > stkMostVolume.Volume Then
                        Set stkMostVolume = kStock
                    End If
                End If
                
                ' populate current row
                resultsheet.Cells(rsr, rsc).Value = skey
                resultsheet.Cells(rsr, rsc + 1).Value = tchange
                resultsheet.Cells(rsr, rsc + 1).NumberFormat = "0.00"
                resultsheet.Cells(rsr, rsc + 2).Value = pchange
                resultsheet.Cells(rsr, rsc + 2).NumberFormat = "0.00%"
                resultsheet.Cells(rsr, rsc + 3).Value = tvolume
                If (tchange > 0) Then
                    resultsheet.Cells(rsr, rsc + 1).Interior.ColorIndex = 4
                ElseIf (tchange = 0) Then
                    resultsheet.Cells(rsr, rsc + 1).Interior.ColorIndex = 5
                Else
                    resultsheet.Cells(rsr, rsc + 1).Interior.ColorIndex = 3
                End If
                
                ' Next row
                rsr = rsr + 1
            Next skey
            
            ' populate winners bracket
            winc = rsc + 6
            resultsheet.Cells(2, winc).Value = stkMostGain.Name
            resultsheet.Cells(2, winc + 1) = stkMostGain.PercentChange
            resultsheet.Cells(2, winc + 1).NumberFormat = "0.00%"
            resultsheet.Cells(3, winc).Value = stkMostLoss.Name
            resultsheet.Cells(3, winc + 1) = stkMostLoss.PercentChange
            resultsheet.Cells(3, winc + 1).NumberFormat = "0.00%"
            resultsheet.Cells(4, winc).Value = stkMostVolume.Name
            resultsheet.Cells(4, winc + 1) = stkMostVolume.Volume
            
            ' update current result sheet column
            rsc = rsc + 9
            rsr = 3
            ' End If Not resultsheet
        End If
        ' next worksheet
    Next ws
End Sub

