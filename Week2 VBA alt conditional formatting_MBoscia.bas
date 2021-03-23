Attribute VB_Name = "Module5"
Option Explicit

Sub TickerMultiSheet2()


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'Step 1) Loop thru multiple worksheets --> to be added after single sheet logic successfully runs.
'Step 2) Define Variables & Create Summary table for each change in primary series
'Step 2a) add Headers to summary table
'Step 3) Create Summary critera
'           a) Begging Price
'           b) Ending Price
'           c) Annual Change in Price(Inc/dec value)
'           d) Annual Change in Price as %
'           e) Total Annual Volume
'Step 4) Output Step 3c-3e
'...End Loop
'Step 5) Add conditional formatting to 3c (positve=green, negative=red)
'
'Bonus
'Step 1) Create headers & item descriptors
'Step 2) loop thru Summary Table extract
'           a) Max Price as %
'           b) Min Price as %
'           c) Max Total Volume
'Step 3)Output related ticker with identified values in respective Cells
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'====================================================================

'Here we gooooo!
'Step 1) Step Through multiple sheets

    Dim ws As Worksheet                  'define variable for worksheet name
    Dim WorksheetName As String
    
For Each ws In Worksheets
    WorksheetName = ws.Name              'assign worksheet name


'Step 2)

        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row   'Add ws.Cells for multi-sheet step-thru
        
        
        'Variables
        Dim i As Long
        Dim Ticker As String
        Dim BegP As Double
        Range("C2") = BegP
        Dim EndP As Double
        Dim Change As Double
        Dim ChangeP As Double
        Dim Volume As Double
        Volume = 0
        
        'Establish Summary Table
        Dim Summary_Row As Integer
        Summary_Row = 2
        ws.Range("J1").Value = "Ticker"
        ws.Range("k1").Value = "$Change"
        ws.Range("l1").Value = "%Change"
        ws.Range("M1").Value = "Annual Volume"

        
        'Begin looping thru data
        For i = 2 To LastRow
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value                      'Extract Unique Ticker Symbol
                ws.Range("J" & Summary_Row).Value = Ticker         'Output Unique Ticker to Summary table when next line changes
 'Step 3)&4)
                EndP = ws.Cells(i, 6).Value                         'Assign Ending Price (kinda odd but trigger is at the end of a range when the chage occurs)
                Change = EndP - BegP                                'Calculate net change of price
                ws.Range("K" & Summary_Row).Value = Change          'Output annual inc/dec value to Summary table
                 
                 'Conditional Formatting for Change value output
                    If Change > 0 Then
                        ws.Cells(Summary_Row, 11).Interior.ColorIndex = 4
                    Else
                        ws.Cells(Summary_Row, 11).Interior.ColorIndex = 3
                    End If
                
                 
                 
                 'Calcuate % change (inc/dec)
                    If BegP = 0 Then
                        ChangeP = 1                             'Bypass div/0 error
                    Else
                        ChangeP = (Change / BegP)
                    End If
                    
                ws.Range("l" & Summary_Row).Value = ChangeP        'Output annual inc/dec as %
                
                Volume = Volume + ws.Cells(i, 7).Value             'Final Volume sum
                ws.Range("M" & Summary_Row).Value = Volume         'Output Final Volume tally to Summary table
                    

                Volume = 0                                      'Reset Volume for next Ticker
                BegP = ws.Cells(i + 1, 3).Value                    'Establish Beginng price for next Ticker
                Summary_Row = Summary_Row + 1                   'Move to the next line on the Summary Table for next value
                        
                
            Else                                                'No change in Ticker, keep adding
                Volume = Volume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        

        
'Step 5 - Formating

     ws.Range("L2", "L" & Summary_Row).NumberFormat = "0.00%"
     


'Bonus

    'Establish Bonus Table & Variables

        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Result"
        ws.Range("P2").Value = "Biggest Incr(%)"
        ws.Range("P3").Value = "Biggest Decr(%)"
        ws.Range("P4").Value = "Highest Volume"
        
        Dim BInc As Double
        Dim BDec As Double
        Dim HVol As Double
        
        Dim HTick As String
        Dim LTick As String
        Dim VTick As String
        
    'Biggest % Increase
    Dim j As Integer
    BInc = 0                                          'Set beginning values
    HTick = ""
    
       For j = 2 To Summary_Row
            If ws.Cells(j, 12).Value > BInc Then           'if next is bigger, reset values
                BInc = ws.Cells(j, 12).Value
                HTick = ws.Cells(j, 10).Value
            Else                                        'no alternate logic
            End If
       Next j                                           'go to next line
       
       ws.Range("Q2").Value = HTick                        'Ouput final Ticker to Bonus Table
       ws.Range("R2").Value = BInc                         'Ouput final value to Bonus Table
        
    'Biggest % Decrease
    Dim k As Integer
    BDec = 0                                            'Set beginning values
    LTick = ""
    
       For k = 2 To Summary_Row
            If ws.Cells(k, 12).Value < BDec Then           'if next is smaller, reset values
                BDec = ws.Cells(k, 12).Value
                LTick = ws.Cells(k, 10).Value
            Else                                        'no alternate logic
            End If
       Next k                                           'go to next line
       
       ws.Range("Q3").Value = LTick                        'Ouput final Ticker to Bonus Table
       ws.Range("R3").Value = BDec                         'Ouput final value to Bonus Table
    
    
  'Hghest Annual Volume
    Dim l As Integer
    HVol = 0                                            'Set beginning values
    VTick = ""
    
       For l = 2 To Summary_Row
            If ws.Cells(l, 13).Value > HVol Then           'if next is smaller, reset values
                HVol = ws.Cells(l, 13).Value
                VTick = ws.Cells(l, 10).Value
            Else                                        'no alternate logic
            End If
       Next l                                           'go to next line
       
       ws.Range("Q4").Value = VTick                        'Ouput final Ticker to Bonus Table
       ws.Range("R4").Value = HVol                         'Ouput final value
        
        
        
        ws.Range("r2", "r3").NumberFormat = "0.00%"
        ws.Cells.EntireColumn.AutoFit                      ' May need to adjust for multi-WS
   
Next ws

End Sub


