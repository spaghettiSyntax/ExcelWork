Option Explicit


Const FOLDER_PATH = "C:\[REMOVED BEFORE GITHUB]\"  'REMEMBER END BACKSLASH


Sub ImportWorksheets()
    '=============================================
    'Process all Excel files in specified folder
    '=============================================
    Dim sFile As String           'file to process
    Dim wsTarget As Worksheet
    Dim wsTarget2 As Worksheet
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim rowTarget As Long         'output row Sheet1
    Dim sheetTarget As Integer    'sheet selection
    Dim memberNameSearch As Long  'search for member names
    Dim dateOfTournament As Date
    
    rowTarget = 1
    sheetTarget = 1
   
    'check the folder exists
    If Not FileFolderExists(FOLDER_PATH) Then
       MsgBox "Specified folder does not exist, exiting!"
       Exit Sub
    End If
   
    'reset application settings in event of error
    On Error GoTo errHandler
    Application.ScreenUpdating = False
    
    'set up the target worksheet
    Set wsTarget = Sheets("Sheet1")
    
    'loop through the Excel files in the folder
    sFile = Dir(FOLDER_PATH & "*.xls*")
    Do Until sFile = ""
        
        Application.DisplayAlerts = False
        
        'open the source file and set the source worksheet - ASSUMED WORKSHEET(1)
        Set wbSource = Workbooks.Open(FOLDER_PATH & sFile)

        For sheetTarget = 1 To 7
            Set wsSource = wbSource.Worksheets(sheetTarget)
            'import the data
            With wsTarget
                If sheetTarget = 1 Then
                    Dim tournamentFee As Double
                    Dim taxRate As Double
                    tournamentFee = wsSource.Range("B4").Value
                    taxRate = wsSource.Range("K4").Value
                    dateOfTournament = wsSource.Range("J1").Value
                    '.Range("A" & rowTarget).Value = "M Num"
                    '.Range("B" & rowTarget).Value = "M Name"
                    '.Range("D" & rowTarget).Value = "Tour Fee"
                    '.Range("E" & rowTarget).Value = "Tax Rate"
                    '.Range("F" & rowTarget).Value = "Tax Paid"
                    '.Range("G" & rowTarget).Value = "Location"
                End If
                If sheetTarget >= 2 Then
                    For memberNameSearch = 9 To 208
                        If wsSource.Range("D" & memberNameSearch).Text <> "" Then
                            .Range("A" & rowTarget).Value = wsSource.Range("C" & memberNameSearch).Value
                            .Range("B" & rowTarget).Value = wsSource.Range("D" & memberNameSearch).Value
                            .Range("D" & rowTarget).Value = tournamentFee
                            .Range("D" & rowTarget).NumberFormat = "$0.00"
                            .Range("E" & rowTarget).Value = taxRate
                            .Range("E" & rowTarget).NumberFormat = "0.00%"
                            .Range("F" & rowTarget).Value = wsSource.Range("AA" & memberNameSearch).Value
                            .Range("G" & rowTarget).Value = dateOfTournament
                            .Range("H" & rowTarget).Value = sFile
                            rowTarget = rowTarget + 1
                        End If
                    Next memberNameSearch
                End If
            End With
        Next sheetTarget
        
        'reset sheetTarget, close the source workbook, and get the next file
        sheetTarget = 1
        wbSource.Close SaveChanges:=False
        sFile = Dir()
    Loop
    
    wsTarget.Range("A1", Range("H1").End(xlDown)).Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlNo
    
errHandler:
    On Error Resume Next
    Application.ScreenUpdating = True
    
    'tidy up
    Set wsSource = Nothing
    Set wbSource = Nothing
    Set wsTarget = Nothing
End Sub

Private Function FileFolderExists(strPath As String) As Boolean
    If Not Dir(strPath, vbDirectory) = vbNullString Then FileFolderExists = True
End Function
