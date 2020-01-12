Option Explicit

Const SOURCE_PATH = "C:\[REMOVED BEFORE GITHUB]\"
Const EMAIL_SOURCE_PATH = "C:\[REMOVED BEFORE GITHUB]"
Const DESTINATION_PATH = "C:\[REMOVED BEFORE GITHUB]\"
Const TEMPLATE_PATH = "C:\[REMOVED BEFORE GITHUB]"

Sub ActivateWordTransferData()

    '==========================================='
    'Process all Excel files in specified folder'
    '==========================================='
    
    Dim sFile As String           'file to process
    Dim eFile As String           'email file to process
    Dim wdapp As Object
    Dim wddoc As Object
    Dim wsTarget As Worksheet
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim sheetTarget As Integer    'sheet selection
    Dim memberEmailSearch As Integer
    Dim memNum_text As String, memNumTo_text As String
    Dim memName_text As String, memNameTo_text As String
    Dim memEmail_text As String, memEmailTo_text As String
    Dim memCreditTotal_text As String, memCreditTotalTo_text As String
    
    'when doing late binding you need to initialize replaceAll prior
    'http://www.vbforums.com/showthread.php?563745-RESOLVED-Find-Replace-Problem
    Const wdReplaceAll = 2
    
    'reset application settings in event of error
    On Error Resume Next
    
    'On Error GoTo errHandler
    Application.ScreenUpdating = False
    
    'set up the target worksheet
    Set wsTarget = Sheets("Sheet1")
    
    'order Matters When Using Dir() use looped Dir after single use Dir()
    'https://stackoverflow.com/questions/51058113/vba-dir-returns-empty-string
    'email file to traverse
    eFile = Dir(EMAIL_SOURCE_PATH)
    'loop through the Excel files in the folder
    sFile = Dir(SOURCE_PATH & "*.xls*")
    
    Do Until sFile = ""
        Application.DisplayAlerts = False
        
        Dim memberName As String
        
        'open the source file and set the source worksheet - ASSUMED WORKSHEET(1)
        Set wbSource = Workbooks.Open(SOURCE_PATH & sFile)

        For sheetTarget = 1 To 2
            Set wsSource = wbSource.Worksheets(sheetTarget)
            
            'grab the data
            With wsTarget
                If sheetTarget = 1 Then
                    memNum_text = Sheets("Cover").Range("A5").Value
                    memNumTo_text = Sheets("Cover").Range("A6").Value
                    memName_text = Sheets("Cover").Range("B8").Value
                    memCreditTotal_text = Sheets("Cover").Range("B2").Value
                    memCreditTotalTo_text = Sheets("Cover").Range("C2").Value
                End If
                If sheetTarget = 2 Then
                    memberName = wsSource.Range("B2").Value
                    memNameTo_text = memberName
                End If
            End With
        Next sheetTarget
        
        'reset sheetTarget, close the source workbook, empty out word objects, and get the next file
        sheetTarget = 1
        wbSource.Close SaveChanges:=False
        
        'traverse email workbook for associated member email
        Set wbSource = Workbooks.Open(EMAIL_SOURCE_PATH)
        
        For sheetTarget = 1 To 1
            Set wsSource = wbSource.Worksheets(sheetTarget)
            
            'grab the data
            With wsTarget
                memEmail_text = Sheets("export").Range("G1").Value
                memEmailTo_text = "Unavailable"
                For memberEmailSearch = 1 To 1198
                    If wsSource.Range("A" & memberEmailSearch).Text <> "" Then
                        Dim memEmailName As String
                        memEmailName = Sheets("export").Range("C" & memberEmailSearch).Value & ", " & Sheets("export").Range("B" & memberEmailSearch).Value
                        If memEmailName = memNameTo_text Then
                            memEmailTo_text = Sheets("export").Range("A" & memberEmailSearch).Value
                            'reset sheetTarget, close the source workbook, returnEmail
                            sheetTarget = 1
                            wbSource.Close SaveChanges:=False
                            Exit For
                        End If
                    End If
                Next memberEmailSearch
            End With
        Next sheetTarget
        
        'start word process
        Set wdapp = GetObject(, "Word.Application")
        
        If Err.Number = 429 Then
            Err.Clear
            Set wdapp = CreateObject("Word.Application")
        End If
        
        wdapp.Visible = True
        
        'this seems to break loop if active
        'If Dir(TEMPLATE_PATH) = "" Then
        '    MsgBox "The file " & TEMPLATE_PATH & vbCrLf & "was not found " & vbCrLf & "C:\[REMOVED BEFORE GITHUB]\.", vbExclamation, "The document does not exist."
        '    Exit Sub
        'End If
        
        wdapp.Activate
        Set wddoc = wdapp.Documents(TEMPLATE_PATH)
        If wddoc Is Nothing Then Set wddoc = wdapp.Documents.Open(TEMPLATE_PATH)
        wddoc.Activate
        
        With wdapp
            .Activate
            With .Selection.Find
                 .ClearFormatting
                 .Replacement.ClearFormatting
                 .Text = memNum_text
                 .Replacement.Text = memNumTo_text
                 .Execute Replace:=wdReplaceAll
            End With
            With .Selection.Find
                 .ClearFormatting
                 .Replacement.ClearFormatting
                 .Text = memName_text
                 .Replacement.Text = memNameTo_text
                 .Execute Replace:=wdReplaceAll
            End With
            With .Selection.Find
                 .ClearFormatting
                 .Replacement.ClearFormatting
                 .Text = memEmail_text
                 .Replacement.Text = memEmailTo_text
                 .Execute Replace:=wdReplaceAll
            End With
            With .Selection.Find
                 .ClearFormatting
                 .Replacement.ClearFormatting
                 .Text = memCreditTotal_text
                 .Replacement.Text = memCreditTotalTo_text
                 .Execute Replace:=wdReplaceAll
            End With
        End With
        
        'save new file from template associated with member
        wddoc.SaveAs Filename:=DESTINATION_PATH & memberName
        wdapp.Quit
        
        Set wddoc = Nothing
        Set wdapp = Nothing
        sFile = Dir()
    Loop
End Sub

Function FileExists(FilePath As String) As Boolean
Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

'function only used for testing, unused in final implementation
Function FindEmail(memNameTo_text As String, memEmailTo_text As String) As String

        Dim sheetTarget As Integer
        Dim wbSource As Workbook
        Dim wsSource As Worksheet
        Dim wsTarget As Worksheet
        Dim memberEmailSearch As Integer
        Dim returnEmail As String
        
        'email file to traverse
        eFile = Dir(EMAIL_SOURCE_PATH)
        
        'traverse email workbook for associated member email
        Set wbSource = Workbooks.Open(EMAIL_SOURCE_PATH)
        
        For sheetTarget = 1 To 1
            Set wsSource = wbSource.Worksheets(sheetTarget)
            
            'grab the data
            With wsTarget
                'memEmail_text = Sheets("export").Range("G1").Value
                For memberEmailSearch = 1 To 1198
                    If wsSource.Range("A" & memberEmailSearch).Text <> "" Then
                        Dim memEmailName As String
                        memEmailName = Sheets("export").Range("C" & memberEmailSearch).Value & ", " & Sheets("export").Range("B" & memberEmailSearch).Value
                        If memEmailName = memNameTo_text Then
                            returnEmail = Sheets("export").Range("A" & memberEmailSearch).Value
                            'reset sheetTarget, close the source workbook, returnEmail
                            sheetTarget = 1
                            wbSource.Close SaveChanges:=False
                        End If
                    End If
                Next memberEmailSearch
            End With
        Next sheetTarget
        FindEmail = returnEmail
End Function
