Attribute VB_Name = "SF_Outlook_Log_Highlighter"
Option Explicit

Const intCol As Integer = 1

'Constants for string compares
Const logReset As String = "TIME ZONE DETAILS" 'String to compare, when new log record created - ensure UPPER CASE
Const errorFound As String = "[Event]SyncEngine status changed to Errored"


Sub OutlookSyncHighlight()
Dim rowCurrent As Integer, intColour As Integer, errorCount As Integer
Dim rowString As String

    With Application
        .ScreenUpdating = False
        .StatusBar = "Processing"
    End With
    
    'Add header row
    Sheets("LogHighlight").Activate
    Cells.Select
    Selection.Interior.ColorIndex = intColour
    
    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    Selection.Font.ColorIndex = 2
    ActiveCell.Value = "S/F Log"
    Range("A1").Interior.ColorIndex = 50
    
    
    
    
    'Start counters
    rowCurrent = 2
    errorCount = 0
    intColour = 19
    
    Do
        'Format cells, and test for string "Time zone details"
        Cells(rowCurrent, intCol).Select
        rowString = ActiveCell.Value
        
        'Change colour of row
        If InStr(UCase(rowString), UCase(logReset)) > 0 Then
            If (intColour = 19) Then 'Green
                intColour = 35
            Else  'White
                intColour = 19
            End If
        ElseIf InStr(UCase(rowString), UCase(errorFound)) > 0 Then
            intColour = 38
            errorCount = errorCount + 1
        End If
            
        Cells(rowCurrent, intCol).Interior.ColorIndex = intColour
            

        'Increment rowCurrent
        rowCurrent = rowCurrent + 1
        Cells(rowCurrent, 1).Select
        
    Loop Until IsEmpty(ActiveCell)

    Columns("A").Select
    Columns("A").EntireColumn.AutoFit
    
    With Application
        .ScreenUpdating = True
        .StatusBar = False
    End With
    
    If errorCount > 0 Then
        MsgBox "Sync errors were found & highlighed." & vbNewLine & _
            "Total number: " & errorCount, vbExclamation
    End If

End Sub




Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    
End Sub
