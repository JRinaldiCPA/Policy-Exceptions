Attribute VB_Name = "myUtility_Macros"
Option Explicit
Sub u_Wipe_Checklist()

' Purpose: To remove the X marks in the Checklist ws.
' Trigger: Button
' Updated: 8/10/2021

' Change Log:
'       8/10/2021: Initial Creation

' ****************************************************************************

myPrivateMacros.DisableForEfficiency

    ' Assign Worksheets
    
    Dim wsChecklist As Worksheet
    Set wsChecklist = ThisWorkbook.Sheets("CHECKLIST")

    ' Declare Integers
       
    Dim intLastRow As Long
        intLastRow = wsChecklist.Cells(Rows.Count, "C").End(xlUp).Row

    Dim i As Integer
    
' ------------
' Wipe the X's
' ------------

    With wsChecklist
                
        For i = 2 To intLastRow
            If .Range("C" & i).Value = "X" Then .Range("C" & i).Value = ""
        Next i
        
    End With

myPrivateMacros.DisableForEfficiencyOff

End Sub
