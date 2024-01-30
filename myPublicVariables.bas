Attribute VB_Name = "myPublicVariables"

' Declare Public Strings
Public str_wsSelected As String

' Declare Worksheets
Public wsChecklist As Worksheet
Public wsLists As Worksheet
Public wsValidation As Worksheet

Option Explicit
Sub o_1_Assign_Global_Variables()

' Purpose: To declare all of the Public variables that were dimensioned "above the line".
' Trigger: Called on uf_Run_Process Initialization
' Updated: 1/31/2022

' Change Log:
'       1/31/2022:  Intial Creation

' ****************************************************************************

' ----------------
' Assign Variables
' ----------------
    
    ' Assign Sheets
    
    Set wsChecklist = ThisWorkbook.Sheets("CHECKLIST")
    Set wsLists = ThisWorkbook.Sheets("LISTS")
    Set wsValidation = ThisWorkbook.Sheets("VALIDATION")

End Sub

