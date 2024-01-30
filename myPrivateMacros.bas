Attribute VB_Name = "myPrivateMacros"
Option Explicit
Sub DisableForEfficiency()

' -----------------------------------------
' Turns off functionality to speed up Excel
' -----------------------------------------

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

End Sub
Sub DisableForEfficiencyOff()

' -----------------------------------------
' Turns functionality back on
' -----------------------------------------

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True

End Sub
Sub Timer_Code()

'My Timer:

Dim sTime As Double

    'Start Timer
    sTime = Timer
    
    '****************** CODE HERE **************************
    
    Debug.Print "Code took: " & (Round(Timer - sTime, 3)) & " seconds"


