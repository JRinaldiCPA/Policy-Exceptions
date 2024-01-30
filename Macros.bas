Attribute VB_Name = "Macros"
Option Explicit
Sub Open_UserForm()

' Purpose: To open the Admin UserForm for the Sageworks Policy Exceptions.
' Trigger: Called by the Quick Access Toolbar
' Updated: 10/5/2021

' Change Log:
'       10/5/2021: Initial Creation

' ****************************************************************************
            
' -------------------
' Open the Admin form
' -------------------
    
    uf_Run_Process.Show vbModeless

End Sub
Sub Export_Data_for_Combined_Data_Set()

' Purpose: To export the data for use in the Combined Data Set project.
' Trigger: Called by the uf_Run_Process
' Updated: 7/15/2022

' Change Log:
'       7/7/2022:   Initial Creation
'       7/15/2022:  Added the code to check the checkbox

' ****************************************************************************

' -------------------------------------------
' Export the loan level Policy Exception data
' -------------------------------------------

    Call fx_Export_Data_for_RiskTrend_Database( _
        wsTarget:=ThisWorkbook.Sheets("(LL) Policy Exceptions"), _
        strFileName:=UCase(Format([v_RunDate], "MMMYY")) & "_POLICY_EXCEPTIONS", _
        dtRunDate:=CDate([v_RunDate]))

    ' Finish the process
    wsChecklist.Range("chk_o8_Export_Data").Value = "X"
    Application.GoTo Reference:=wsChecklist.Range("chk_o8_Export_Data").Offset(0, -2), Scroll:=True

End Sub
Sub ZZZ_TEST_FILTER()

' Tested on 5/23/2023

Dim str_FilterValue As String
    str_FilterValue = "Asset Based Lending, Commercial Real Estate, Middle Market Banking, Sponsor and Specialty Finance, Wealth"
    
'Dim cell As cell

'Dim rng As Range

'Dim i As Long

'Dim strTEST As String
Dim strTEST2 As String

Dim arryTEST() As Variant
    arryTEST = Application.Transpose(ThisWorkbook.Sheets("Lists").Range("ZZZ_LOBS"))


strTEST2 = Join(arryTEST, ", ") '-> This is the way
   
'Debug.Print strTEST
Debug.Print strTEST2
    
'Debug.Print ThisWorkbook.Names("ZZZ_LOBS")

'For Each rng In Range("ZZZ_LOBS")
'    Debug.Print cell
'Next rng


End Sub
Sub TEST_fx_Convert_to_Values()

Call myPrivateMacros.DisableForEfficiency

'My Timer:

Dim sTime As Double

    'Start Timer
    sTime = Timer
    
    '****************** CODE HERE **************************

    Call fx_Convert_to_Values(ws_Target:=ThisWorkbook.Worksheets("Test"), str_TargetField_Name:="Account Number / Loan Number")
    
    Debug.Print "Code took: " & (Round(Timer - sTime, 3)) & " seconds"

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
