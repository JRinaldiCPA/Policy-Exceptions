Attribute VB_Name = "o2_Import_Sageworks_Data"
' Declare Workbooks
Dim wbSageworksChainedReport As Workbook

' Declare Worksheets

Dim ws5010a_LL_Dest As Worksheet
Dim ws5003_Dest As Worksheet

Option Explicit
Sub o_01_MAIN_PROCEDURE()

' Purpose:  To import the Sageworks Chained Report to create the Policy Exception Report.
' Trigger:  uf_Run_Process > cmd_Import_Chained_Report
' Updated:  5/23/2023
' Author:   James Rinaldi

' Change Log:
'       6/21/2021:  Initial Creation
'       1/3/2022:   Added the If statements around the check boxes
'       1/4/2022:   Added the code to "check" the applicable step once completed
'       1/31/2022:  Added the code to jump to the Checklist when complete
'       5/23/2023:  Added the DebugMode to NOT jump to the checklist

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency
Call o2_Import_Sageworks_Data.o_02_Assign_Private_Variables

' ---------------
' Import the data
' ---------------
    
    If uf_Run_Process.chk_Import_5003_Ticklers = True Then
        Call o2_Import_Sageworks_Data.o_21_Import_5003_FS_and_AnnualReview_Data
    End If
           
    If uf_Run_Process.chk_Import_5010a_PE = True Then
        Call o2_Import_Sageworks_Data.o_22_Import_5010a_LoanLevel_Data
    End If
    
#If DebugMode <> 1 Then
    ' Finish the process
    wsChecklist.Range("chk_o2_Import_Sageworks_Data").Value = "X"
    Application.GoTo Reference:=wsChecklist.Range("chk_o2_Import_Sageworks_Data").Offset(0, -2), Scroll:=True

#End If

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_02_Assign_Private_Variables()

' Purpose: To assign the Private Variables that were declared above the line.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 6/21/2021

' Change Log:
'       6/21/2021: Initial Creation
'       6/29/2021: Replaced the Validation check with a function

' ****************************************************************************
    
' ----------------
' Assign Variables
' ----------------
    
    ' Assign Worksheets
    
    Set ws5010a_LL_Dest = ThisWorkbook.Sheets("5010a - Policy Exceptions")
    Set ws5003_Dest = ThisWorkbook.Sheets("5003 - Ticklers & AR")
    
End Sub
Sub o_21_Import_5003_FS_and_AnnualReview_Data()

' Purpose: To import the 5003 Report Timely Receipt of FS & Annual Reviews tab from the Sageworks Chained Report.
' Trigger: Called by Main Procedure
' Updated: 6/12/2023

' Change Log:
'       6/29/2021:  Initial Creation
'       6/29/2021:  Added the "If Not XXX Is Nothing" to check if the wbSageworksChainedReport was already set
'       7/2/2021:   Added the uf_wsSelector code in case the name of the sheet changes
'       7/4/2021:   Added the code for fx_Select_Worksheet to import the worksheet
'       9/15/2021:  Switched to import the LL report and added code to fix the Header values
'       9/16/2021:  Updated with the fx_Rename_Header_Field function
'       6/28/2022:  Created the 'bolImport5001' and updated the wbClose paramater
'       6/30/2022:  Switched to just saying True to close, as we are using seperate reports
'       6/5/2023:   Added the strWsNamedRange optional varible to allow the user to update the ws Name
'       6/12/2023:  Updated to rename the field '14 Digit Account Number OR Loan Number' to 'Account Number / Loan Number'

' ****************************************************************************

' ----------------
' Assign Variables
' ----------------
    
    ' Dim Workbooks
    
    If Not wbSageworksChainedReport Is Nothing Then
        ' Do Nothing, already set
    Else
        Set wbSageworksChainedReport = fx_Open_Workbook(strPromptTitle:="Select the current (2.1) Sageworks Chained Report w/ 5003 data (Ticklers)")
    End If
        
    ' Dim Worksheets
    
    Dim ws5003_Source As Worksheet
    Set ws5003_Source = fx_Select_Worksheet( _
        strWbName:=wbSageworksChainedReport.Name, _
        strWsName:=ThisWorkbook.Worksheets(1).Evaluate("wsName_5003"), _
        strWsNamedRange:="wsName_5003")
    
    ' Dim Integers
    
    Dim intHeaderRow As Long
        intHeaderRow = 6

    ' Dim Boolean
    
    Dim bolImport5010a As Boolean
        bolImport5010a = uf_Run_Process.chk_Import_5010a_PE

' -----------------
' Fix Header Values
' -----------------

    Call fx_Rename_Header_Field( _
        ws:=ws5003_Source, _
        intHeaderRow:=intHeaderRow, _
        str_Value_To_Update:="RTB High (Branch)", _
        str_Updated_Value:="RTB High")

    Call fx_Rename_Header_Field( _
        ws:=ws5003_Source, _
        intHeaderRow:=intHeaderRow, _
        str_Value_To_Update:="RTB Low (User Defined 14)", _
        str_Updated_Value:="RTB Low")

    Call fx_Rename_Header_Field( _
        ws:=ws5003_Source, _
        intHeaderRow:=intHeaderRow, _
        str_Value_To_Update:="14 Digit Account Number OR Loan Number", _
        str_Updated_Value:="Account Number / Loan Number")

' ---------------
' Update the data
' ---------------

    Call fx_Copy_in_Data_for_Matching_Fields( _
        wsSource:=ws5003_Source, _
        wsDest:=ws5003_Dest, _
        int_Source_HeaderRow:=intHeaderRow, _
        str_ModuleName:="o_22_Import_5003_FS_and_AnnualReview_Data", _
        str_ControlTotalField:="Timely Receipt of Financial Statements - Count", _
        int_CurRow_wsValidation:=6, _
        bol_CloseSourceWb:=True)

    ' Check the validation totals
    Call fx_Validate_Control_Totals(int1stTotalRow:=6, int2ndTotalRow:=7)

End Sub
Sub o_22_Import_5010a_LoanLevel_Data()

' Purpose: To import the 5010a (LL) tab from the Sageworks Chained Report.
' Trigger: Called by Main Procedure
' Updated: 6/12/2023

' Change Log:
'       6/22/2021:  Initial Creation
'       6/23/2021:  Added the code to pass int_CurRow_wsValidation
'       6/29/2021:  Updated to allow for multiple options for the WS name
'       7/2/2021:   Updated to handle if the Sageworks Chained Report is already open
'       7/2/2021:   Added the uf_wsSelector code in case the name of the sheet changes
'       7/2/2021:   Updated the source name to be '5010a - Loan Level - Without WC'
'       7/4/2021:   Added the code for fx_Select_Worksheet to import the worksheet
'       1/31/2022:  Added the code to close the Chained Report and empty the variable if importing individually
'                   Create the bolImport5003 variable to pass to the Copy In function
'       6/5/2023:   Added the strWsNamedRange optional varible to allow the user to update the ws Name
'       6/12/2023:  Updated to rename the field '14 Digit Account Number OR Loan Number' to 'Account Number / Loan Number'

' ****************************************************************************

' ----------------
' Assign Variables
' ----------------
    
    ' Dim Workbooks
    
    Set wbSageworksChainedReport = fx_Open_Workbook(strPromptTitle:="Select the current (2.2) Sageworks Chained Report w/ 5010a data (Policy Exceptions)")
        
    ' Dim Worksheets
    
    Dim ws5010a_LL_Source As Worksheet
    Set ws5010a_LL_Source = fx_Select_Worksheet( _
        strWbName:=wbSageworksChainedReport.Name, _
        strWsName:=ThisWorkbook.Worksheets(1).Evaluate("wsName_5010a"), _
        strWsNamedRange:="wsName_5010a")

    ' Dim Integers
    
    Dim intHeaderRow As Long
        intHeaderRow = 4

    ' Dim Boolean
    
    Dim bolImport5003 As Boolean
        bolImport5003 = uf_Run_Process.chk_Import_5003_Ticklers

' -----------------
' Fix Header Values
' -----------------

    Call fx_Rename_Header_Field( _
        ws:=ws5010a_LL_Source, _
        intHeaderRow:=intHeaderRow, _
        str_Value_To_Update:="14 Digit Account Number OR Loan Number", _
        str_Updated_Value:="Account Number / Loan Number")

' -------------------------
' Import the Sageworks data
' -------------------------

    Call fx_Copy_in_Data_for_Matching_Fields( _
        wsSource:=ws5010a_LL_Source, _
        wsDest:=ws5010a_LL_Dest, _
        int_Source_HeaderRow:=intHeaderRow, _
        str_ModuleName:="o_21_Import_5010a_LoanLevel_Data", _
        str_ControlTotalField:="Risk Exposure (Loan Level)", _
        int_CurRow_wsValidation:=4, _
        bol_CloseSourceWb:=True)
        
        If uf_Run_Process.chk_Import_5003_Ticklers = False Then
            Set wbSageworksChainedReport = Nothing ' Wipe if I only imported the 5010a
        End If

    ' Check the validation totals
    Call fx_Validate_Control_Totals(int1stTotalRow:=4, int2ndTotalRow:=5)

End Sub
