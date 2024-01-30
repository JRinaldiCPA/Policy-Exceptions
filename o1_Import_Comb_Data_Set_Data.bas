Attribute VB_Name = "o1_Import_Comb_Data_Set_Data"
' Declare Workbooks
Dim wbCDS As Workbook

' Declare Worksheets
Dim wsCDS_Source As Worksheet
Dim wsCDS_Dest As Worksheet

Option Explicit
Sub o_01_MAIN_PROCEDURE()

' Purpose:  To import the Combined Data Set data as the starting point for the Policy Exceptions process.
' Trigger:  uf_Run_Process
' Updated:  5/22/2023
' Author:   James Rinaldi

' Change Log:
'       6/21/2021:  Initial Creation
'       1/4/2022:   Added the code to "check" the applicable step once completed
'       1/31/2022:  Added the code to jump to the Checklist when complete
'       9/15/2022:  Updated to switch the source from the Adjusted Risk Trend to the Corporate Scorecard
'       5/22/2023:  Updated to switch the source from the Corp Scorecard to the Combined Data Set
'       5/23/2023:  Added the DebugMode to NOT jump to the checklist

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency
Call o1_Import_Comb_Data_Set_Data.o_02_Assign_Private_Variables

' --------------
' Run the import
' --------------
    
    Call o1_Import_Comb_Data_Set_Data.o_1_Import_Data
    
#If DebugMode <> 1 Then
    
    ' Finish the process
    wsChecklist.Range("chk_o1_Import_Comb_Data_Set_Data").Value = "X"
    Application.GoTo Reference:=wsChecklist.Range("chk_o1_Import_Comb_Data_Set_Data").Offset(0, -2), Scroll:=True
    
#End If
    
Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_02_Assign_Private_Variables()

' Purpose: To assign the Private Variables that were declared above the line.
' Trigger: uf_Run_Process > cmd_Import_CDS_Data
' Updated: 6/21/2021

' Change Log:
'       6/21/2021: Initial Creation
'       6/29/2021: Replaced the Validation check with a function

' ****************************************************************************
    
' ----------------
' Assign Variables
' ----------------
    
    ' Assign Worksheets
    
    Set wsCDS_Dest = ThisWorkbook.Sheets("(LL) Policy Exceptions")
    
End Sub
Sub o_1_Import_Data()

' Purpose: To import data from Combined Data Set to the (LL) Policy Exceptions.
' Trigger: Called by Main Procedure
' Updated: 6/12/2023

' Change Log:
'       6/21/2021:  Initial Creation
'       6/29/2021:  Updated to account for the Raw file being imported
'       6/29/2021:  Replaced the Count with an ISREF to pull the 1st worksheet if the given name isn't present
'       7/2/2021:   Replaced the Sageworks Risk Trend wb with the Sageworks Chained Report wb
'       7/2/2021:   Added the uf_wsSelector code in case the name of the sheet changes
'       7/4/2021:   Added code to update the headers on the Sageworks Risk Trend to be more clear
'       7/4/2021:   Replaced the code for the SageworksRiskTrend with the new fx_Select_Worksheet function
'       9/16/2021:  Updated with the fx_Rename_Header_Field function
'       1/3/2022:   Updated to switch to using the Adjusted Risk Trend as a source
'                   Switched to using the fx_Rename_Header_Fields function
'       9/15/2022:  Updated to switch the source from the Adjusted Risk Trend to the Corporate Scorecard
'       5/22/2023:  Updated to switch the source from the Corp Scorecard to the Combined Data Set
'                   Added the 'Fix Header Values' section to adjust the CDS header fields
'       6/12/2023:  Added the code to delete the Consumer Loans based on Sub-Portfolio

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Dim Workbooks
    
    Set wbCDS = fx_Open_Workbook(strPromptTitle:="Select the current (1) Combined Data Set")
        
    ' Dim Worksheets
    
    
    Set wsCDS_Source = fx_Select_Worksheet(wbCDS.Name, "COMBINED_DATA_SET")

    ' Dim Integers
    
    Dim intHeaderRow As Long
        intHeaderRow = 1

' -----------------
' Fix Header Values
' -----------------
    
    Call fx_Rename_Header_Field( _
        ws:=wsCDS_Source, _
        intHeaderRow:=intHeaderRow, _
        str_Value_To_Update:="RTB High", _
        str_Updated_Value:="Line of Business")

    Call fx_Rename_Header_Field( _
        ws:=wsCDS_Source, _
        intHeaderRow:=intHeaderRow, _
        str_Value_To_Update:="RTB Low", _
        str_Updated_Value:="Region")

    Call fx_Rename_Header_Field( _
        ws:=wsCDS_Source, _
        intHeaderRow:=intHeaderRow, _
        str_Value_To_Update:="Direct Outstanding", _
        str_Updated_Value:="Outstanding")

' -------------------------
' Delete the Consumer Loans
' -------------------------

    Call fx_Delete_Unused_Data( _
        ws:=wsCDS_Source, _
        str_Target_Field:="Sub-Portfolio", _
        str_Value_To_Delete:="Consumer, Consumer Other, Home Equity, Lending Club, Residential", _
        bol_DeleteValues_PassArray:=True)

' -------------------
' Import the CDS data
' -------------------

    Call fx_Copy_in_Data_for_Matching_Fields( _
        wsSource:=wsCDS_Source, _
        wsDest:=wsCDS_Dest, _
        int_Source_HeaderRow:=intHeaderRow, _
        str_ModuleName:="o_1_Import_Data", _
        str_ControlTotalField:="Scorecard Exposure", _
        int_CurRow_wsValidation:=2, _
        bol_CloseSourceWb:=True)

    ' Check the validation totals
    Call fx_Validate_Control_Totals(int1stTotalRow:=2, int2ndTotalRow:=3)

End Sub
