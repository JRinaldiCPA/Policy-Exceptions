Attribute VB_Name = "o5_Merge_Exception_Data"
' Declare Worksheets

Dim wsPolicyExceptions_LL As Worksheet

Dim ws5010a As Worksheet
Dim ws5003 As Worksheet
Dim ws7348 As Worksheet

Option Explicit
Sub o_01_MAIN_PROCEDURE()

' Purpose:  To merge the Policy Exception Data into the Sageworks Risk Trend.
' Trigger:  uf_Run_Process > cmd_Merge_Exception_Data
' Updated:  5/24/2023
' Author:   James Rinaldi

' Change Log:
'       6/29/2021:  Initial Creation
'       1/4/2022:   Added the msgbox to confirm the process is complete
'                   Added the code to "check" the applicable step once completed
'       1/31/2022:  Added the code to jump to the Checklist when complete
'       9/15/2022:  Removed the 'o_9_Create_Borrower_v_Loan_Exception_Flag' procedure
'       10/3/2022:  Added the code to refresh the pivots at the end of the process
'       5/23/2023:  Added the DebugMode to NOT jump to the checklist
'       5/24/2023:  Updated to use the Named Ranges in wsLists for the Borrower Level Exception str_FilterValue

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency
    
' -------------------
' Manipulate the data
' -------------------
        
    Call o5_Merge_Exception_Data.o_02_Assign_Private_Variables
    Call o5_Merge_Exception_Data.o_03_Apply_Global_Changes
    
    Call o5_Merge_Exception_Data.o_1_Import_Collateral_Data
    Call o5_Merge_Exception_Data.o_2_Import_Debt_Repay_Capacity_Data
    Call o5_Merge_Exception_Data.o_3_Import_Financial_Covenants_Data
    Call o5_Merge_Exception_Data.o_4_Import_Max_Amortization_Data
    Call o5_Merge_Exception_Data.o_5_Import_Max_Tenor_Data
    
    Call o5_Merge_Exception_Data.o_6_Import_Timely_Receipt_FS_Tickler_Data
    Call o5_Merge_Exception_Data.o_7_Import_Annual_Review_Tickler_Data
    Call o5_Merge_Exception_Data.o_8_Import_House_Limit_Exception_Data
    
#If DebugMode <> 1 Then
    
    ' Finish the process
    wsChecklist.Range("chk_o5_Merge_Exception_Data").Value = "X"
    Application.GoTo Reference:=wsChecklist.Range("chk_o5_Merge_Exception_Data").Offset(0, -2), Scroll:=True
    
#End If

    ' Refresh the pivot tables
    Call fx_Update_Named_Range("PE_Data")
    ThisWorkbook.RefreshAll

    MsgBox _
     Title:="Your good to go bro / ladybro", _
     Buttons:=vbOKOnly + vbInformation, _
     Prompt:="----------------------------------------------------------------------" & Chr(10) & Chr(10) & _
     "The Policy Exception / Tickler data was merged w/ the Risk Trend data in '(LL) Policy Exceptions'." _
     & Chr(10) & Chr(10) _
     & "You are good to click on the next step." & Chr(10) & Chr(10) & _
     "----------------------------------------------------------------------"

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_02_Assign_Private_Variables()

' Purpose: To assign the Private Variables that were declared above the line.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 5/8/2023

' Change Log:
'       6/29/2021:  Initial Creation
'       11/16/2021: Added the ws3666
'       5/8/2023:   Removed the BO 3666 as a source for BB Ticklers, as we replaced it with the 5003

' ****************************************************************************
    
' ---------------------
' Assign your variables
' ---------------------
    
    Set wsPolicyExceptions_LL = ThisWorkbook.Sheets("(LL) Policy Exceptions")
    
    Set ws5010a = ThisWorkbook.Sheets("5010a - Policy Exceptions")
    Set ws5003 = ThisWorkbook.Sheets("5003 - Ticklers & AR")
    Set ws7348 = ThisWorkbook.Sheets("7348 - BB Policy Exceptions")
    
End Sub
Sub o_03_Apply_Global_Changes()

' Purpose: To make any required changes prior to starting the process.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 6/30/2021

' Change Log:
'       6/30/2021: Initial Creation

' ****************************************************************************

' ----------------------
' Remove the Autofilters
' ----------------------
    
    If wsPolicyExceptions_LL.AutoFilterMode = True Then wsPolicyExceptions_LL.AutoFilter.ShowAllData
    If ws5010a.AutoFilterMode = True Then ws5010a.AutoFilter.ShowAllData
    If ws5003.AutoFilterMode = True Then ws5003.AutoFilter.ShowAllData
    
End Sub
Sub o_1_Import_Collateral_Data()

' Purpose: To import the '1 - Collateral & Other Support' Policy Exception data into the Sageworks Risk Trend.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 6/12/2023

' Change Log:
'       6/29/2021:  Initial Creation
'       10/21/2021: Added the code for BB / SB
'       9/15/2022:  Updated to reflect the new field names
'                   Updated to include the code to flag the balances w/ exceptions
'                   Created the 'Apply the Borrower Level Exceptions' section
'       5/24/2023:  Updated to use the Named Ranges in wsLists for the Borrower Level Exception str_FilterValue
'       6/12/2023:  Updated the field from '14 Digit Account Number' to 'Account Number / Loan Number'

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

Dim arryBLE_LOBs() As Variant
    arryBLE_LOBs = Application.Transpose(ThisWorkbook.Sheets("Lists").Range("Lists_BLE_1_Collateral"))

' --------------------------------
' Import the Policy Exception Data
' --------------------------------
    
    ' Import the Commercial Count from the 5010a
    Call fx_Count_on_Single_Field( _
        wsSource:=ws5010a, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Exception Name", _
        str_Source_MatchField:="Account Number / Loan Number", _
        str_Dest_TargetField:="1 - Collateral (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="1 - Collateral & Other Support")
    
    ' Import the Business Banking Count from the 7348
    Call fx_Count_on_Single_Field( _
        wsSource:=ws7348, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="1 - Collateral & Other Support", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="1 - Collateral (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="X")

' -----------------------------------
' Apply the Borrower Level Exceptions
' -----------------------------------

    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="1 - Collateral (#)", _
        str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="1 - Collateral (#)", _
        str_Dest_MatchField:="Helper", _
        str_FilterField_Dest:="Line of Business", _
        str_FilterValue:=Join(arryBLE_LOBs, ", "), _
        bol_FilterPassArray:=True)
        
' -------------------------------
' Calculate the Exception Balance
' -------------------------------

    Call fx_Sum_on_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="1 - Collateral (#)", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="1 - Collateral ($)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="<>", _
        bolNotBlankSum:=True, _
        strSumField:="Scorecard Exposure")
        
End Sub
Sub o_2_Import_Debt_Repay_Capacity_Data()

' Purpose: To import the '2 - Debt Repayment Capacity / Liquidity' Policy Exception data into the Sageworks Risk Trend.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 6/12/2023

' Change Log:
'       6/29/2021:  Initial Creation
'       10/21/2021: Added the code for BB / SB
'       9/15/2022:  Updated to reflect the new field names
'                   Updated to include the code to flag the balances w/ exceptions
'                   Created the 'Apply the Borrower Level Exceptions' section
'       5/24/2023:  Updated to use the Named Ranges in wsLists for the Borrower Level Exception str_FilterValue
'       6/12/2023:  Updated the field from '14 Digit Account Number' to 'Account Number / Loan Number'

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

Dim arryBLE_LOBs() As Variant
    arryBLE_LOBs = Application.Transpose(ThisWorkbook.Sheets("Lists").Range("Lists_BLE_2_DebtRepay"))

' --------------------------------
' Import the Policy Exception Data
' --------------------------------
    
    ' Import the Commercial Count from the 5010a
    Call fx_Count_on_Single_Field( _
        wsSource:=ws5010a, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Exception Name", _
        str_Source_MatchField:="Account Number / Loan Number", _
        str_Dest_TargetField:="2 - Debt Repayment / Liquidity (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="2 - Debt Repayment Capacity / Liquidity")

    ' Import the Business Banking Count from the 7348
    Call fx_Count_on_Single_Field( _
        wsSource:=ws7348, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="2 - Debt Repayment Capacity / Liquidity", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="2 - Debt Repayment / Liquidity (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="X")

' -----------------------------------
' Apply the Borrower Level Exceptions
' -----------------------------------
    
    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="2 - Debt Repayment / Liquidity (#)", _
        str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="2 - Debt Repayment / Liquidity (#)", _
        str_Dest_MatchField:="Helper", _
        str_FilterField_Dest:="Line of Business", _
        str_FilterValue:=Join(arryBLE_LOBs, ", "), _
        bol_FilterPassArray:=True)
        
' -------------------------------
' Calculate the Exception Balance
' -------------------------------

    Call fx_Sum_on_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="2 - Debt Repayment / Liquidity (#)", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="2 - Debt Repayment / Liquidity ($)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="<>", _
        bolNotBlankSum:=True, _
        strSumField:="Scorecard Exposure")

End Sub
Sub o_3_Import_Financial_Covenants_Data()

' Purpose: To import the '3 - Financial Covenants' Policy Exception data into the Sageworks Risk Trend.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 6/12/2023

' Change Log:
'       6/29/2021:  Initial Creation
'       10/21/2021: Added the code for BB / SB
'       9/15/2022:  Updated to reflect the new field names
'                   Updated to include the code to flag the balances w/ exceptions
'                   Created the 'Apply the Borrower Level Exceptions' section
'       5/24/2023:  Updated to use the Named Ranges in wsLists for the Borrower Level Exception str_FilterValue
'       6/12/2023:  Updated the field from '14 Digit Account Number' to 'Account Number / Loan Number'

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

Dim arryBLE_LOBs() As Variant
    arryBLE_LOBs = Application.Transpose(ThisWorkbook.Sheets("Lists").Range("Lists_BLE_3_FinancialCovenants"))

' --------------------------------
' Import the Policy Exception Data
' --------------------------------
    
    ' Import the Commercial Count from the 5010a
    Call fx_Count_on_Single_Field( _
        wsSource:=ws5010a, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Exception Name", _
        str_Source_MatchField:="Account Number / Loan Number", _
        str_Dest_TargetField:="3 - Financial Covenants (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="3 - Financial Covenants")

    ' Import the Business Banking Count from the 7348
    Call fx_Count_on_Single_Field( _
        wsSource:=ws7348, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="3 - Financial Covenants", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="3 - Financial Covenants (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="X")

' -----------------------------------
' Apply the Borrower Level Exceptions
' -----------------------------------
    
    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="3 - Financial Covenants (#)", _
        str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="3 - Financial Covenants (#)", _
        str_Dest_MatchField:="Helper", _
        str_FilterField_Dest:="Line of Business", _
        str_FilterValue:=Join(arryBLE_LOBs, ", "), _
        bol_FilterPassArray:=True)
        
' -------------------------------
' Calculate the Exception Balance
' -------------------------------

    Call fx_Sum_on_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="3 - Financial Covenants (#)", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="3 - Financial Covenants ($)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="<>", _
        bolNotBlankSum:=True, _
        strSumField:="Scorecard Exposure")

End Sub
Sub o_4_Import_Max_Amortization_Data()

' Purpose: To import the '4 - Maximum Amortization / Ability to Amortize' Policy Exception data into the Sageworks Risk Trend.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 6/12/2023

' Change Log:
'       6/29/2021:  Initial Creation
'       10/21/2021: Added the code for BB / SB
'       9/15/2022:  Updated to reflect the new field names
'                   Updated to include the code to flag the balances w/ exceptions
'                   Created the 'Apply the Borrower Level Exceptions' section
'       5/24/2023:  Updated to use the Named Ranges in wsLists for the Borrower Level Exception str_FilterValue
'       6/12/2023:  Updated the field from '14 Digit Account Number' to 'Account Number / Loan Number'

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

Dim arryBLE_LOBs() As Variant
    arryBLE_LOBs = Application.Transpose(ThisWorkbook.Sheets("Lists").Range("Lists_BLE_4_MaxAmort"))

' --------------------------------
' Import the Policy Exception Data
' --------------------------------
    
    ' Import the Commercial Count from the 5010a
    Call fx_Count_on_Single_Field( _
        wsSource:=ws5010a, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Exception Name", _
        str_Source_MatchField:="Account Number / Loan Number", _
        str_Dest_TargetField:="4 - Maximum Amortization (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="4 - Maximum Amortization / Ability to Amortize")
    
    ' Import the Business Banking Count from the 7348
    Call fx_Count_on_Single_Field( _
        wsSource:=ws7348, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="4 - Maximum Amortization / Ability to Amortize", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="4 - Maximum Amortization (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="X")

' -----------------------------------
' Apply the Borrower Level Exceptions
' -----------------------------------
    
    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="4 - Maximum Amortization (#)", _
        str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="4 - Maximum Amortization (#)", _
        str_Dest_MatchField:="Helper", _
        str_FilterField_Dest:="Line of Business", _
        str_FilterValue:=Join(arryBLE_LOBs, ", "), _
        bol_FilterPassArray:=True)
        
' -------------------------------
' Calculate the Exception Balance
' -------------------------------

    Call fx_Sum_on_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="4 - Maximum Amortization (#)", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="4 - Maximum Amortization ($)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="<>", _
        bolNotBlankSum:=True, _
        strSumField:="Scorecard Exposure")

End Sub
Sub o_5_Import_Max_Tenor_Data()

' Purpose: To import the '5 - Maximum Tenor' Policy Exception data into the Sageworks Risk Trend.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 6/12/2023

' Change Log:
'       6/29/2021:  Initial Creation
'       10/21/2021: Added the code for BB / SB
'       9/15/2022:  Updated to reflect the new field names
'                   Updated to include the code to flag the balances w/ exceptions
'                   Created the 'Apply the Borrower Level Exceptions' section
'       5/24/2023:  Updated to use the Named Ranges in wsLists for the Borrower Level Exception str_FilterValue
'       6/12/2023:  Updated the field from '14 Digit Account Number' to 'Account Number / Loan Number'

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

Dim arryBLE_LOBs() As Variant
    arryBLE_LOBs = Application.Transpose(ThisWorkbook.Sheets("Lists").Range("Lists_BLE_5_MaxTenor"))

' --------------------------------
' Import the Policy Exception Data
' --------------------------------
    
    ' Import the Commercial Count from the 5010a
    Call fx_Count_on_Single_Field( _
        wsSource:=ws5010a, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Exception Name", _
        str_Source_MatchField:="Account Number / Loan Number", _
        str_Dest_TargetField:="5 - Maximum Tenor     (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="5 - Maximum Tenor")
    
    ' Import the Business Banking Count from the 7348
    Call fx_Count_on_Single_Field( _
        wsSource:=ws7348, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="5 - Maximum Tenor", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="5 - Maximum Tenor     (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="X")

' -----------------------------------
' Apply the Borrower Level Exceptions
' -----------------------------------
    
    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="5 - Maximum Tenor     (#)", _
        str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="5 - Maximum Tenor     (#)", _
        str_Dest_MatchField:="Helper", _
        str_FilterField_Dest:="Line of Business", _
        str_FilterValue:=Join(arryBLE_LOBs, ", "), _
        bol_FilterPassArray:=True)
        
' -------------------------------
' Calculate the Exception Balance
' -------------------------------

    Call fx_Sum_on_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="5 - Maximum Tenor     (#)", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="5 - Maximum Tenor     ($)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="<>", _
        bolNotBlankSum:=True, _
        strSumField:="Scorecard Exposure")

End Sub
Sub o_6_Import_Timely_Receipt_FS_Tickler_Data()

' Purpose: To import the 'Timely Receipt of Financials' Tickler data into the Sageworks Risk Trend.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 6/12/2023

' Change Log:
'       6/30/2021:  Initial Creation
'       11/16/2021: Updated to include the 3666 data
'       9/15/2022:  Updated to reflect the new field names
'                   Updated to include the code to flag the balances w/ exceptions
'                   Created the 'Apply the Borrower Level Exceptions' section
'       5/8/2023:   Removed the BB related logic to delete those borrowers, as we are now using the 5003 as the source not the BO 3666
'       6/12/2023:  Updated the field from '14 Digit Account Number' to 'Account Number / Loan Number'

' ****************************************************************************

' -----------------------
' Import the Tickler Data
' -----------------------
    
    Call fx_Update_Single_Field( _
        wsSource:=ws5003, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Timely Receipt of Financial Statements - Count", _
        str_Source_MatchField:="Account Number / Loan Number", _
        str_Dest_TargetField:="Timely Receipt of Financials (#)", _
        str_Dest_MatchField:="Account Number", _
        bol_SkipDuplicates:=True)

' -----------------------------------
' Apply the Borrower Level Exceptions
' -----------------------------------
    
    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Timely Receipt of Financials (#)", _
        str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="Timely Receipt of Financials (#)", _
        str_Dest_MatchField:="Helper")
        
' -------------------------------
' Calculate the Exception Balance
' -------------------------------

    Call fx_Sum_on_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Timely Receipt of Financials (#)", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="Timely Receipt of Financials ($)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="1", _
        strSumField:="Scorecard Exposure")
        
End Sub
Sub o_7_Import_Annual_Review_Tickler_Data()

' Purpose: To import the 'Annual Review' Tickler data into the Sageworks Risk Trend.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 6/12/2023

' Change Log:
'       6/30/2021: Initial Creation
'       10/28/2021: Switched to matching on Account Number, now that I am using the Risk Trend
'       11/16/2021: Updated to include the 3666 data
'       9/15/2022:  Updated to reflect the new field names
'                   Updated to include the code to flag the balances w/ exceptions
'                   Created the 'Apply the Borrower Level Exceptions' section
'       5/8/2023:   Removed the BB related logic to delete those borrowers, as we are now using the 5003 as the source not the BO 3666
'       6/12/2023:  Updated the field from '14 Digit Account Number' to 'Account Number / Loan Number'

' ****************************************************************************

' -----------------------
' Import the Tickler Data
' -----------------------
    
    Call fx_Update_Single_Field( _
        wsSource:=ws5003, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Annual Review - Count", _
        str_Source_MatchField:="Account Number / Loan Number", _
        str_Dest_TargetField:="Annual Review (#)", _
        str_Dest_MatchField:="Account Number", _
        bol_SkipDuplicates:=True)

' -----------------------------------
' Apply the Borrower Level Exceptions
' -----------------------------------
    
    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Annual Review (#)", _
        str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="Annual Review (#)", _
        str_Dest_MatchField:="Helper")
        
' -------------------------------
' Calculate the Exception Balance
' -------------------------------

    Call fx_Sum_on_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Annual Review (#)", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="Annual Review ($)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="<>", _
        bolNotBlankSum:=True, _
        strSumField:="Scorecard Exposure")

End Sub
Sub o_8_Import_House_Limit_Exception_Data()

' Purpose: To import the 'House Limit Exposure (HLE) Breach' Policy Exception data into the Sageworks Risk Trend.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 6/12/2023

' Change Log:
'       6/29/2021: Initial Creation
'       9/15/2022:  Updated to reflect the new field names
'                   Updated to include the code to flag the balances w/ exceptions
'                   Created the 'Apply the Borrower Level Exceptions' section
'       6/12/2023:  Updated the field from '14 Digit Account Number' to 'Account Number / Loan Number'

' ****************************************************************************

' --------------------------------
' Import the Policy Exception Data
' --------------------------------
    
    Call fx_Count_on_Single_Field( _
        wsSource:=ws5010a, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Exception Name", _
        str_Source_MatchField:="Account Number / Loan Number", _
        str_Dest_TargetField:="House Limit Exposure (HLE) Breach (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="House Limit Exposure (HLE) Breach")

' -----------------------------------
' Apply the Borrower Level Exceptions
' -----------------------------------
    
    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="House Limit Exposure (HLE) Breach (#)", _
        str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="House Limit Exposure (HLE) Breach (#)", _
        str_Dest_MatchField:="Helper")
        
' -------------------------------
' Calculate the Exception Balance
' -------------------------------

    Call fx_Sum_on_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="House Limit Exposure (HLE) Breach (#)", _
        str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="House Limit Exposure (HLE) Breach ($)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="<>", _
        bolNotBlankSum:=True, _
        strSumField:="Scorecard Exposure")

End Sub
