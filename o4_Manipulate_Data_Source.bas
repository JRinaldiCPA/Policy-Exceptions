Attribute VB_Name = "o4_Manipulate_Data_Source"
' Declare Worksheets

Dim wsPolicyExcept_LL As Worksheet
Dim ws5010a As Worksheet
Dim ws5003 As Worksheet

Dim ws7348 As Worksheet

Option Explicit
Sub o_01_MAIN_PROCEDURE()

' Purpose:  To manipulate the data sources being used to create the Sageworks Policy Exception Report.
' Trigger: uf_Run_Process > cmd_Manipulate_Data_Sources
' Updated:  7/12/2023
' Author:   James Rinaldi

' Change Log:
'       6/29/2021:  Initial Creation
'       1/4/2022:   Added the msgbox to confirm the process is complete
'                   Added the code to "check" the applicable step once completed
'       1/31/2022:  Added the code to jump to the Checklist when complete
'       5/23/2023:  Added the DebugMode to NOT jump to the checklist
'       7/12/2023:  Added o_5_Create_Borrower_Reviewed_Field_in_CDS_Data

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency
    
' -------------------
' Manipulate the data
' -------------------
        
    Call o4_Manipulate_Data_Source.o_02_Assign_Private_Variables
        
    Call o4_Manipulate_Data_Source.o_1_Manipulate_CDS_Data
        
    Call o4_Manipulate_Data_Source.o_2_Manipulate_5010a_LoanLevel_Data
    Call o4_Manipulate_Data_Source.o_3_Manipulate_5003_FS_and_AnnualReview_Data
    Call o4_Manipulate_Data_Source.o_4_Manipulate_7348_BB_PolicyException_Data
    
    Call o4_Manipulate_Data_Source.o_5_Create_Borrower_Reviewed_Field_in_CDS_Data
        
#If DebugMode <> 1 Then
    
    ' Finish the process
    wsChecklist.Range("chk_o4_Manipulate_Data_Source").Value = "X"
    Application.GoTo Reference:=wsChecklist.Range("chk_o4_Manipulate_Data_Source").Offset(0, -2), Scroll:=True
    
#End If
    
    'MsgBox "The data sources have been manuipulated.  You are good to click on the next step."
    
       MsgBox _
        Title:="Your good to go bro / ladybro", _
        Buttons:=vbOKOnly + vbInformation, _
        Prompt:="----------------------------------------------------------------------" & Chr(10) & Chr(10) & _
        "The data sources have been manuipulated.  You are good to click on the next step." & Chr(10) & Chr(10) & _
        "----------------------------------------------------------------------"

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_02_Assign_Private_Variables()

' Purpose: To assign the Private Variables that were declared above the line.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 5/8/2023

' Change Log:
'       6/29/2021:  Initial Creation
'       10/20/2021: Added the BB worksheets
'       5/8/2023:   Removed the BO 3666 as a source for BB Ticklers, as we replaced it with the 5003

' ****************************************************************************
    
' ----------------
' Assign Variables
' ----------------
    
    Set wsPolicyExcept_LL = ThisWorkbook.Sheets("(LL) Policy Exceptions")
    
    Set ws5010a = ThisWorkbook.Sheets("5010a - Policy Exceptions")
    Set ws5003 = ThisWorkbook.Sheets("5003 - Ticklers & AR")
    Set ws7348 = ThisWorkbook.Sheets("7348 - BB Policy Exceptions")
    
End Sub
Sub o_1_Manipulate_CDS_Data()

' Purpose: To manipulate the data imported from the Combined Data Set
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 7/12/2023

' Change Log:
'       6/30/2021:  Initial Creation
'       9/15/2021:  Added the code to remove the BB and SB customers
'       10/5/2021:  Added the code to convert "Private Bank" => Wealth
'       10/5/2021:  Added the code to remove the Webster internal loans
'       10/21/2021: Commented out the code to remove the Business Banking & Small Business Customers
'       10/28/2021: Added the code to Update Small Business => Business Banking
'       1/8/2022:   Update the code to sort to be more dynamic by using the Column References
'       4/5/2022:   Replaced the 'fx_Rename_SmallBusiness_to_BusinessBanking' and 'fx_Rename_PrivateBank_to_Wealth' with 'fx_Convert_LOB_Name'
'                   Added the code to Convert CRE > EF => Equipment Finance
'       5/18/2022:  Turned the code back on to convert PSF -> MM
'       9/15/2022:  Added the code to create the Region v2 from 'o7_Create_Policy_Exceptions_ws'
'                   Removed the 'CRE => EF' code, as this is no longer needed
'       5/5/2023:   Removed old code about Wealth as it's been updated in the source data
'       5/24/2023:  Removed the code related to Region v2
'                   Added the code to remove the Webster Internal Loans
'                   Added the code to remove the PPP loans from L-WBS BB
'       6/12/2023:  Added the code to update the LOB to be 'Equipment Finance' based on a Sub-Portfolio of 'Equipment Finance' for L-SNB
'       6/13/2023:  Added the code to create the PE Update Portfolio and PE Update Sub-Portfolio fields
'       7/5/2023:   Added code to convert the Account Number to values, to normalize across all data sets
'       7/12/2023:  Added code to apply the "Borrower Reviewed" field
'                   Broke the "Borrower Reviewed" code out to prevent any unexpected issues from reordering the subs
'                   Added code to remove the '-' in the EF loans before converting to values to match the format of the Sageworks data

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Dim Worksheets
    
    Dim wsData As Worksheet
    Set wsData = wsPolicyExcept_LL
    
    ' Dim Integers
    
    Dim intLastCol As Long
        intLastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim intLastRow As Long
        intLastRow = Application.WorksheetFunction.Max( _
        wsData.Cells(Rows.Count, "A").End(xlUp).Row, _
        wsData.Cells(Rows.Count, "D").End(xlUp).Row)
        
    Dim i As Long
    
    ' Dim Cell References
    
    Dim arryHeader() As Variant
        arryHeader = Application.Transpose( _
        wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))
        
    Dim col_LOB As Long
        col_LOB = fx_Create_Headers_v2("Line of Business", arryHeader)

    Dim col_Region As Long
        col_Region = fx_Create_Headers_v2("Region", arryHeader)

    Dim col_Borrower As Long
        col_Borrower = fx_Create_Headers_v2("Borrower Name", arryHeader)
            
    Dim col_AcctNum As Long
        col_AcctNum = fx_Create_Headers_v2("Account Number", arryHeader)

    Dim col_Helper As Long
        col_Helper = fx_Create_Headers_v2("Helper", arryHeader)
        
    Dim col_SubPort As Long
        col_SubPort = fx_Create_Headers_v2("Sub-Portfolio", arryHeader)

' ------------------------------------------
' Convert Small Business => Business Banking
' ------------------------------------------
        
    Call fx_Convert_LOB_Name( _
        rngToUpdate:=wsData.Range(wsData.Cells(2, col_LOB), wsData.Cells(intLastRow, col_LOB)), _
        bolSmallBusiness_to_BusinessBanking:=True)

' ----------------------------------------------------
' Convert Middle Market (PSF) => Public Sector Finance
' ----------------------------------------------------
        
    Call fx_Convert_LOB_Name( _
        rngToUpdate:=wsData.Range(wsData.Cells(2, col_LOB), wsData.Cells(intLastRow, col_LOB)), _
        rngRegion:=wsData.Range(wsData.Cells(2, col_Region), wsData.Cells(intLastRow, col_Region)), _
        bolPublicSectorFinance:=True)
        
' ----------------------------------------------------
' Convert LOB to Equipment Finance using Sub Portfolio
' ----------------------------------------------------
        
    Call fx_Convert_LOB_Name( _
        rngToUpdate:=wsData.Range(wsData.Cells(2, col_LOB), wsData.Cells(intLastRow, col_LOB)), _
        rngSubPortfolio:=wsData.Range(wsData.Cells(2, col_SubPort), wsData.Cells(intLastRow, col_SubPort)), _
        bol_EF_Using_SubPort:=True)
        
' ------------------------
' Normalize Account Number
' ------------------------
    
    ' Remove the '-' from the EF loans to match Sageworks
    
    For i = 2 To intLastRow
        If wsData.Cells(i, col_LOB).Value2 = "Equipment Finance" Then
            wsData.Cells(i, col_AcctNum).Value2 = Replace(wsData.Cells(i, col_AcctNum).Value2, "-", "")
        End If
    Next i

    Call fx_Convert_to_Values(ws_Target:=wsData, str_TargetField_Name:="Account Number")
        
' -------------------
' Create Helper field
' -------------------

    For i = 2 To intLastRow
        wsData.Cells(i, col_Helper).Value2 = wsData.Cells(i, col_LOB).Value2 & " - " & wsData.Cells(i, col_Borrower).Value2
    Next i

' --------------------
' Remove the PPP Loans
' --------------------

    For i = intLastRow To 2 Step -1
    
        If wsData.Cells(i, col_Region) = "01761 P3 Portfolio" Then
            wsData.Rows(i).EntireRow.Delete
        ElseIf wsData.Cells(i, col_Region) = "00161 SB Central" And wsData.Cells(i, col_SubPort) = "" Then
            wsData.Rows(i).EntireRow.Delete
        End If
    
    Next i

' ---------------------------------
' Remove the Webster Internal Loans
' ---------------------------------

    Call fx_Remove_Webster_Internal_Loans(wsTarget:=wsData, strAccountNumber:="Account Number")

' ----------------------------------------------------------------------
' Pull in the 'PE Update Portfolio' and 'PE Update Sub-Portfolio' fields
' ----------------------------------------------------------------------

    Call fx_Update_Single_Field( _
        wsSource:=wsLists, wsDest:=wsData, _
        str_Source_TargetField:="3. Updated Portfolio", str_Source_MatchField:="3. Sub-Portfolio", _
        str_Dest_TargetField:="PE Updated Portfolio", str_Dest_MatchField:="Sub-Portfolio", _
        bol_MissingLookupData_MsgBox:=True, _
        strWsNameLookup:="LISTS")
        
    Call fx_Update_Single_Field( _
        wsSource:=wsLists, wsDest:=wsData, _
        str_Source_TargetField:="3. Updated Sub-Portfolio", str_Source_MatchField:="3. Sub-Portfolio", _
        str_Dest_TargetField:="PE Updated Sub-Portfolio", str_Dest_MatchField:="Sub-Portfolio", _
        bol_MissingLookupData_MsgBox:=True, _
        strWsNameLookup:="LISTS")
    
' -------------
' Sort the data
' -------------

With wsData

    .UsedRange.Sort _
        Key1:=.Cells(1, col_LOB), Order1:=xlAscending, _
        Key2:=.Cells(1, col_Region), Order2:=xlAscending, _
        Key3:=.Cells(1, col_Borrower), Order3:=xlAscending, _
        Header:=xlYes

End With

End Sub
Sub o_2_Manipulate_5010a_LoanLevel_Data()

' Purpose: To manipulate the 5010(a) Loan Level data.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 7/5/2023

' Change Log:
'       6/29/2021:  Initial Creation
'       6/30/2021:  Added a sort by Account Number to aid in merging the data, and switched to the Sort method
'       7/3/2021:   Added the section to remove the Paid Off loans
'       9/15/2021:  Normalized the code to delete the Paid Off loans and Proposed Borrowers
'       10/5/2021:  Added the code to convert "Private Bank" => Wealth
'       1/4/2022:   Temporarily disabled the Loan Paid Off code
'       1/8/2022:   Update the code to sort to be more dynamic by using the Column References
'       5/5/2023:   Replaced the 'fx_Rename_PrivateBank_to_Wealth' code with 'fx_Convert_LOB_Name'
'                   Added the msgBoxes for Matt when all the data is deleted
'       6/12/2023:  Updated to rename the field '14 Digit Account Number' to 'Account Number / Loan Number'
'       7/5/2023:   Added code to convert the Account Number to values, to normalize across all data sets'
'                   Updated the 'Remove the obselete exceptions' to not remove the "No Exceptions" values

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Dim Worksheets
    
    Dim wsData As Worksheet
    Set wsData = ws5010a
    
    ' Dim Integers
    
    Dim intLastCol As Long
        intLastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim intLastRow As Long
        intLastRow = Application.WorksheetFunction.Max( _
        wsData.Cells(Rows.Count, "D").End(xlUp).Row, _
        wsData.Cells(Rows.Count, "E").End(xlUp).Row)
        
        If intLastRow = 1 Then Exit Sub
        
    Dim i As Long
    
    ' Dim Cell References
    
    Dim arryHeader() As Variant
        arryHeader = Application.Transpose( _
        wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))

    Dim col_RiskExp As Long
        col_RiskExp = fx_Create_Headers_v2("Risk Exposure (Loan Level)", arryHeader)
        
    Dim col_Except_Status
        col_Except_Status = fx_Create_Headers_v2("Exception Status", arryHeader)
            
    Dim col_Except_Name
        col_Except_Name = fx_Create_Headers_v2("Exception Name", arryHeader)
            
    Dim col_PaidOff
        col_PaidOff = fx_Create_Headers_v2("Loan Paid Off?", arryHeader)
            
    Dim col_LOB As Long
        col_LOB = fx_Create_Headers_v2("RTB High", arryHeader)
            
    Dim col_Region As Long
        col_Region = fx_Create_Headers_v2("RTB Low", arryHeader)
            
    Dim col_Borrower As Long
        col_Borrower = fx_Create_Headers_v2("Customer Name", arryHeader)
            
    Dim col_AcctNum As Long
        col_AcctNum = fx_Create_Headers_v2("Account Number / Loan Number", arryHeader)
            
    ' Dim Arrays
    Dim arryCurExcept As Variant
        arryCurExcept = Get_arryCurExceptions
               
' ------------------------
' Normalize Account Number
' ------------------------

    Call fx_Convert_to_Values(ws_Target:=wsData, str_TargetField_Name:="Account Number / Loan Number")
               
' -------------------------
' Delete the Paid Off loans
' -------------------------

    Call fx_Delete_Unused_Data( _
        ws:=wsData, _
        str_Target_Field:="Loan Paid Off?", _
        str_Value_To_Delete:="Paid Off")
        
    ' Reset the intLastRow and check to make sure there is data left
    intLastRow = Application.WorksheetFunction.Max(wsData.Cells(Rows.Count, "D").End(xlUp).Row, wsData.Cells(Rows.Count, "E").End(xlUp).Row)
    
    If intLastRow = 1 Then
        MsgBox "Warning, all of the data was deleted when removing Paid Off loans from the 5010a data."
        Exit Sub
    End If
        
' -----------------------------
' Remove the Proposed Borrowers
' -----------------------------

    Call fx_Delete_Unused_Data( _
        ws:=wsData, _
        str_Target_Field:="Risk Exposure (Loan Level)", _
        str_Value_To_Delete:="")

    intLastRow = wsData.Cells(Rows.Count, "D").End(xlUp).Row ' Reset Last Row

    ' Reset the intLastRow and check to make sure there is data left
    intLastRow = Application.WorksheetFunction.Max(wsData.Cells(Rows.Count, "D").End(xlUp).Row, wsData.Cells(Rows.Count, "E").End(xlUp).Row)
    
    If intLastRow = 1 Then
        MsgBox "Warning, all of the data was deleted when removing Proposed Borrowers from the 5010a data."
        Exit Sub
    End If

' ------------------------------
' Remove the obselete exceptions
' ------------------------------

    For i = intLastRow To 2 Step -1
    
        If wsData.Cells(i, col_Except_Status) <> "Active" And wsData.Cells(i, col_Except_Status) <> "Retroactively Applied" Then
            wsData.Rows(i).EntireRow.Delete
        End If
        
        ' 7/5/2023: Disabled to allow the "No Exceptions" to pass through, and there are no longer "obselete" exceptions included
'        If IsError(Application.Match(wsData.Cells(i, col_Except_Name), arryCurExcept, 0)) Then
'            wsData.Rows(i).EntireRow.Delete
'        End If

    Next i

' -----------------------------
' Update Private Bank => Wealth
' -----------------------------
        
    Call fx_Convert_LOB_Name( _
        rngToUpdate:=wsData.Range(wsData.Cells(2, col_LOB), wsData.Cells(intLastRow, col_LOB)), _
        bolPrivateBank_to_Wealth:=True)

' -------------
' Sort the data
' -------------

With wsData.Sort

    .SortFields.Clear
    .SortFields.Add Key:=Cells(1, col_LOB), Order:=xlAscending
    .SortFields.Add Key:=Cells(1, col_Region), Order:=xlAscending
    .SortFields.Add Key:=Cells(1, col_Borrower), Order:=xlAscending
    .SortFields.Add Key:=Cells(1, col_AcctNum), Order:=xlAscending

    .Header = xlYes
    .SetRange Range(wsData.Cells(1, 1), wsData.Cells(intLastRow, intLastCol))
    .Apply

End With

End Sub
Sub o_3_Manipulate_5003_FS_and_AnnualReview_Data()

' Purpose: To manipulate the 5003 Report Timely Receipt of FS & Annual Reviews data.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 7/12/2023

' Change Log:
'       6/29/2021:  Initial Creation
'       6/30/2021:  Added the code to create the Helper field
'       6/30/2021:  Added the code to remove the 0s for the Count
'       6/30/2021:  Added the code to address the BB customers incorrectly showing as <> BB
'       7/3/2021:   Added 'Trimm Inc' to list to delete and switched from ElseIf to Or
'       9/15/2021:  Normalized the code to delete the BB customers
'       9/15/2021:  Disabled the columns related to balances, since we don't use those
'       10/5/2021:  Added the code to convert "Private Bank" => Wealth
'       10/5/2021:  Added the code to delete 'Small Business' customers, if they appear
'       10/5/2021:  Removed the code for Proposed Borrowers, as we are no longer importing the balance fields after switching from BL to LL
'       1/8/2022:   Update the code to sort to be more dynamic by using the Column References
'       5/5/2023:   Replaced the 'fx_Rename_PrivateBank_to_Wealth' code with 'fx_Convert_LOB_Name'
'       5/8/2023:   Removed the BB related logic to delete those borrowers, as we are now using the 5003 as the source not the BO 3666
'       5/23/2023:  Removed the code to delete the BB borrowers w/ an innacurate LOB, as were not purging the BB data anymore (ex. North Cottage Program Inc)
'       7/5/2023:   Added code to convert the Account Number to values, to normalize across all data sets
'       7/12/2023:  Removed the no longer used Helper field code

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Dim Worksheets
    
    Dim wsData As Worksheet
    Set wsData = ws5003
    
    ' Dim Integers
    
    Dim intLastCol As Long
        intLastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim intLastRow As Long
        intLastRow = Application.WorksheetFunction.Max( _
        wsData.Cells(Rows.Count, "A").End(xlUp).Row, _
        wsData.Cells(Rows.Count, "D").End(xlUp).Row)
        
    Dim i As Long
    
    ' Dim "Ranges"
    
    Dim arryHeader() As Variant
        arryHeader = Application.Transpose( _
        wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))

    Dim col_FS_RiskExp As Long
        col_FS_RiskExp = fx_Create_Headers_v2("Timely Receipt of Financial Statements - Risk Exposure", arryHeader)
        
    Dim col_AR_RiskExp As Long
        col_AR_RiskExp = fx_Create_Headers_v2("Annual Review - Risk Exposure", arryHeader)
        
    Dim col_FS_Count As Long
        col_FS_Count = fx_Create_Headers_v2("Timely Receipt of Financial Statements - Count", arryHeader)
        
    Dim col_AR_Count As Long
        col_AR_Count = fx_Create_Headers_v2("Annual Review - Count", arryHeader)
        
    Dim col_LOB As Long
        col_LOB = fx_Create_Headers_v2("RTB High", arryHeader)
        
    Dim col_Region As Long
        col_Region = fx_Create_Headers_v2("RTB Low", arryHeader)

    Dim col_Borrower As Long
        col_Borrower = fx_Create_Headers_v2("Company Name", arryHeader)

'    Dim col_Helper As Long
'        col_Helper = fx_Create_Headers_v2("Helper", arryHeader)
        
' ------------------------
' Normalize Account Number
' ------------------------

    Call fx_Convert_to_Values(ws_Target:=wsData, str_TargetField_Name:="Account Number / Loan Number")

' ----------------------------
' Overwrite the N/As and Zeros
' ----------------------------

    For i = 2 To intLastRow

        ' Financial Statement - Count
        If wsData.Cells(i, col_FS_Count) = 0 Then
            wsData.Cells(i, col_FS_Count) = ""
        End If
    
        ' Annual Review - Count
        If wsData.Cells(i, col_AR_Count) = 0 Then
            wsData.Cells(i, col_AR_Count) = ""
        End If
    
    Next i

' -------------------------------
' Fix the Equipment Finance Dept.
' -------------------------------

    For i = 2 To intLastRow
    
        If wsData.Cells(i, col_Region) = "770 Equipment Finance" Then
            wsData.Cells(i, col_LOB).Value2 = "Equipment Finance"
        End If
    
    Next i

' -----------------------------
' Update Private Bank => Wealth
' -----------------------------
        
    Call fx_Convert_LOB_Name( _
        rngToUpdate:=wsData.Range(wsData.Cells(2, col_LOB), wsData.Cells(intLastRow, col_LOB)), _
        bolPrivateBank_to_Wealth:=True)

' -------------------
' Create Helper field
' -------------------

'    For i = 2 To intLastRow
'        wsData.Cells(i, col_Helper).Value2 = wsData.Cells(i, col_LOB).Value2 & " - " & wsData.Cells(i, col_Borrower).Value2
'    Next i

' -------------
' Sort the data
' -------------

With wsData

    .UsedRange.Sort _
        Key1:=.Cells(1, col_LOB), Order1:=xlAscending, _
        Key2:=.Cells(1, col_Region), Order2:=xlAscending, _
        Key3:=.Cells(1, col_Borrower), Order3:=xlAscending, _
        Header:=xlYes

End With

End Sub
Sub o_4_Manipulate_7348_BB_PolicyException_Data()

' Purpose: To manipulate the '7348 - BB Policy Exceptions' data.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 7/5/2023

' Change Log:
'       10/20/2021: Initial Creation
'       5/24/2023:  Added the 'arryHeader_wsLists' and related wsList variables to be dynamic
'       7/5/2023:   Added code to convert the Account Number to values, to normalize across all data sets

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Declare Worksheets
    
    Dim wsData As Worksheet
    Set wsData = ws7348
    
    Dim wsLists As Worksheet
    Set wsLists = ThisWorkbook.Sheets("Lists")
    
    ' Declare Integers
    
    Dim intLastCol As Long
        intLastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim intLastRow As Long
        intLastRow = Application.WorksheetFunction.Max( _
        wsData.Cells(Rows.Count, "A").End(xlUp).Row, _
        wsData.Cells(Rows.Count, "B").End(xlUp).Row)
        
    ' Declare Loop Variables
        
    Dim i As Long
    
    Dim x As Long
    
    ' Declare Strings
    
    Dim strException As String
            
    ' Declare Dictionary
    
    Dim dict_BB_Exceptions As Scripting.Dictionary
    Set dict_BB_Exceptions = New Scripting.Dictionary
    
        dict_BB_Exceptions.CompareMode = TextCompare
    
    ' Declare "Ranges"
    
    Dim arryHeader() As Variant
        arryHeader = Application.Transpose( _
        wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))

    ' Declare Exception Code Columns

    Dim col_Exception_1 As Long
        col_Exception_1 = fx_Create_Headers_v2("Exception Code 1", arryHeader)
        
    Dim col_Exception_8 As Long
        col_Exception_8 = fx_Create_Headers_v2("Exception Code 8", arryHeader)
        
    ' Declare Policy Exception Columns
    
    Dim col_1_Collat_Excpt As Long
        col_1_Collat_Excpt = fx_Create_Headers_v2("1 - Collateral & Other Support", arryHeader)
    
    Dim col_2_DebtRepay_Excpt As Long
        col_2_DebtRepay_Excpt = fx_Create_Headers_v2("2 - Debt Repayment Capacity / Liquidity", arryHeader)
        
    Dim col_3_FinCovs_Excpt As Long
        col_3_FinCovs_Excpt = fx_Create_Headers_v2("3 - Financial Covenants", arryHeader)
        
    Dim col_4_MaxAmort_Excpt As Long
        col_4_MaxAmort_Excpt = fx_Create_Headers_v2("4 - Maximum Amortization / Ability to Amortize", arryHeader)
        
    Dim col_5_MaxTenor_Excpt As Long
        col_5_MaxTenor_Excpt = fx_Create_Headers_v2("5 - Maximum Tenor", arryHeader)
    
    ' Dim wsList Variables
    
    Dim arryHeader_wsLists() As Variant
        arryHeader_wsLists = Application.Transpose( _
        wsLists.Range(wsLists.Cells(1, 1), wsLists.Cells(1, 999)))

    Dim col_wsLists_ExcpCode As Long
        col_wsLists_ExcpCode = fx_Create_Headers_v2("2. Exception Code", arryHeader_wsLists)
        
    Dim col_wsLists_ExcpType As Long
        col_wsLists_ExcpType = fx_Create_Headers_v2("2. Exception Type", arryHeader_wsLists)
        
    Dim intLastRow_wsLists As Long
        intLastRow_wsLists = wsLists.Cells(Rows.Count, col_wsLists_ExcpCode).End(xlUp).Row
        
' ------------------------
' Normalize Account Number
' ------------------------

    Call fx_Convert_to_Values(ws_Target:=wsData, str_TargetField_Name:="Account Number")
        
' ---------------------------------
' Fill the BB Exceptions Dictionary
' ---------------------------------
    
    For i = 2 To intLastRow_wsLists
        dict_BB_Exceptions.Add Key:=CStr(wsLists.Cells(i, col_wsLists_ExcpCode).Value2), Item:=CStr(wsLists.Cells(i, col_wsLists_ExcpType).Value2)
    Next i
    
' -------------------------------------------------------
' Flag the Policy Exceptions based on the Exception Codes
' -------------------------------------------------------

    For i = 2 To intLastRow
    
        For x = col_Exception_1 To col_Exception_8
    
            If wsData.Cells(i, x) <> "" Then ' Only run the code if the Exception column isn't blank
    
                If dict_BB_Exceptions.Exists(CStr(wsData.Cells(i, x).Value2)) Then
                    strException = dict_BB_Exceptions.Item(CStr(wsData.Cells(i, x).Value2))
                    
                    Select Case strException
                        Case "Collateral and Other Support"
                            wsData.Cells(i, col_1_Collat_Excpt) = "X"
                        Case "Debt Repayment Capacity / Liquidity"
                            wsData.Cells(i, col_2_DebtRepay_Excpt) = "X"
                        Case "Financial Covenants"
                            wsData.Cells(i, col_3_FinCovs_Excpt) = "X"
                        Case "Maximum Amortization / Ability to Amortize"
                            wsData.Cells(i, col_4_MaxAmort_Excpt) = "X"
                        Case "Maximum Tenor"
                            wsData.Cells(i, col_5_MaxTenor_Excpt) = "X"
                    End Select
                    
                End If
    
            End If
            
        Next x
    
    Next i

End Sub
Sub o_5_Create_Borrower_Reviewed_Field_in_CDS_Data()

' Purpose: To create the Borrower Reviewed data w/in the Combined Data Set
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 7/12/2023

' Change Log:
'       7/12/2023:  Initial Creation, broke the "Borrower Reviewed" code out from the 'o_1_Manipulate_CDS_Data' sub

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Dim Worksheets
    
    Dim wsData As Worksheet
    Set wsData = wsPolicyExcept_LL
    
    ' Dim Integers
    
    Dim intLastCol As Long
        intLastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim intLastRow As Long
        intLastRow = Application.WorksheetFunction.Max( _
        wsData.Cells(Rows.Count, "A").End(xlUp).Row, _
        wsData.Cells(Rows.Count, "D").End(xlUp).Row)
        
    Dim i As Long
    
    ' Dim Cell References
    
    Dim arryHeader() As Variant
        arryHeader = Application.Transpose( _
        wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))
        
    Dim col_LegacyBank As Long
        col_LegacyBank = fx_Create_Headers_v2("Legacy Bank Name", arryHeader)
        
    Dim col_Helper As Long
        col_Helper = fx_Create_Headers_v2("Helper", arryHeader)
        
    Dim col_BorrowerReviewed As Long
        col_BorrowerReviewed = fx_Create_Headers_v2("Borrower Reviewed", arryHeader)

' ---------------------------------------------
' Create the 'Borrower Reviewed' data for L-WBS
' ---------------------------------------------

    ' Clear the old data just in case the manipulation is re-run
    wsData.Range(wsData.Cells(2, col_BorrowerReviewed), wsData.Cells(intLastRow, col_BorrowerReviewed)).Value2 = ""

    For i = 2 To intLastRow
    
        If wsData.Cells(i, col_LegacyBank) = "WBS" Then
            wsData.Cells(i, col_BorrowerReviewed).Value2 = "Yes"
        End If
    
    Next i
        
' ---------------------------------------------
' Create the 'Borrower Reviewed' data for L-SNB
' ---------------------------------------------
        
    ' Pull Big Commercial Policy Exception data from the 5010a
    
    Call fx_Update_Single_Field( _
        wsSource:=ws5010a, wsDest:=wsPolicyExcept_LL, _
        str_Source_TargetField:="Exception Name", str_Source_MatchField:="Account Number / Loan Number", _
        str_Dest_TargetField:="Borrower Reviewed", str_Dest_MatchField:="Account Number")
    
    ' Pull All Commercial Tickler data from the 5003
    
    Call fx_Update_Single_Field( _
        wsSource:=ws5003, wsDest:=wsPolicyExcept_LL, _
        str_Source_TargetField:="Company Name", str_Source_MatchField:="Account Number / Loan Number", _
        str_Dest_TargetField:="Borrower Reviewed", str_Dest_MatchField:="Account Number")
    
    ' Pull Small Commercial Policy Exception data from the 7348
             
    Call fx_Update_Single_Field( _
        wsSource:=ws7348, wsDest:=wsPolicyExcept_LL, _
        str_Source_TargetField:="NAME FULL", str_Source_MatchField:="Account Number", _
        str_Dest_TargetField:="Borrower Reviewed", str_Dest_MatchField:="Account Number")
             
' --------------------------------------------------------------------
' Apply the 'Borrower Reviewed' data from Loan level to Borrower level
' --------------------------------------------------------------------
             
    ' Replace all values pulled in with a Yes
    For i = 2 To intLastRow
    
        If wsData.Cells(i, col_BorrowerReviewed) <> "" Then
            wsData.Cells(i, col_BorrowerReviewed).Value2 = "Yes"
        End If
    
    Next i
    
    ' Apply to the Borrower level based on Helper
       
    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExcept_LL, wsDest:=wsPolicyExcept_LL, _
        str_Source_TargetField:="Borrower Reviewed", str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="Borrower Reviewed", str_Dest_MatchField:="Helper")

End Sub

