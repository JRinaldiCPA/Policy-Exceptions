Attribute VB_Name = "o3_Import_BB_Data"
' Declare Worksheets

Option Explicit
Sub o_01_MAIN_PROCEDURE()

' Purpose:  To import the Business Banking BO reports to create the Policy Exception Report.
' Trigger:  uf_Run_Process > cmd_Import_BB_Exceptions
' Updated:  5/23/2023
' Author:   James Rinaldi

' Change Log:
'       10/20/2021: Initial Creation
'       11/16/2021: Created the code to import the BB Ticklers
'       1/3/2022:   Added the If statements around the check boxes
'       1/4/2022:   Added the code to "check" the applicable step once completed
'       1/31/2022:  Added the code to jump to the Checklist when complete
'       5/8/2023:   Removed the BO 3666 as a source for BB Ticklers, as we replaced it with the 5003
'       5/23/2023:  Added the DebugMode to NOT jump to the checklist

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency
Call o3_Import_BB_Data.o_02_Assign_Private_Variables

' ---------------
' Import the data
' ---------------
        
    If uf_Run_Process.chk_Import_7348_PE = True Then
        Call o3_Import_BB_Data.o_1_Import_7348_BB_Policy_Exception_Data
    End If
        
#If DebugMode <> 1 Then
        
    ' Finish the process
    wsChecklist.Range("chk_o3_Import_BB_Data").Value = "X"
    Application.GoTo Reference:=wsChecklist.Range("chk_o3_Import_BB_Data").Offset(0, -2), Scroll:=True
    
#End If

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_02_Assign_Private_Variables()

' Purpose: To assign the Private Variables that were declared above the line.
' Trigger: Called: o_01_MAIN_PROCEDURE
' Updated: 10/20/2021

' Change Log:
'       10/20/2021: Initial Creation

' ****************************************************************************
    
' ---------------------
' Assign your variables
' ---------------------
    
End Sub
Sub o_1_Import_7348_BB_Policy_Exception_Data()

' Purpose: To import data from the BO '7348 - BB Exception Exposure Report' (Policy Exceptions)
' Trigger: Called by Main Procedure
' Updated: 6/7/2023

' Change Log:
'       10/20/2021: Initial Creation
'       1/3/2022:   Updated the intHeaderRow to be dynamic, instead of using the 3rd row
'       2/22/2023:  Updated the calc for intLastRow to use Find, so that if the header info from the BO report is long it doesn't inadvertantly impact the LastCol.
'       6/7/2023:   Updated the logic for intLastCol to default to 15 if there is an issue with the source

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Dim Workbooks
    
    Dim wb7348_BB_PolicyExceptions As Workbook
    Set wb7348_BB_PolicyExceptions = fx_Open_Workbook(strPromptTitle:="Select the current (3) BO 7348 BB Policy Exception Report")
        
    ' Dim Worksheets
    
    Dim ws7348_Source As Worksheet
    Set ws7348_Source = wb7348_BB_PolicyExceptions.Sheets(1)

    Dim ws7348_Dest As Worksheet
    Set ws7348_Dest = ThisWorkbook.Sheets("7348 - BB Policy Exceptions")

    ' Dim Integers
    
    Dim intHeaderRow As Long
        intHeaderRow = ws7348_Source.Range("A:A").Find("Account Number").Row

    Dim intLastRow As Long
        intLastRow = ws7348_Source.Range("A:A").Find("Count per Loan:").Row - 1

    Dim intLastCol As Long
        On Error Resume Next
        intLastCol = ws7348_Source.Range("A3:Z99999").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        If intLastCol = 0 Then intLastCol = 15
        On Error GoTo 0
            
    Dim i As Integer
    
    Dim x As Integer: x = 1
    
    ' Dim Cell References
    
    Dim arryHeader() As Variant
        arryHeader = Application.Transpose(ws7348_Source.Range(ws7348_Source.Cells(intHeaderRow, 1), ws7348_Source.Cells(intHeaderRow, intLastCol)))

    Dim col_ExceptCode1 As Integer
        col_ExceptCode1 = fx_Create_Headers_v2("EXCEPTION CODES", arryHeader)

' -----------------
' Fix Header Values
' -----------------

    For i = col_ExceptCode1 To intLastCol
        ws7348_Source.Cells(intHeaderRow, i) = "Exception Code " & x
        x = x + 1
        
        If x > 8 Then
            MsgBox "There are more then 8 Exception Codes, need to update the headers in the Policy Exception workbook"
        End If
        
    Next i

' ------------------
' Import the BB data
' ------------------

    Call fx_Copy_in_Data_for_Matching_Fields( _
        wsSource:=ws7348_Source, _
        wsDest:=ws7348_Dest, _
        int_Source_HeaderRow:=intHeaderRow, _
        int_LastRowtoImport:=intLastRow, _
        str_ModuleName:="o3_Import_BB_Data", _
        str_ControlTotalField:="EXPOSURE", _
        int_CurRow_wsValidation:=8, _
        bol_CloseSourceWb:=True)
        
    ' Check the validation totals
    Call fx_Validate_Control_Totals(int1stTotalRow:=8, int2ndTotalRow:=9)

End Sub
