Attribute VB_Name = "myFunctions"
Option Explicit
Function Global_Error_Handling(SubName, ErrSource, ErrNum, ErrDesc)

    Dim strTempVer As String
       strTempVer = Mid(String:=ThisWorkbook.Name, Start:=InStr(ThisWorkbook.Name, " (v") + 3, Length:=4)

If Err.Number <> 0 Then MsgBox _
    Title:="I am Error", _
    Buttons:=vbCritical, _
    Prompt:="Something went awry, try to hit Cancel and redo the last step. " _
    & "If that doesn't resolve it then reach out to James Rinaldi for a fix. " _
    & "This tool has a growth mindset, with each issue addressed we itterate to a better version." & Chr(10) & Chr(10) _
    & "Please take a screenshot of this message, and send it to James." & Chr(10) _
    & "Include a brief description of what you were doing when it occurred." & Chr(10) & Chr(10) _
    & "Error Source: " & ErrSource & " " & strTempVer & Chr(10) _
    & "Subroutine: " & SubName & Chr(10) _
    & "Error Desc.: #" & ErrNum & " - " & ErrDesc & Chr(10)
    
    myPrivateMacros.DisableForEfficiencyOff
    
    End

'Or include all of the details in an auto email to me and just prompt them for what happened.

End Function
Function fx_Create_Headers(strHeaderTitle As String, arry_Header As Variant)

' Purpose: To determine the column number for a specific title in the header.
' Trigger: Called
' Updated: 12/11/2020

' Change Log:
'       5/1/2020: Initial Creation
'       12/11/2020: Updated to use an array instead of the range, reducing the time to run by 75%.

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim i As Integer

' ----------------------
' Loop through the array
' ----------------------

    For i = LBound(arry_Header) To UBound(arry_Header)
        If arry_Header(i, 1) = strHeaderTitle Then
            fx_Create_Headers = i
            Exit Function
        End If
    Next i

End Function
Function fx_Create_Headers_v2(str_Target_FieldName As String, arry_Target_Header As Variant) As Long

' Purpose: To determine the column number for a specific title in the header.
' Trigger: Called
' Updated: 7/3/2023

' Change Log:
'       5/1/2020: Intial Creation
'       12/11/2020: Updated to use an array instead of the range, reducing the time to run by 75%.
'       7/3/2023:   Switched to using a Dictionary which is ~50% faster, and easier to troubleshoot

' --------------------------------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    intColNum_Source = fx_Create_Headers_v2(str_Target_FieldName, arry_Target_Header)

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim i As Long

    Dim dict_LookupData As Scripting.Dictionary
    Set dict_LookupData = New Scripting.Dictionary
        dict_LookupData.CompareMode = TextCompare
        
    Dim arryTemp() As Variant
        arryTemp = WorksheetFunction.Transpose(arry_Target_Header)

' ----------------------------------------
' Fill the Dictionary with the Lookup Data
' ----------------------------------------

On Error Resume Next
    For i = LBound(arryTemp) To UBound(arryTemp)
        dict_LookupData.Add Key:=arryTemp(i), Item:=i
    Next i
On Error GoTo 0

' --------------------------------------------------
' Loop through the array to find the matching column
' --------------------------------------------------

    If dict_LookupData.Exists(str_Target_FieldName) Then
        fx_Create_Headers_v2 = dict_LookupData.Item(str_Target_FieldName)
    End If

End Function
Function fx_Copy_in_Data_for_Matching_Fields(wsSource As Worksheet, wsDest As Worksheet, _
 Optional int_Source_HeaderRow As Long, Optional int_LastRowtoImport As Long, _
 Optional bol_CloseSourceWb As Boolean, Optional bol_ImportVisibleFieldsOnly As Boolean, _
 Optional str_ModuleName As String, Optional str_ControlTotalField As String, Optional int_CurRow_wsValidation As Long)
    
' Purpose: To copy the data from the source to the destination, wherever the fields match.

' Trigger: Called
' Updated: 5/23/2023
'
' Change Log:
'       9/18/2020:  Intial Creation based on CV Mod Agg. CV Tracker import code
'       11/3/2020:  Updated to include the strSourceDesc and strDestDesc to feed into the validation function
'       11/3/2020:  Updated to include the str_ModuleName to feed into the validation function
'       11/3/2020:  Removed the 'DisableforEfficiency' as it was disabiling it in my Main Procedure.
'       2/12/2021:  Updated to account for pulling in only visible data
'       2/12/2021:  Switched from the filtered boolean to int_LastRowtoImport
'       2/12/2021:  Updated to use Arrays instead of Ranges for the import
'       2/12/2021:  Added the code related to bol_CloseSourceWb
'       2/25/2021:  Updated the code for str_ControlTotalField and str_ModuleName to make it optional
'       2/25/2021:  Removed the old col_Bal_Dest and col_Bal_Source references
'       3/9/2021:   Added the code to use intLastUsedRow_Dest to delete any extraneous rows
'       3/15/2021:  Updated to include the Optional intHeaderRow field to handle ignoring headers
'       5/17/2021:  Updated the code related to the Data Validation to use the intFirstRowData_Source instead of defaulting to 1
'       5/17/2021:  Updated the intHeaderRow variable to use the code from intHeaderRow_Source
'       6/16/2021:  Made some minor improvements to the variables to make them more resilient.
'       6/16/2021:  Added the code to apply the formatting from the first row to the rest.
'       6/21/2021:  Added Option to only import the visible fields
'       6/22/2021:  Updated the code to assign intHeaderRow_Source, and related code in the data
'       3/1/2022:   Added the int_CurRow_wsValidation variable to pass to fx_Create_Data_Validation_Control_Totals
'       5/5/2023:   Added a check for intLastUsedRow_Dest so that if it is 1 it changes to 2
'       5/23/2023:  Attempted to use an array but it failed when pulling in the L-SNB Acct #s (LIQ accounts starting in "=")
'                   Switched to using Copy => Paste for errors in the import

' ********************************************************************************************************************************************************

'   USE EXAMPLE: _
        Call fx_Copy_in_Data_for_Matching_Fields( _
            wsSource:=wsSource, _
            wsDest:=wsData, _
            int_Source_HeaderRow:=1, _
            int_LastRowtoImport:=0, _
            str_ModuleName:="o_11_Import_Data", _
            str_ControlTotalField:="New Direct Outstanding", _
            bol_CloseSourceWb:=True, _
            int_CurRow_wsValidation:=2, _
            bol_ImportVisibleFieldsOnly:=True)

' LEGEND MANDATORY:
'   wsSource:  The Source worksheet that the data is being copied FROM
'   wsDest:  The destination worksheet that the data is being copied TO

' LEGEND OPTIONAL:
'   int_Source_HeaderRow: The row that the header data is located, default is 1
'   int_LastRowtoImport: The last row of data to import, default is the max of LastRow for Col A, Col B, and Col C
'   bol_CloseSourceWb: Closes the parent workbook of the wsSource, so long as it isn't the workbook running the code
'   bol_ImportVisibleFieldsOnly: If set to True then imports only visibile fields in from the wsSource
'   str_ModuleName: Used to pass the module name of the module running the code to fx_Create_Data_Validation_Control_Totals
'   str_ControlTotalField: Used to pass the control total fiel name to fx_Create_Data_Validation_Control_Totals
'   int_CurRow_wsValidation: Used to pass the current row for the wsValidation to fx_Create_Data_Validation_Control_Totals


' ********************************************************************************************************************************************************

' -----------------------------------------------
' Turn off any filtering from the source and dest
' -----------------------------------------------
        
    If wsSource.AutoFilterMode = True Then wsSource.AutoFilter.ShowAllData
        
    If wsDest.AutoFilterMode = True Then wsDest.AutoFilter.ShowAllData

' -----------------
' Declare Variables
' -----------------

    'Dim "Source" Integers
    
    Dim int_LastRow_Source As Long
        If int_LastRowtoImport > 0 Then 'If I passed the int_LastRowtoImport variable use it
            int_LastRow_Source = int_LastRowtoImport
        Else
            int_LastRow_Source = WorksheetFunction.Max( _
            wsSource.Cells(Rows.Count, "A").End(xlUp).Row, _
            wsSource.Cells(Rows.Count, "B").End(xlUp).Row, _
            wsSource.Cells(Rows.Count, "C").End(xlUp).Row)
        End If

    Dim intHeaderRow_Source As Long
        
        If int_Source_HeaderRow > 0 Then 'If I passed the intHeaderRow variable use it
            intHeaderRow_Source = int_Source_HeaderRow
        Else
            intHeaderRow_Source = 1
        End If

    Dim intFirstRowData_Source As Long
        intFirstRowData_Source = intHeaderRow_Source + 1

    Dim int_LastCol_Source As Long
        int_LastCol_Source = WorksheetFunction.Max( _
        wsSource.Cells(intHeaderRow_Source, Columns.Count).End(xlToLeft).Column, _
        wsSource.Rows(intHeaderRow_Source).Find("").Column - 1)
        
    'Dim "Dest" Integers

    Dim int_LastRow_Dest As Long
        int_LastRow_Dest = WorksheetFunction.Max( _
        wsDest.Cells(Rows.Count, "A").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "B").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "C").End(xlUp).Row)
        
        If int_LastRow_Dest = 1 Then int_LastRow_Dest = 2
    
    Dim intLastUsedRow_Dest As Long
        intLastUsedRow_Dest = wsDest.Range("A1").SpecialCells(xlCellTypeLastCell).Row
        
        If intLastUsedRow_Dest = 1 Then intLastUsedRow_Dest = 2 ' Added 5/5/2023
    
    Dim int_LastCol_Dest As Long
        int_LastCol_Dest = wsDest.Cells(1, Columns.Count).End(xlToLeft).Column
        int_LastCol_Dest = WorksheetFunction.Max( _
        wsDest.Cells(1, Columns.Count).End(xlToLeft).Column, _
        wsDest.Rows(1).Find("").Column - 1)
    
    'Dim Other Integers
        
    Dim int_CurRowValidation As Long
        
    'Dim Ranges / "Ranges"
    
    Dim arry_Header_Source() As Variant
        arry_Header_Source = Application.Transpose( _
        wsSource.Range(wsSource.Cells(intHeaderRow_Source, 1), wsSource.Cells(intHeaderRow_Source, int_LastCol_Source)))
        
    Dim arry_Header_Dest() As Variant
        arry_Header_Dest = Application.Transpose( _
        wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(1, int_LastCol_Dest)))
        
    Dim arryTEST() As Variant
    
    'Dim Values for Loops
        
    Dim strFieldName As String
    
    Dim intColNum_Source As Long
    
    Dim intColNum_Dest As Long
        
    Dim i As Long
    
    'Dim Strings
    
    Dim strSourceDesc As String
        strSourceDesc = wsSource.Parent.Name & " - " & wsSource.Name

    Dim strDestDesc As String
        strDestDesc = wsDest.Parent.Name & " - " & wsDest.Name

' ----------------------------------
' Copy over the data from the source
' ----------------------------------
        
    'Clear out the old data and cell fill
    wsDest.Range(wsDest.Cells(2, 1), wsDest.Cells(int_LastRow_Dest, int_LastCol_Dest)).ClearContents
    wsDest.Range(wsDest.Cells(2, 1), wsDest.Cells(int_LastRow_Dest, int_LastCol_Dest)).Interior.Color = xlNone
    wsDest.Range(wsDest.Cells(int_LastRow_Dest + 1, 1), wsDest.Cells(intLastUsedRow_Dest, 1)).EntireRow.Delete

    'Loop through the fields in the destination
    For intColNum_Dest = 1 To int_LastCol_Dest
        strFieldName = wsDest.Cells(1, intColNum_Dest).Value2
        intColNum_Source = fx_Create_Headers(strFieldName, arry_Header_Source)
        
        If intColNum_Source > 0 Then
            If bol_ImportVisibleFieldsOnly = True Then
                If wsDest.Columns(intColNum_Dest).Hidden = False Then
                    wsDest.Range(wsDest.Cells(2, intColNum_Dest), wsDest.Cells(int_LastRow_Source - intHeaderRow_Source + 1, intColNum_Dest)).Value2 = _
                    wsSource.Range(wsSource.Cells(intFirstRowData_Source, intColNum_Source), wsSource.Cells(int_LastRow_Source, intColNum_Source)).Value2
                End If
            Else
                On Error Resume Next
                wsDest.Range(wsDest.Cells(2, intColNum_Dest), wsDest.Cells(int_LastRow_Source - intHeaderRow_Source + 1, intColNum_Dest)).Value2 = _
                wsSource.Range(wsSource.Cells(intFirstRowData_Source, intColNum_Source), wsSource.Cells(int_LastRow_Source, intColNum_Source)).Value2
                
                If Err.Number = 7 Then 'Handle the LIQ Loan # errors
                    wsSource.Range(wsSource.Cells(intFirstRowData_Source, intColNum_Source), wsSource.Cells(int_LastRow_Source, intColNum_Source)).Copy
                    wsDest.Range(wsDest.Cells(2, intColNum_Dest), wsDest.Cells(int_LastRow_Source - intHeaderRow_Source + 1, intColNum_Dest)).PasteSpecial xlPasteValues
                        
                    On Error GoTo 0
                End If
                                        
            End If
        End If
        
    Next intColNum_Dest

    'Reset the Last Row variable
    int_LastRow_Dest = WorksheetFunction.Max( _
        wsDest.Cells(Rows.Count, "A").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "B").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "C").End(xlUp).Row)

' ------------------------------------------------------------
' Output the control totals to the Validation ws, if it exists
' ------------------------------------------------------------
    
If str_ControlTotalField <> "" Then
    
    Call fx_Create_Data_Validation_Control_Totals( _
        wsDataSource:=wsSource, _
        str_ModuleName:=str_ModuleName, _
        strSourceName:=strSourceDesc, _
        intHeaderRow:=intHeaderRow_Source, _
        int_LastRowtoImport:=int_LastRowtoImport, _
        str_ControlTotalField:=str_ControlTotalField, _
        int_CurRow_wsValidation:=int_CurRow_wsValidation)
    
    Call fx_Create_Data_Validation_Control_Totals( _
        wsDataSource:=wsDest, _
        str_ModuleName:=str_ModuleName, _
        strSourceName:=strDestDesc, _
        intHeaderRow:=1, _
        str_ControlTotalField:=str_ControlTotalField, _
        int_CurRow_wsValidation:=int_CurRow_wsValidation + 1)
    
End If
    
' -----------------------------------------------
' Apply the formatting to all rows from the first
' -----------------------------------------------
    
    Call fx_Steal_First_Row_Formating( _
        ws:=wsDest, _
        intFirstRow:=2, _
        int_LastRow:=int_LastRow_Dest, _
        int_LastCol:=int_LastCol_Dest)
    
' -------------------------
' Close the Source workbook
' -------------------------
    
    If wsSource.Parent.Name <> ThisWorkbook.Name And bol_CloseSourceWb = True Then
        wsSource.Parent.Close savechanges:=False
    End If

End Function
Function fx_Create_Data_Validation_Control_Totals(wsDataSource As Worksheet, str_ModuleName As String, strSourceName As String, intHeaderRow As Long, str_ControlTotalField As String, Optional int_LastRowtoImport As Long, Optional int_CurRow_wsValidation As Long, Optional dblTotalsFromSource As Double, Optional intRecordCountFromSource As Long, Optional dblAdjustControlTotal As Double)

' Purpose: To output the data validation control totals to the wsValidation, if it exists.
' Trigger: Called
' Updated: 1/5/2022

' Use Example: _
    Call fx_Create_Data_Validation_Control_Totals( _
        wsDataSource:=wsDest, _
        str_ModuleName:=str_ModuleName, _
        strSourceName:=strDestDesc, _
        intHeaderRow:=1, _
        str_ControlTotalField:=str_ControlTotalField, _
        int_CurRow_wsValidation:=3)

' Use Example 2: _
    Call fx_Create_Data_Validation_Control_Totals( _
    wsDataSource:=ws3666_Source, _
    str_ModuleName:="o3_Import_BB_Data", _
    strSourceName:=ws3666_Source.Parent.Name & " - " & ws3666_Source.Name, _
    intHeaderRow:=intHeaderRow, _
    str_ControlTotalField:="Line Commitment", _
    int_CurRow_wsValidation:=10, _
    intRecordCountFromSource:=intDestRowCounter - 2, _
    dblTotalsFromSource:=dblControlTotal)

' Change Log:
'       9/26/2020: Initial Creation
'       11/3/2020: Updated to activate ThisWorkbook before checking for the Validation ws
'       12/19/2020: Made the intRecordCount more resiliant
'       12/19/2020: Added the ThisWorkbook.Name to the ISREF check
'       2/12/2021: Added the code for int_LastRowtoImport
'       5/17/2021: Updated the calculation for strRng_Totals to go down to intRecordCount + intHeaderRow
'       6/16/2021: Added the code related to int_CurRow_wsValidation
'       6/22/2021: Updated the code for intRecordCount
'       6/22/2021: Updated the Check for the Validation worksheet to reference ThisWorkbook, and avoid the .Activate
'       1/5/2022:  Added the code to bypass the sums and whatnot if dblTotalsFromSource or intRecordCountFromSource is passed
'                   Added the dblAdjustControlTotal code
    
' ****************************************************************************

    ' Only run of the VALIDATION ws exists
    If Evaluate("ISREF(" & "'[" & ThisWorkbook.Name & "]" & "Validation'" & "!A1)") = False Then
        Debug.Print "fx_Create_Data_Validation_Control_Totals failed becuase there is no ws called 'VALIDATION' in the Workbook"
        Exit Function
    End If

' ----------------------------
' Declare Validation Variables
' ----------------------------

    'Dim Worksheets

    Dim wsValidation As Worksheet
    Set wsValidation = ThisWorkbook.Sheets("VALIDATION")

    Dim wsSource As Worksheet
    Set wsSource = wsDataSource

    ' Dim Cell References

    Dim int_LastCol As Integer
        int_LastCol = wsSource.Cells(intHeaderRow, Columns.Count).End(xlToLeft).Column
      
    Dim intCurRow As Long
        If int_CurRow_wsValidation > 1 Then ' If I passed the int_CurRow_wsValidation variable use it
            intCurRow = int_CurRow_wsValidation
        Else
            intCurRow = wsValidation.Cells(Rows.Count, "A").End(xlUp).Row + 1
        End If

    ' Bypass the code if dblTotals was passed
    If dblTotalsFromSource > 0 Or intRecordCountFromSource > 0 Then GoTo Bypass

' ------------------------
' Declare Source Variables
' ------------------------

    'Dim "Ranges"
    
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(wsSource.Range(wsSource.Cells(intHeaderRow, 1), wsSource.Cells(intHeaderRow, int_LastCol)))
        
    Dim intColTotals As Integer
        intColTotals = fx_Create_Headers_v2(str_ControlTotalField, arry_Header)
    
    'Dim Integers

    Dim intRecordCount As Long
        If int_LastRowtoImport > 0 Then ' If I passed the int_LastRowtoImport variable use it
            intRecordCount = int_LastRowtoImport - intHeaderRow
        Else
            intRecordCount = WorksheetFunction.Max( _
            wsSource.Cells(Rows.Count, "A").End(xlUp).Row, _
            wsSource.Cells(Rows.Count, "B").End(xlUp).Row, _
            wsSource.Cells(Rows.Count, "C").End(xlUp).Row) - intHeaderRow
        End If
    
    'Dim Other Variables

    Dim strCol_Totals As String
        strCol_Totals = Split(Cells(1, intColTotals).Address, "$")(1)
    
    Dim strRng_Totals As String
        strRng_Totals = strCol_Totals & "1:" & strCol_Totals & intRecordCount + intHeaderRow
        
    Dim dblTotals As Double
        dblTotals = Round(Application.WorksheetFunction.Sum(wsSource.Range(strRng_Totals)), 2) - dblAdjustControlTotal

Bypass:

    ' Assign the Optional variables if they were passed
    If dblTotalsFromSource > 0 Then dblTotals = dblTotalsFromSource
    If intRecordCountFromSource > 0 Then intRecordCount = intRecordCountFromSource

' ------------------------------------------------------
' Output the validation totals from the passed variables
' ------------------------------------------------------

    With wsValidation
        .Range("A" & intCurRow) = Format(Now, "m/d/yyyy hh:mm")   'Date / Time
        .Range("B" & intCurRow) = str_ModuleName                   'Code Module
        .Range("C" & intCurRow) = strSourceName                   'Source
        .Range("D" & intCurRow) = Format(dblTotals, "$#,##0")     'Total
        .Range("E" & intCurRow) = Format(intRecordCount, "0,0")   'Count
    End With

End Function
Function fx_Validate_Control_Totals(int1stTotalRow As Long, int2ndTotalRow As Long)

' Purpose: To validate that the control totals for the data imported match.
' Trigger: Called
' Updated: 6/8/2022

' Use Example: _
    Call fx_Validate_Control_Totals(int1stTotalRow:=2, int2ndTotalRow:=3)

' Change Log:
'       6/22/2021:  Intial Creation
'       1/31/2022:  Updated to switch to an Information msgbox
'       6/8/2022:   Updated the control total formatting for the count not matching

' ****************************************************************************

    ' Only run of the VALIDATION ws exists
    If Evaluate("ISREF(" & "'[" & ThisWorkbook.Name & "]" & "Validation'" & "!A1)") = False Then
        Debug.Print "fx_Validate_Control_Totals failed becuase there is no ws called 'VALIDATION' in the Workbook"
        Exit Function
    End If

' -----------
' Declare your variables
' -----------
    
    ' Dim Worksheets
    
    Dim wsValidation As Worksheet
    Set wsValidation = ThisWorkbook.Sheets("VALIDATION")
    
    ' Dim Integers
    
    Dim int1stTotal As Double
        int1stTotal = wsValidation.Cells(int1stTotalRow, "D").Value
    
    Dim int2ndTotal As Double
        int2ndTotal = wsValidation.Cells(int2ndTotalRow, "D").Value
    
    Dim int1stCount As Long
        int1stCount = wsValidation.Cells(int1stTotalRow, "E").Value
    
    Dim int2ndCount As Long
        int2ndCount = wsValidation.Cells(int2ndTotalRow, "E").Value
        
    ' Dim Booleans
    
    Dim bolTotalsMatch As Boolean
        If int1stTotal = int2ndTotal And int1stCount = int2ndCount Then
            bolTotalsMatch = True
        Else
            bolTotalsMatch = False
        End If
    
' -----------
' Output the messagebox with the results
' -----------
   
    If bolTotalsMatch = True Then
    MsgBox Title:="Validation Totals Match", _
        Buttons:=vbOKOnly + vbInformation, _
        Prompt:="The validation totals match, you're golden. " & Chr(10) & Chr(10) _
        & "1st Validation Total: " & Format(int1stTotal, "$#,##0") & Chr(10) _
        & "1st Validation Count: " & Format(int1stCount, "0,0")
       
    ElseIf bolTotalsMatch = False Then
    MsgBox Title:="Validation Totals Don't Match", _
        Buttons:=vbCritical, _
        Prompt:="The validation totals from the workbook don't match what was imported. " _
        & "Please review the totals in the Validation worksheet to determine what went awry. " & Chr(10) & Chr(10) _
        & "1st Validation Total Variance: " & Format(int1stTotal - int2ndTotal, "$#,##0") & Chr(10) & Chr(10) _
        & "1st Validation Count Variance: " & Format(int1stCount - int2ndCount, "0,0")

    End If
    
End Function
Function fx_Count_on_Single_Field(wsSource As Worksheet, wsDest As Worksheet, str_Source_TargetField As String, str_Source_MatchField As String, str_Dest_TargetField As String, str_Dest_MatchField As String, str_Criteria As String, Optional bol_CloseSourceWb As Boolean)

' Purpose: To import the summed / counted data from the Source to the Destination.
    ' 1) Add the data to be referenced into the dictionary
    ' 2) Look for the data to be updated for matches to the reference dictionary
    ' 3) Replace the existing data with the reference data

' Trigger: Called
' Updated: 5/5/2023

' Change Log:
'       6/29/2021:  Initial Creation, based on fx_Update_Single_Field
'       6/29/2021:  Added the 'And arry_Target_Source(i - 1) = str_Criteria' to ensure we were comparing the same Exception AND the same Account Number
'       11/16/2021: Added the strSumField
'       9/15/2022:  Split out the Count and Sum functions to simplify
'       5/5/2023:   Updated so that the int_LastRow_wsSource is a minimum of 2

' ********************************************************************************************************************************************************

' Use Example: _
    Call fx_Count_on_Single_Field( _
        wsSource:=ws5010a, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Exception Name", _
        str_Source_MatchField:="14 Digit Account Number", _
        str_Dest_TargetField:="1 - Collateral (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="1 - Collateral & Other Support")

' LEGEND MANDATORY:
'   TBD:

' LEGEND OPTIONAL:
'   bolNotBlankSum: When the value in the terget isn't blank then do the Sum

' ****************************************************************************

' -------------------------------
' Declare your wsSource variables
' -------------------------------
        
With wsSource
        
    ' Dim Integers
    
    Dim int_LastCol_wsSource As Long
        int_LastCol_wsSource = .Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim int_LastRow_wsSource As Long
        int_LastRow_wsSource = WorksheetFunction.Max( _
        .Cells(Rows.Count, "A").End(xlUp).Row, _
        .Cells(Rows.Count, "B").End(xlUp).Row, _
        .Cells(Rows.Count, "C").End(xlUp).Row)
        
        If int_LastRow_wsSource = 1 Then int_LastRow_wsSource = 2
        
    ' Dim "Ranges"
        
    Dim arry_Header_wsSource() As Variant
        arry_Header_wsSource = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol_wsSource)))
        
    Dim col_Target_Source As Integer
        col_Target_Source = fx_Create_Headers_v2(str_Source_TargetField, arry_Header_wsSource)

    Dim col_Match_Source As Integer
        col_Match_Source = fx_Create_Headers_v2(str_Source_MatchField, arry_Header_wsSource)
        
    ' Dim Arrays
    
    Dim arry_Target_Source() As Variant
        arry_Target_Source = Application.Transpose(.Range(.Cells(1, col_Target_Source), .Cells(int_LastRow_wsSource, col_Target_Source)))
    
    Dim arry_Match_Source() As Variant
        arry_Match_Source = Application.Transpose(.Range(.Cells(1, col_Match_Source), .Cells(int_LastRow_wsSource, col_Match_Source)))
        
End With
        
' -----------------------------
' Declare your wsDest variables
' -----------------------------
        
With wsDest
        
    ' Dim wsDest Integers
    
    Dim int_LastCol_wsDest As Long
        int_LastCol_wsDest = wsDest.Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim int_LastRow_wsDest As Long
        int_LastRow_wsDest = WorksheetFunction.Max( _
        wsDest.Cells(Rows.Count, "A").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "B").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "C").End(xlUp).Row)
 
    ' Dim wsDest "Ranges"
    
    Dim arry_Header_wsDest() As Variant
        arry_Header_wsDest = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol_wsDest)))
        
    Dim col_Target_Dest As Integer
        col_Target_Dest = fx_Create_Headers_v2(str_Dest_TargetField, arry_Header_wsDest)

    Dim col_Match_Dest As Integer
        col_Match_Dest = fx_Create_Headers_v2(str_Dest_MatchField, arry_Header_wsDest)
 
    ' Dim Ranges
        
    Dim rng_Target_Dest As Range
    Set rng_Target_Dest = .Range(.Cells(1, col_Target_Dest), .Cells(int_LastRow_wsDest, col_Target_Dest))
        
    ' Dim Arrays
    
    Dim arry_Target_Dest() As Variant
        arry_Target_Dest = Application.Transpose(.Range(.Cells(1, col_Target_Dest), .Cells(int_LastRow_wsDest, col_Target_Dest)))
    
    Dim arry_Match_Dest() As Variant
        arry_Match_Dest = Application.Transpose(.Range(.Cells(1, col_Match_Dest), .Cells(int_LastRow_wsDest, col_Match_Dest)))
 
End With
        
' ----------------------------
' Declare your Other variables
' ----------------------------
        
    ' Dim Dictionaries
    
    Dim dict_LookupData As Scripting.Dictionary
        Set dict_LookupData = New Scripting.Dictionary
        dict_LookupData.CompareMode = TextCompare
        
    ' Declare Loop Variables
    
    Dim i As Long
        
    Dim intCount As Long

' ----------------------------------------
' Fill the Dictionary with the Lookup Data
' ----------------------------------------
    
On Error Resume Next
       
        For i = 2 To UBound(arry_Target_Source)
            If arry_Target_Source(i) = str_Criteria Then
                If arry_Match_Source(i) = arry_Match_Source(i - 1) And arry_Target_Source(i - 1) = str_Criteria Then
                    intCount = intCount + 1
                Else
                    intCount = intCount + 1
                    dict_LookupData.Add Key:=arry_Match_Source(i), Item:=intCount
                    intCount = 0
                End If
            
                ' To handle the last record
                If i = UBound(arry_Target_Source) Then
                    intCount = intCount + 1
                    dict_LookupData.Add Key:=arry_Match_Source(i), Item:=intCount
                    intCount = 0
                End If
            
            End If
        
        Next i
        
On Error GoTo 0
    
' -----------------------------------------------------------------------------
' Loop through the Lookup Data to pull in the updated data for the target field
' -----------------------------------------------------------------------------

    For i = 2 To UBound(arry_Match_Dest)
        If dict_LookupData.Exists(arry_Match_Dest(i)) Then
            arry_Target_Dest(i) = dict_LookupData.Item(arry_Match_Dest(i))
        End If
    Next i

    ' Output the values from the array
    rng_Target_Dest.Value2 = Application.Transpose(arry_Target_Dest)
    
    ' Empty the Dictionary
    dict_LookupData.RemoveAll

' -------------------------
' Close the Source workbook
' -------------------------
    
    If wsSource.Parent.Name <> ThisWorkbook.Name And bol_CloseSourceWb = True Then
        wsSource.Parent.Close savechanges:=False
    End If

End Function
Function fx_Sum_on_Single_Field(wsSource As Worksheet, wsDest As Worksheet, str_Source_TargetField As String, str_Source_MatchField As String, str_Dest_TargetField As String, str_Dest_MatchField As String, str_Criteria As String, Optional strSumField As String, Optional bolNotBlankSum As Boolean, Optional bol_CloseSourceWb As Boolean)

' Purpose: To import the summed / counted data from the Source to the Destination.
    ' 1) Add the data to be referenced into the dictionary
    ' 2) Look for the data to be updated for matches to the reference dictionary
    ' 3) Replace the existing data with the reference data

' Trigger: Called
' Updated: 11/16/2021

' Change Log:
'       6/29/2021: Initial Creation, based on fx_Update_Single_Field
'       6/29/2021: Added the 'And arry_Target_Source(i - 1) = str_Criteria' to ensure _
                    we were comparing the same Exception AND the same Account Number
'       11/16/2021: Added the strSumField
'       9/15/2022:  Created the bolNotBlankSum boolean to allow me to sum when not blank

' ********************************************************************************************************************************************************

' Use Example: _
    Call fx_Sum_on_Single_Field( _
        wsSource:=ws5010a_LL, wsDest:=wsSageworksRT, _
        str_Source_TargetField:="Exception Name", _
        str_Source_MatchField:="14 Digit Account Number", _
        str_Dest_TargetField:="1 - Collateral & Other Support", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="1 - Collateral & Other Support", _
        strSumField:="Risk Exposure (Loan Level)")

' LEGEND MANDATORY:
'   TBD:

' LEGEND OPTIONAL:
'   bolNotBlankSum: When the value in the terget isn't blank then do the Sum

' ****************************************************************************

' -------------------------------
' Declare your wsSource variables
' -------------------------------
        
With wsSource
        
    ' Dim Integers
    
    Dim int_LastCol_wsSource As Long
        int_LastCol_wsSource = .Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim int_LastRow_wsSource As Long
        int_LastRow_wsSource = WorksheetFunction.Max( _
        .Cells(Rows.Count, "A").End(xlUp).Row, _
        .Cells(Rows.Count, "B").End(xlUp).Row, _
        .Cells(Rows.Count, "C").End(xlUp).Row)
        
    ' Dim "Ranges"
        
    Dim arry_Header_wsSource() As Variant
        arry_Header_wsSource = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol_wsSource)))
        
    Dim col_Target_Source As Integer
        col_Target_Source = fx_Create_Headers_v2(str_Source_TargetField, arry_Header_wsSource)

    Dim col_Match_Source As Integer
        col_Match_Source = fx_Create_Headers_v2(str_Source_MatchField, arry_Header_wsSource)
        
    Dim col_SumField_Source As Integer
        col_SumField_Source = fx_Create_Headers_v2(strSumField, arry_Header_wsSource)
        If strSumField = "" Then col_SumField_Source = 1

    ' Dim Arrays
    
    Dim arry_Target_Source() As Variant
        arry_Target_Source = Application.Transpose(.Range(.Cells(1, col_Target_Source), .Cells(int_LastRow_wsSource, col_Target_Source)))
    
    Dim arry_Match_Source() As Variant
        arry_Match_Source = Application.Transpose(.Range(.Cells(1, col_Match_Source), .Cells(int_LastRow_wsSource, col_Match_Source)))
    
    Dim arry_SumField_Source() As Variant
        arry_SumField_Source = Application.Transpose(.Range(.Cells(1, col_SumField_Source), .Cells(int_LastRow_wsSource, col_SumField_Source)))
        
End With
        
' -----------------------------
' Declare your wsDest variables
' -----------------------------
        
With wsDest
        
    ' Dim wsDest Integers
    
    Dim int_LastCol_wsDest As Long
        int_LastCol_wsDest = wsDest.Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim int_LastRow_wsDest As Long
        int_LastRow_wsDest = WorksheetFunction.Max( _
        wsDest.Cells(Rows.Count, "A").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "B").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "C").End(xlUp).Row)
 
    ' Dim wsDest "Ranges"
    
    Dim arry_Header_wsDest() As Variant
        arry_Header_wsDest = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol_wsDest)))
        
    Dim col_Target_Dest As Integer
        col_Target_Dest = fx_Create_Headers_v2(str_Dest_TargetField, arry_Header_wsDest)

    Dim col_Match_Dest As Integer
        col_Match_Dest = fx_Create_Headers_v2(str_Dest_MatchField, arry_Header_wsDest)
 
    ' Dim Ranges
        
    Dim rng_Target_Dest As Range
    Set rng_Target_Dest = .Range(.Cells(1, col_Target_Dest), .Cells(int_LastRow_wsDest, col_Target_Dest))
        
    ' Dim Arrays
    
    Dim arry_Target_Dest() As Variant
        arry_Target_Dest = Application.Transpose(.Range(.Cells(1, col_Target_Dest), .Cells(int_LastRow_wsDest, col_Target_Dest)))
    
    Dim arry_Match_Dest() As Variant
        arry_Match_Dest = Application.Transpose(.Range(.Cells(1, col_Match_Dest), .Cells(int_LastRow_wsDest, col_Match_Dest)))
 
End With
        
' ----------------------------
' Declare your Other variables
' ----------------------------
        
    ' Dim Dictionaries
    
    Dim dict_LookupData As Scripting.Dictionary
        Set dict_LookupData = New Scripting.Dictionary
        dict_LookupData.CompareMode = TextCompare
        
    ' Declare Loop Variables
    
    Dim i As Long
        
    Dim dblSum As Double
    
    Dim intCount As Long

' ----------------------------------------
' Fill the Dictionary with the Lookup Data
' ----------------------------------------
    
On Error Resume Next
       
        For i = 2 To UBound(arry_Target_Source)
            If arry_Target_Source(i) = str_Criteria Or bolNotBlankSum = True And arry_Target_Source(i) <> "" And arry_Target_Source(i) <> 0 Then
                If arry_Match_Source(i) = arry_Match_Source(i - 1) And arry_Target_Source(i - 1) = str_Criteria Then
                    dblSum = dblSum + arry_SumField_Source(i)
                Else
                    dblSum = dblSum + arry_SumField_Source(i)
                    dict_LookupData.Add Key:=arry_Match_Source(i), Item:=dblSum
                    dblSum = 0
                End If
            
                ' To handle the last record
                If i = UBound(arry_Target_Source) Then
                    dblSum = dblSum + arry_SumField_Source(i)
                    dict_LookupData.Add Key:=arry_Match_Source(i), Item:=dblSum
                    dblSum = 0
                End If
            
            End If
        
        Next i
        
On Error GoTo 0
    
' -----------------------------------------------------------------------------
' Loop through the Lookup Data to pull in the updated data for the target field
' -----------------------------------------------------------------------------

    For i = 2 To UBound(arry_Match_Dest)
        If dict_LookupData.Exists(arry_Match_Dest(i)) Then
            arry_Target_Dest(i) = dict_LookupData.Item(arry_Match_Dest(i))
        End If
    Next i

    ' Output the values from the array
    rng_Target_Dest.Value2 = Application.Transpose(arry_Target_Dest)
    
    ' Empty the Dictionary
    dict_LookupData.RemoveAll

' -------------------------
' Close the Source workbook
' -------------------------
    
    If wsSource.Parent.Name <> ThisWorkbook.Name And bol_CloseSourceWb = True Then
        wsSource.Parent.Close savechanges:=False
    End If

End Function
Function fx_Update_Single_Field(wsSource As Worksheet, wsDest As Worksheet, _
    str_Source_TargetField As String, str_Source_MatchField As String, str_Dest_TargetField As String, str_Dest_MatchField As String, _
    Optional int_SourceHeaderRow As Long, Optional bol_ConvertMatchSourcetoValues As Boolean, _
    Optional bol_CloseSourceWb As Boolean, Optional bol_SkipDuplicates As Boolean, Optional bol_BlanksOnly As Boolean, Optional str_OnlyUseValue As String, _
    Optional bol_MissingLookupData_MsgBox As Boolean, Optional bol_MissingLookupData_UseExistingData As Boolean, _
    Optional strMissingLookupData_ValuetoUse, Optional strWsNameLookup As String, _
    Optional str_FilterField_Dest As String, Optional str_FilterValue As String, Optional bol_FilterPassArray As Boolean)

' Purpose: To update the data in the Target Field in the Destination, based on data from the Target Field in the Source.
    ' 1) Add the data to be referenced into the dictionary
    ' 2) Look for the data to be updated for matches to the reference dictionary
    ' 3) Replace the existing data with the reference data

' Trigger: Called
' Updated: 6/13/2023

' Change Log:
'       2/16/2021:  Initial creation, based on fx_Update_Data_SIC
'       2/17/2021:  Updated to convert over to pulling in the applicable ranges.
'       2/26/2021:  Tweaked the names of the paramaters
'       2/26/2021:  Rewrote to include as much of the code as possible in the function.
'       6/22/2021:  Added the bol_CloseSourceWb variable and related code.
'       6/30/2021:  Added the code to ignore duplicates, just output the value once
'       7/14/2021:  Added the option to convert the Match_Source to values (for Acct #s w/ leading 0s)
'       7/14/2021:  Added the option to pass the int_SourceHeaderRow and the related code
'       10/5/2021:  Updated to use the passed Target & Match fields to determine the int_LastRow
'       10/11/2021: Added the option for bol_BlanksOnly
'       10/12/2021: Added the option to ONLY update with a single value, if present (Ex. updating NPL flag for a borrower)
'       10/13/2021: Updated to convert the range from Text => General formatting, if bol_ConvertMatchSourcetoValues = True
'       4/18/2022:  Added the MsgBox for any missing data, and the bol to use it
'                   Updated the code to determine if there are missing fields to use a dictionary, and created a process to handle blanks
'       4/19/2022:  Added 'strMissingLookupData_ValuetoUse' to allow a user to pass a value that will be used for blanks
'                   Updated the names of some of the variables to help clarify
'       6/13/2022:  Added the 'strWsNameLookup' and related code when a value is missing.
'                   Added code to remove the leading line break in str_missingvalues
'       9/15/2022:  Added the 'Or InStr(1, str_FilterValue, arry_Dest_Filter(i)) > 0' to allow an 'array' to be passed as a criteria
'                   Simplified how the Arrays are determined
'       6/13/2023:  Added a simple example and some clarifications

' ********************************************************************************************************************************************************

' USE EXAMPLE 1 (BASIC): _
    Call fx_Update_Single_Field( _
        wsSource:=wsLists, wsDest:=wsData, _
        str_Source_TargetField:="3. Updated Portfolio", str_Source_MatchField:="3. Sub-Portfolio", _
        str_Dest_TargetField:="PE Updated Portfolio", str_Dest_MatchField:="Sub-Portfolio", _
        bol_MissingLookupData_MsgBox:=True)

' USE EXAMPLE 2: _
    Call fx_Update_Single_Field( _
        wsSource:=wsDetailDash, wsDest:=wsSageworks, _
        int_SourceHeaderRow:=4, _
        str_Source_TargetField:="14 Digit Acct#", _
        str_Source_MatchField:="Full Customer #", _
        str_Dest_TargetField:="Account Number", _
        str_Dest_MatchField:="Full Customer #", _
        bol_ConvertMatchSourcetoValues:=True, _
        strWsNameLookup:="County Lookup", _
        bol_SkipDuplicates:=True, _
        bol_CloseSourceWb:=True, _
        bol_BlanksOnly:= True, _
        str_OnlyUseValue:= "Y", _
        bol_MissingLookupData_MsgBox:=True, _
        bol_MissingLookupData_UseExistingData:=True)

' USE EXAMPLE 3: _
    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="1 - Collateral (#)", _
        str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="1 - Collateral (#)", _
        str_Dest_MatchField:="Helper", _
        str_FilterField_Dest:="Line of Business", _
        str_FilterValue:="Asset Based Lending, Commercial Real Estate, Middle Market Banking, Sponsor and Specialty Finance, Wealth", _
        bol_FilterPassArray:=True)

' LEGEND MANDATORY:
'   wsSource:
'   wsDest:
'   str_Source_TargetField:
'   str_Source_MatchField:
'   str_Dest_TargetField:
'   str_Dest_MatchField:

' LEGEND OPTIONAL:
'   int_SourceHeaderRow: The header row for the Source ws, if blank will default to 1
'   bol_ConvertMatchSourcetoValues: Converts to values only from the Source ws for the Match fields
'   bol_CloseSourceWb: Closes the Source wb when the code has finished
'   bol_SkipDuplicates: Removes duplicate values so the looked up data will only be used once
'   bol_BlanksOnly: Only updates the data if the field is currently blank
'   str_OnlyUseValue: Used to allow ONLY a single value to be used
'   bol_MissingLookupData_MsgBox: Outputs a message box with a list of fields that are missing from the lookup
'   bol_MissingLookupData_UseExistingData: Will use the existing data instead of the lookup value
'   strMissingLookupData_ValuetoUse: If the value isn't in the lookup, and I didn't include a blank in the lookups, will use this value instead
'   strWsNameLookup: If the value isn't in the lookup it will say where it was looking to help with troubleshooting
'   str_FilterField_Dest: Used to filter down the values to be imported on
'   str_FilterValue: Used to filter down the values to be imported on
'   bol_FilterPassArray: Allows an array of values to be passed instead of a single value for the filter

' ********************************************************************************************************************************************************

' -------------------------------
' Declare your wsSource variables
' -------------------------------
        
With wsSource
        
    ' Dim wsSource Range References
    
    Dim intHeaderRow_wsSource As Long
    
        If int_SourceHeaderRow <> 0 Then
            intHeaderRow_wsSource = int_SourceHeaderRow
        Else
            intHeaderRow_wsSource = 1
        End If
    
    Dim int_LastCol_wsSource As Long
        int_LastCol_wsSource = .Cells(intHeaderRow_wsSource, Columns.Count).End(xlToLeft).Column
        
    ' Dim wsSource Column References
        
    Dim arry_Header_wsSource() As Variant
        arry_Header_wsSource = Application.Transpose(.Range(.Cells(intHeaderRow_wsSource, 1), .Cells(intHeaderRow_wsSource, int_LastCol_wsSource)))
        
    Dim col_Source_Target As Integer
        col_Source_Target = fx_Create_Headers_v2(str_Source_TargetField, arry_Header_wsSource)

    Dim col_Source_Match As Integer
        col_Source_Match = fx_Create_Headers_v2(str_Source_MatchField, arry_Header_wsSource)

    ' Dim wsSource Range References
        
    Dim int_LastRow_wsSource As Long
        int_LastRow_wsSource = WorksheetFunction.Max( _
        .Cells(Rows.Count, col_Source_Target).End(xlUp).Row, _
        .Cells(Rows.Count, col_Source_Match).End(xlUp).Row)
        
        If int_LastRow_wsSource = 1 Then int_LastRow_wsSource = 2
        
    ' Dim wsSource Ranges
        
    Dim rng_Source_Target As Range
    Set rng_Source_Target = .Range(.Cells(1, col_Source_Target), .Cells(int_LastRow_wsSource, col_Source_Target))
        
    Dim rng_Source_Match As Range
    Set rng_Source_Match = .Range(.Cells(1, col_Source_Match), .Cells(int_LastRow_wsSource, col_Source_Match))
        
    If bol_ConvertMatchSourcetoValues = True Then
        rng_Source_Match.NumberFormat = "General"
        rng_Source_Match.Value = rng_Source_Match.Value
        Set rng_Source_Match = .Range(.Cells(1, col_Source_Match), .Cells(int_LastRow_wsSource, col_Source_Match))
    End If
        
    ' Dim wsSource Arrays
    
    Dim arry_Source_Target() As Variant
        arry_Source_Target = Application.Transpose(rng_Source_Target)
    
    Dim arry_Source_Match() As Variant
        arry_Source_Match = Application.Transpose(rng_Source_Match)
        
End With
        
' -----------------------------
' Declare your wsDest variables
' -----------------------------
        
With wsDest
        
    ' Dim wsDest Range References
    
    Dim int_LastCol_wsDest As Long
        int_LastCol_wsDest = wsDest.Cells(1, Columns.Count).End(xlToLeft).Column
 
    ' Dim wsDest Column References
    
    Dim arry_Header_wsDest() As Variant
        arry_Header_wsDest = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol_wsDest)))
        
    Dim col_Dest_Target As Integer
        col_Dest_Target = fx_Create_Headers_v2(str_Dest_TargetField, arry_Header_wsDest)

    Dim col_Dest_Match As Integer
        col_Dest_Match = fx_Create_Headers_v2(str_Dest_MatchField, arry_Header_wsDest)
        
    Dim col_Dest_FilterField As Integer
        col_Dest_FilterField = fx_Create_Headers_v2(str_FilterField_Dest, arry_Header_wsDest)
        If col_Dest_FilterField = 0 Then col_Dest_FilterField = 999
 
    ' Dim wsDest Range References
 
    Dim int_LastRow_wsDest As Long
        int_LastRow_wsDest = WorksheetFunction.Max( _
        wsDest.Cells(Rows.Count, col_Dest_Target).End(xlUp).Row, _
        wsDest.Cells(Rows.Count, col_Dest_Match).End(xlUp).Row)
        
        If int_LastRow_wsDest = 1 Then int_LastRow_wsDest = 2
 
    ' Dim Ranges
        
    Dim rng_Dest_Target As Range
    Set rng_Dest_Target = .Range(.Cells(1, col_Dest_Target), .Cells(int_LastRow_wsDest, col_Dest_Target))
        
    ' Dim Arrays
    
    Dim arry_Dest_Target() As Variant
        arry_Dest_Target = Application.Transpose(.Range(.Cells(1, col_Dest_Target), .Cells(int_LastRow_wsDest, col_Dest_Target)))
    
    Dim arry_Dest_Match() As Variant
        arry_Dest_Match = Application.Transpose(.Range(.Cells(1, col_Dest_Match), .Cells(int_LastRow_wsDest, col_Dest_Match)))

    Dim arry_Dest_Filter() As Variant
        arry_Dest_Filter = Application.Transpose(.Range(.Cells(1, col_Dest_FilterField), .Cells(int_LastRow_wsDest, col_Dest_FilterField)))
 
End With
        
' ----------------------------
' Declare your Other variables
' ----------------------------
        
    ' Dim Dictionaries
    
    Dim dict_LookupData As Scripting.Dictionary
        Set dict_LookupData = New Scripting.Dictionary
            dict_LookupData.CompareMode = TextCompare
        
    Dim dict_MissingFields As Scripting.Dictionary
        Set dict_MissingFields = New Scripting.Dictionary
            dict_MissingFields.CompareMode = TextCompare
        
    ' Declare Loop Variables
    
    Dim i As Long
    
    Dim cntr_MissingFields As Integer
    
    Dim str_MissingValues As String
    
    Dim val As Variant
    
    ' Declare Message Variables
    
    Dim strMissingDataMessage As String
    
    If strWsNameLookup <> "" Then
        strMissingDataMessage = strWsNameLookup
    Else
        strMissingDataMessage = "(ex. 'Collateral Lookup')"
    End If

' ----------------------------------------
' Fill the Dictionary with the Lookup Data
' ----------------------------------------
    
On Error Resume Next
        
    For i = 1 To UBound(arry_Source_Target)
        If arry_Source_Target(i) <> "" And arry_Source_Match(i) <> "" Then
            If str_OnlyUseValue <> "" Then
                If arry_Source_Target(i) = str_OnlyUseValue Then ' Only import if the target matches the passed value
                    dict_LookupData.Add Key:=arry_Source_Match(i), Item:=arry_Source_Target(i)
                End If
            Else
                dict_LookupData.Add Key:=arry_Source_Match(i), Item:=arry_Source_Target(i)
            End If
        End If
    Next i

On Error GoTo 0
    
' -----------------------------------------------------------------------------
' Loop through the Lookup Data to pull in the updated data for the target field
' -----------------------------------------------------------------------------

    For i = 2 To UBound(arry_Dest_Match)
        
        If str_FilterField_Dest = "" Or arry_Dest_Filter(i) = str_FilterValue Or bol_FilterPassArray = True And InStr(1, str_FilterValue, arry_Dest_Filter(i)) > 0 Then
        
            If dict_LookupData.Exists(arry_Dest_Match(i)) Then
                
                If bol_BlanksOnly = True Then
                    If arry_Dest_Target(i) = "" Or arry_Dest_Target(i) = 0 Then
                        arry_Dest_Target(i) = dict_LookupData.Item(arry_Dest_Match(i))
                    End If
                Else
                    arry_Dest_Target(i) = dict_LookupData.Item(arry_Dest_Match(i))
                End If
                
                If bol_SkipDuplicates = True Then dict_LookupData.Remove (arry_Dest_Match(i)) ' Remove so it can only be imported once
            ElseIf arry_Dest_Match(i) = Empty Then
            
                ' If I have a record for a blank in the lookups use that, or use the strMissingLookupData_ValuetoUse if that was passed, otherwise abort
                If dict_LookupData.Exists(" ") = True Then
                    arry_Dest_Target(i) = dict_LookupData.Item(" ")
                
                ElseIf IsMissing(strMissingLookupData_ValuetoUse) = False Then
                    arry_Dest_Target(i) = strMissingLookupData_ValuetoUse
                    GoTo MissingDataMsgBox
                Else
                    GoTo MissingDataMsgBox
                End If
            
            Else

MissingDataMsgBox:

                If bol_MissingLookupData_MsgBox = True Then ' Let the user know that the data is missing
                    
                    ' Load the dictionary with each of the exceptions noted
                    On Error Resume Next
                        dict_MissingFields.Add Key:=arry_Dest_Match(i), Item:=arry_Dest_Match(i)
                        cntr_MissingFields = cntr_MissingFields + 1
                    On Error GoTo 0
                    
                End If
                
                If bol_MissingLookupData_UseExistingData = True Then ' Use the existing data to fill in the blank
                    arry_Dest_Target(i) = arry_Dest_Match(i)
                End If
                
                'wsDest.Cells(i, col_Dest_Target).Interior.Color = RGB(252, 213, 180) ' Highlight the missing data (disabled 6/13/2022)
            End If
        
        End If
        
    Next i

' -------------------------------------------------------
' Create the MsgBox if bol_MissingLookupData_MsgBox = True
' -------------------------------------------------------

        ' Create the list of fields
        For Each val In dict_MissingFields
            str_MissingValues = str_MissingValues & Chr(10) & "  > " & CStr(dict_MissingFields(val))
        Next val
        
        If Left(str_MissingValues, 1) = vbLf Then
            str_MissingValues = Right(str_MissingValues, Len(str_MissingValues) - 2)
        End If
        
        ' Output the Messagebox if there were any results
        If cntr_MissingFields > 0 Then
                
            MsgBox Title:="Missing Values in Lookup", _
                Buttons:=vbOKOnly + vbExclamation, _
                Prompt:="There is a missing record in the lookup table for:" & Chr(10) _
                & "'" & str_MissingValues & "'" & Chr(10) & Chr(10) _
                & "Please review the lookups in the applicable worksheet: " & Chr(10) _
                & strMissingDataMessage & Chr(10) & Chr(10) _
                & "Once the data has been reviewed re-run this process, or manually update the data."
        
        End If

    'Output the values from the array
    rng_Dest_Target.Value2 = Application.Transpose(arry_Dest_Target)

' -------------------------
' Close the Source workbook
' -------------------------
    
    If wsSource.Parent.Name <> ThisWorkbook.Name And bol_CloseSourceWb = True Then
        wsSource.Parent.Close savechanges:=False
    End If

End Function
Function fx_Open_Workbook(strPromptTitle As String) As Workbook
             
' Purpose: This function will prompt the user for the workbook to open and returns that workbook.
' Trigger: Called Function
' Updated: 5/5/2023

' Use Example: Set wbTEST = fx_Open_Workbook(strPromptTitle:="Select the current Sageworks data dump")

' Change Log:
'       2/12/2021:  Initial Creation
'       2/12/2021:  Added the code to abort if the user selects cancel.
'       2/12/2021:  Added the code to determine if the Workbook is already open.
'       6/16/2021:  Added the code to ChDrive and ChDir
'       5/5/2023:   Repointed to 'Source Report' as the starting point, not '(Source Data)'

' ****************************************************************************
             
On Error Resume Next
    ChDrive ThisWorkbook.path
    ChDir ThisWorkbook.path & "\Source Reports"
On Error GoTo 0
             
' -----------
' Declare your variables
' -----------
             
    Dim str_wbPath As String
        str_wbPath = Application.GetOpenFilename( _
        Title:=strPromptTitle, FileFilter:="Excel Workbooks (*.xls*;*.csv),*.xls*;*.csv")
             
        If str_wbPath = "False" Then
            MsgBox "No Workbook was selected, the code cannont continue."
            myPrivateMacros.DisableForEfficiencyOff
            End
        End If
        
' -----------
' Determine if the Workbook is already open
' -----------
        
    Dim bolAlreadyOpen As Boolean
        
     Dim str_wbName As String
         str_wbName = Right(str_wbPath, Len(str_wbPath) - InStrRev(str_wbPath, "\"))
        
    On Error Resume Next
        Dim wb As Workbook
        Set wb = Workbooks(str_wbName)
        bolAlreadyOpen = Not wb Is Nothing
    On Error GoTo 0
        
' -----------
' Obtain the Workbook
' -----------
        
    If bolAlreadyOpen = True Then
        Set fx_Open_Workbook = Workbooks(str_wbName)
    Else
        Set fx_Open_Workbook = Workbooks.Open(str_wbPath, UpdateLinks:=False, ReadOnly:=True)
    End If

End Function
Function fx_Steal_First_Row_Formating(ws As Worksheet, Optional intFirstRow As Long, Optional int_LastCol As Long, Optional int_LastRow As Long, Optional intSingleRow As Long)

' Purpose: To copy the formatting from the first row of data and apply to the rest of the data.
' Trigger: Called
' Updated: 1/8/2022

' Use Example: _
    Call fx_Steal_First_Row_Formating( _
        ws:=wsQCReview, _
        intFirstRow:=2, _
        int_LastRow:=int_LastRow, _
        int_LastCol:=int_LastCol)

' Use Example 2: Call fx_Steal_First_Row_Formating(ws:=wsQCReview)

' Change Log:
'       5/17/2021:  Intial Creation
'       6/16/2021:  Added the 'Application.Goto' to reset the copy paste
'       12/6/2021:  Added the option to pass only a single row
'       12/8/2021:  Added the rngCur so the screen doesn't jump around
'       1/8/2022:   Updated some of the passed variables to be optional
'                   Defaulted intFirstRow to be 2 if not passed

' ****************************************************************************
  
' -----------------
' Declare Variables
' -----------------

    Dim rngCur As Range
    Set rngCur = ActiveCell

    ' Declare Integers
    
    If intFirstRow = 0 Then intFirstRow = 2
    
    If int_LastRow = 0 Then
       int_LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    End If
    
    If int_LastCol = 0 Then
       int_LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    End If
    
    ' Declare Ranges
    
With ws

    Dim rngFormat As Range
        Set rngFormat = .Range(.Cells(intFirstRow, 1), .Cells(intFirstRow, int_LastCol))
    
    Dim rngTarget As Range
        If intSingleRow <> 0 Then
            Set rngTarget = .Range(.Cells(intSingleRow, 1), .Cells(intSingleRow, int_LastCol))
        ElseIf int_LastRow <> 0 Then
            Set rngTarget = .Range(.Cells(intFirstRow + 1, 1), .Cells(int_LastRow, int_LastCol))
        Else
            MsgBox "There was no row passed to the Steal First Row function."
        End If

End With

' ---------------------------------------------------------------------------------------------------
' Copy the formatting from the first row of data (intFirstRow) to the remaining rows (thru int_LastRow)
' ---------------------------------------------------------------------------------------------------
    
    rngFormat.Copy: rngTarget.PasteSpecial xlPasteFormats
    
    ' Go back to where you were before the code
    Application.CutCopyMode = False
    Application.GoTo Reference:=rngCur, Scroll:=False
    
End Function
Function fx_Select_Worksheet(strWbName As String, strWsName As String, Optional strWsNamedRange As String) As Worksheet

' Purpose: To determine if a sheet exists, otherwise prompt the user for the sheet to select.
' Trigger: Called
' Updated: 6/5/2023

' Use Example:
'    Dim wsSageworksRiskTrend As Worksheet
'    Set wsSageworksRiskTrend = fx_Select_Worksheet(wbSageworksChainedReport.Name, "1022 - Risk Trend 2873 ME Extra111")

'    Dim ws5003_Source As Worksheet
'    Set ws5003_Source = fx_Select_Worksheet( _
        strWbName:=wbSageworksChainedReport.Name, _
        strWsName:=ThisWorkbook.Worksheets(1).Evaluate("wsName_5003"), _
        strWsNamedRange:="wsName_5003")

' LEGEND OPTIONAL:
'   strWsNamedRange: The name of the Named Range that corresponds to the applicable name of the worksheet.  If this is passed AND no sheet is found, update with the value selected by the user.

' Change Log:
'       7/4/2021:   Initial Creation, combined fx_Sheets_Exists and fx_Pick_Worksheet
'       6/5/2023:   Added the strWsNamedRange optional varible to allow the user to update the ws Name

' ****************************************************************************

' ---------------------------------
' Determine if the worksheet exists
' ---------------------------------

On Error GoTo SheetNotFound

    If Evaluate("ISREF('[" & strWbName & "]" & strWsName & "'!A1)") = True Then
        Set fx_Select_Worksheet = Workbooks(strWbName).Sheets(strWsName)
        Exit Function
    End If

SheetNotFound:

' -----------------------------------
' Prompt the user to select the sheet
' -----------------------------------

    MsgBox "The " & strWsName & " was not found, please select the correct Worksheet."
        uf_wsSelector.Show
    
' ---------------------------
' Pass the selected worksheet
' ---------------------------
            
    If str_wsSelected = "" Then
        MsgBox "No Worksheet was selected, the code cannont continue."
        myPrivateMacros.DisableForEfficiencyOff
        End
    Else
        Set fx_Select_Worksheet = Workbooks(strWbName).Worksheets(str_wsSelected)
        If strWsNamedRange <> "" Then ThisWorkbook.Names(strWsNamedRange).Value = str_wsSelected 'Updated the Named Range if it exists
    End If

End Function
Function fx_Delete_Unused_Data(ws As Worksheet, str_Target_Field As String, str_Value_To_Delete As String, Optional bol_DeleteValues_PassArray As Boolean)

' Purpose: To delete data from the passed worksheet where the "Value To Delete" is in the Target Field.
' Trigger: Called
' Updated: 6/12/2023

' Use Example: _
    Call fx_Delete_Unused_Data( _
        ws:=wsSageworksRT_Dest, _
        str_Target_Field:="Line of Business", _
        str_Value_To_Delete:="Small Business")

' LEGEND OPTIONAL:
'   bol_FilterPassArray: Allows an array of values to be passed instead of a single value for the filter

' Change Log:
'       9/15/2021:  Initial Creation
'       5/5/2023:   Added the 'Exit Function' if the int_LastRow is 1, to handle situations where there is no data (ex. Policy Exceptions - Paid Off Loans)
'                   Updated the delete to keep the first row of data for the formatting
'       6/12/2023:  Added the option to pass an array of values to be deleted, and the code around arryValuesToDelete

' ****************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    ' Declare Cell References
       
    Dim int_LastRow As Long
        int_LastRow = WorksheetFunction.Max( _
            ws.Cells(Rows.Count, "A").End(xlUp).Row, _
            ws.Cells(Rows.Count, "B").End(xlUp).Row, _
            ws.Cells(Rows.Count, "C").End(xlUp).Row)
       
       If int_LastRow = 1 Then Exit Function
       
    Dim int_LastCol As Integer
        int_LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column

    ' Declare "Ranges"
    
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(ws.Range(ws.Cells(1, 1), ws.Cells(1, int_LastCol)))

    Dim col_Target As Integer
        col_Target = fx_Create_Headers_v2(str_Target_Field, arry_Header)
        
    ' Declare Arrays
    
    Dim arryValuesToDelete() As String
    
    If bol_DeleteValues_PassArray = True Then
        arryValuesToDelete = Split(str_Value_To_Delete, ", ")
    End If
    
' ----------------------
' Delete the Unused Data
' ----------------------
        
On Error Resume Next
        
With ws
    
    ' Sort the data to make deleting MUCH faster
    .Range(.Cells(1, 1), .Cells(int_LastRow, int_LastCol)).Sort _
        Key1:=.Cells(1, col_Target), Order1:=xlAscending, Header:=xlYes
    
    ' Filter then delete the filtered data
    If bol_DeleteValues_PassArray = False Then
    
        .Range("A1").AutoFilter Field:=col_Target, Criteria1:=str_Value_To_Delete, Operator:=xlFilterValues
            .Range("A2:A3").SpecialCells(xlCellTypeVisible).EntireRow.ClearContents
            .Range("A3:A" & int_LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .Range("A1").AutoFilter Field:=col_Target
        
    Else
    
        .Range("A1").AutoFilter Field:=col_Target, Criteria1:=arryValuesToDelete, Operator:=xlFilterValues
            .Range("A2:A3").SpecialCells(xlCellTypeVisible).EntireRow.ClearContents
            .Range("A3:A" & int_LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .Range("A1").AutoFilter Field:=col_Target
    
    End If
    
End With

On Error GoTo 0

End Function
Function fx_Rename_Header_Field(ws As Worksheet, intHeaderRow As Long, str_Value_To_Update As String, str_Updated_Value As String)

' Purpose: To rename one of the fields in the header of the passed workbook.
' Trigger: Called
' Updated: 9/16/2021

' Use Example: _
    Call fx_Rename_Header_Field( _
        ws:=wsSageworksRT_Dest, _
        intHeaderRow:=8, _
        str_Value_To_Update:="Line of Business (Branch)", _
        str_Updated_Value:="Line of Business")

' Change Log:
'       9/16/2021: Initial Creation

' ****************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    ' Declare Cell References
              
    Dim int_LastCol As Integer
        int_LastCol = ws.Cells(intHeaderRow, Columns.Count).End(xlToLeft).Column

    ' Declare "Ranges"
    
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(ws.Range(ws.Cells(intHeaderRow, 1), ws.Cells(intHeaderRow, int_LastCol)))

    Dim col_Target As Integer
        col_Target = fx_Create_Headers_v2(str_Value_To_Update, arry_Header)
    
' -----------------------
' Update the Header Field
' -----------------------
        
On Error Resume Next
    
    ws.Cells(intHeaderRow, col_Target).Value2 = str_Updated_Value
    '.Cells(2, .Rows(2).Find("BookBalance").Column).Value = "Outstanding"
    
On Error GoTo 0

'       9/16/2021: To be added
'        > If it's missing then go to errhandler and output a message box
'            > Use the message box to determine if I continue or abort

End Function
Function fx_Remove_Webster_Internal_Loans(wsTarget As Worksheet, strAccountNumber As String)

' Purpose: To remove the Webster Internal Loans from the passed data set.
' Trigger: Called
' Updated: 10/12/2021

' Use Example:
'    Call fx_Remove_Webster_Internal_Loans(wsTarget:=wsData, strAccountNumber:="Account Number")

' Change Log:
'       10/12/2021: Intial Creation

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    ' Declare Range References
    
    Dim int_LastRow As Long
        int_LastRow = WorksheetFunction.Max( _
            wsTarget.Cells(Rows.Count, "A").End(xlUp).Row, _
            wsTarget.Cells(Rows.Count, "B").End(xlUp).Row, _
            wsTarget.Cells(Rows.Count, "C").End(xlUp).Row)

    Dim int_LastCol As Long
        int_LastCol = WorksheetFunction.Max( _
        wsTarget.Cells(1, Columns.Count).End(xlToLeft).Column, _
        wsTarget.Rows(1).Find("").Column - 1)
        
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose( _
            wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(1, int_LastCol)))
    
    Dim colAcctNum As Long
        colAcctNum = fx_Create_Headers_v2("Account Number", arry_Header)

    If colAcctNum = 0 Then
        MsgBox "Couldn't find the Account Number field"
        Stop
    End If

    ' Declare Loop Variables
    
    Dim i As Long
    
    ' Declare Dictionary Variables
    
    Dim dict_WebsterLoans As Scripting.Dictionary
    Set dict_WebsterLoans = New Scripting.Dictionary

' ----------------------------------------
' Load the dictionary w/ the Webster loans
' ----------------------------------------

With dict_WebsterLoans
    .Add Key:="00026000841488", Item:="WEBSTER CAPITAL FINANCE"
    .Add Key:="00260000583712", Item:="WEBSTER CAPITAL FINANCE"
    .Add Key:="00026000733553", Item:="WEBSTER BANK-SPONSOR & SPECIALTY"
    .Add Key:="00260000561917", Item:="WEBSTER BANK-SPONSOR & SPECIALTY"
    .Add Key:="00260000521422", Item:="WEBSTER FINANCIAL CORP"
    .Add Key:="26000841488", Item:="WEBSTER CAPITAL FINANCE"
    .Add Key:="260000583712", Item:="WEBSTER CAPITAL FINANCE"
    .Add Key:="26000733553", Item:="WEBSTER BANK-SPONSOR & SPECIALTY"
    .Add Key:="260000561917", Item:="WEBSTER BANK-SPONSOR & SPECIALTY"
    .Add Key:="260000521422", Item:="WEBSTER FINANCIAL CORP"
End With

' --------------------------------------------
' Delete the loan if it is present in the data
' --------------------------------------------

    For i = int_LastRow To 2 Step -1
    
        If dict_WebsterLoans.Exists(CStr(wsTarget.Cells(i, colAcctNum).Value2)) Then
            wsTarget.Rows(i).Delete
        End If

    Next i

End Function
Function fx_Convert_LOB_Name(rngToUpdate As Range, Optional rngRegion As Range, Optional rngSubPortfolio As Range, _
Optional bolPrivateBank_to_Wealth As Boolean, Optional bolSmallBusiness_to_BusinessBanking As Boolean, Optional bolPublicSectorFinance As Boolean, _
Optional bol_EF_To_Seperate_LOB_not_CRE As Boolean, Optional bol_EF_Using_SubPort As Boolean)

' Purpose: To rename (convert) the given LOB name in the passed range to what it should be.
' Trigger: Called
' Updated: 6/12/2023

' Use Example: _
    Call fx_Convert_LOB_Name( _
        rngToUpdate:=wsData.Range(wsData.Cells(2, col_LOB), wsData.Cells(int_LastRow, col_LOB)), _
        bolSmallBusiness_to_BusinessBanking:=True)

' Use Example 2: _
    Call fx_Convert_LOB_Name( _
        rngToUpdate:=wsData.Range(wsData.Cells(2, col_LOB), wsData.Cells(int_LastRow, col_LOB)), _
        rngRegion:=wsData.Range(wsData.Cells(2, col_Region), wsData.Cells(int_LastRow, col_Region)), _
        bolPublicSectorFinance:=True)

' LEGEND OPTIONAL:
'   bol_EF_To_Seperate_LOB_not_CRE: Use the Region (RTB Low) to convert the LOB to be 'Equipment Finance' not 'Commercial Real Estate'
'   rngSubPortfolio: The range that includes the data for Sub Portfolio that corresponds with the rngToUpdate
'   bol_EF_Using_SubPort: Use the Sub Portfolio field to convert the LOB to be 'Equipment Finance'

' Change Log:
'       10/5/2021:  Initial Creation
'       10/30/2021: Combined the two related functions ('fx_Rename_PrivateBank_to_Wealth' and 'fx_Rename_SmallBusiness_to_BusinessBanking')
'       10/30/2021: Removed the Application.Transpose, isn't necessary
'       3/17/2022:  Updated to include the conversion for Public Sector Finance
'       4/4/2022:   Added the conversion for EF to go back to EF vs being part of CRE
'       5/18/2022:  Switched from PSF being broken out to moving it back under Middle Market Banking
'       6/12/2023:  Added the option to do Equipment Finance based on SubPort
'                   Added Application.transpose to make 1 dimensional arrays and updated 'arry' to 'arryUpdateValues' to be more explicit

' Note: As of 4/4/2022 MM Healthcare is aligned with CRE, so that update isn't necessary

' ****************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    ' Declare Arrays
    
    Dim arryUpdateValues() As Variant
        arryUpdateValues() = Application.Transpose(rngToUpdate)

    Dim arryRegion() As Variant
        If rngRegion Is Nothing Then
            ' Do Nothing
        Else
            arryRegion() = Application.Transpose(rngRegion)
        End If

    Dim arrySubPortfolio() As Variant
        If rngSubPortfolio Is Nothing Then
            ' Do Nothing
        Else
            arrySubPortfolio() = Application.Transpose(rngSubPortfolio)
        End If

    Dim i As Long

' ----------------------------
' Update the Data in the Array
' ----------------------------
        
    For i = LBound(arryUpdateValues) To UBound(arryUpdateValues)
        
        ' Private Bank -> Wealth
        If bolPrivateBank_to_Wealth = True Then
            If arryUpdateValues(i) = "Webster Private Bank" Or arryUpdateValues(i) = "Private Bank" Then
                arryUpdateValues(i) = "Wealth"
            End If
        End If
                
        ' Small Business -> Business Banking
        If bolSmallBusiness_to_BusinessBanking = True Then
            If arryUpdateValues(i) = "Small Business" Then
                arryUpdateValues(i) = "Business Banking"
            End If
        End If
        
        ' Public Sector Finance
        If bolPublicSectorFinance = True Then
            If Right(arryRegion(i), 21) = "Public Sector Finance" Then
                'arryUpdateValues(i) = "Public Sector Finance"
                arryUpdateValues(i) = "Middle Market Banking"
            End If
        End If
        
        ' Equipment Finance by Region
        If bol_EF_To_Seperate_LOB_not_CRE = True Then
            If arryRegion(i) = "00770 Equipment Finance" Or arryRegion(i) = "00771 Equip Fin Remediation" Then
                arryUpdateValues(i) = "Equipment Finance"
            End If
        End If
        
        ' Equipment Finance by Sub Portfolio
        If bol_EF_Using_SubPort = True Then
            If arrySubPortfolio(i) = "Equipment Finance" Then
                arryUpdateValues(i) = "Equipment Finance"
            End If
        End If
        
    Next i
    
    rngToUpdate = Application.Transpose(arryUpdateValues)
        
End Function
Function fx_Clear_Old_Data(ws As Worksheet, Optional bolDeleteHeader As Boolean, Optional bolClearFormatting As Boolean, Optional bolKeepFirstRow As Boolean)

' Purpose: To clear the existing data as the first step of an import process.
' Trigger: Called
' Updated: 1/31/2022

' Use Example: _
    Call fx_Clear_Old_Data( _
        ws:=wsSageworksRT_Dest, _
        bolDeleteHeader:=False, _
        bolKeepFirstRow:=True, _
        bolClearFormatting:=True)

' Change Log:
'       10/5/2021:  Intial Creation
'       10/20/2021: Updated to include the firstrow and allow formatting to be cleared.
'       1/8/2022:   Updated so that the first row of formatting gets retained
'       1/31/2022:  Allow the first row of data to be retained, to keep formulas intact
'                   Made 'bolClearFormatting' and 'bolDeleteHeader' Optional

' ********************************************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Declare Integers
    
    Dim int_LastRow As Long
        int_LastRow = WorksheetFunction.Max( _
            ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
            ws.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row, _
            ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row)
            
        If int_LastRow = 1 Then int_LastRow = 2
            
    Dim intFirstRow As Long
        If bolDeleteHeader = True Then
            intFirstRow = 1
        ElseIf bolKeepFirstRow = True Then
            intFirstRow = 3
        Else
            intFirstRow = 2
        End If
    
    Dim int_LastCol As Long
        int_LastCol = WorksheetFunction.Max( _
        ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column, _
        ws.Cells(int_LastRow, Columns.Count).End(xlToLeft).Column, _
        ws.Rows(1).Find("").Column - 1)

' -------------
' Wipe Old Data
' -------------
   
    ws.Range(ws.Cells(intFirstRow, 1), ws.Cells(int_LastRow, int_LastCol)).ClearContents
    
    If bolClearFormatting = True Then
        ws.Range(ws.Cells(intFirstRow + 1, 1), ws.Cells(int_LastRow + 1, int_LastCol)).ClearFormats
    End If

End Function
Function fx_Update_Named_Range(strNamedRangeName As String)

' Purpose: To update the passed Named Range a change in the Change Log.
' Trigger: Called
' Updated: 3/6/2022

' Use Example: _
    Call fx_Update_Named_Range("ChangeLog_Data")

' Change Log:
'       12/8/2021:  Intial Creation
'       3/4/2022:   Added Error Handling for the int_LastRow and int_LastCol to handle if all of the empty rows/cols are hidden
'       3/6/2022:   Replaced the int_LastRow and int_LastCol w/ functions

' ****************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strWsName As String
        strWsName = ThisWorkbook.Names(strNamedRangeName).RefersToRange.Parent.Name
    
    Dim wsNamedRange As Worksheet
    Set wsNamedRange = ThisWorkbook.Sheets(strWsName)
    
    Dim int_LastRow As Long
        int_LastRow = fx_Find_LastRow(wsNamedRange)
        
    Dim int_LastCol As Integer
        int_LastCol = fx_Find_LastColumn(wsNamedRange)

' ----------------------
' Update the Named Range
' ----------------------

   ThisWorkbook.Names(strNamedRangeName).RefersToR1C1 = wsNamedRange.Range(wsNamedRange.Cells(1, 1), wsNamedRange.Cells(int_LastRow, int_LastCol))

End Function
Function fx_Find_LastRow(wsTarget As Worksheet, Optional intTargetColumn As Long, Optional bolIncludeSpecialCells As Boolean, Optional bolIncludeUsedRange As Boolean) As Long

' Purpose: To output the the Last Row for the passed ws using multiple options.
' Trigger: Called
' Updated: 3/6/2022

' Use Example: _
'   int_LastRow = fx_Find_LastRow(wsData)

' Use Example 2: Using all of the optional variables _
'   int_LastRow = fx_Find_LastRow(wsTarget:=wsTest, intTargetColumn:=2, bolIncludeSpecialCells:=True, bolIncludeUsedRange:=True)

'bolIncludeUsedRange: If this is True then the last row of the UsedRange will be included in the Max formula
'bolIncludeSpecialCells: If this is True then the SpecialCells(xlCellTypeLastCell) row will be included in the Max formula

' Change Log:
'       11/29/2021: Initial Creation
'       3/6/2022:   Overhauled to include error handling, and the if statements to breakout the determination of the Last Row
'                   Added the fx_Find_Row code as an alternative to handle filtered data

' ****************************************************************************
  
' -----------------
' Declare Variables
' -----------------

On Error Resume Next

    Dim int_LastRow_1st As Long
    If intTargetColumn <> 0 Then
        int_LastRow_1st = wsTarget.Cells(wsTarget.Rows.Count, intTargetColumn).End(xlUp).Row
    Else
        int_LastRow_1st = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    End If
        
    Dim int_LastRow_2nd As Long
    If bolIncludeSpecialCells = True Then
        int_LastRow_2nd = wsTarget.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row
    End If
    
    Dim int_LastRow_3rd As Long
    If bolIncludeUsedRange = True Then
        int_LastRow_3rd = wsTarget.UsedRange.Rows(wsTarget.UsedRange.Rows.Count).Row
    End If
    
    Dim int_LastRow_4th As Long
        int_LastRow_4th = fx_Find_Row(ws:=wsTarget, strTarget:="") - 1

    Dim int_LastRow_Max As Long

On Error Resume Next

' ---------------------------------
' Determine which int_LastRow to use
' ---------------------------------

    int_LastRow_Max = WorksheetFunction.Max(int_LastRow_1st, int_LastRow_2nd, int_LastRow_3rd, int_LastRow_4th)
        If int_LastRow_Max = 0 Or int_LastRow_Max = 1 Then int_LastRow_Max = 2 ' Don't pass 0 or 1
        fx_Find_LastRow = int_LastRow_Max
        
End Function
Function fx_Find_LastColumn(wsTarget As Worksheet, Optional bolIncludeSpecialCells As Boolean, Optional bolIncludeUsedRange As Boolean) As Long

' Purpose: To output the the Last Column for the passed ws using multiple options.
' Trigger: Called
' Updated: 3/6/2022

' Use Example: _
'   int_LastCol = fx_Find_LastColumn(wsData)

'bolIncludeUsedRange: If this is True then the last Col of the UsedRange will be included in the Max formula
'bolIncludeSpecialCells: If this is True then the SpecialCells(xlCellTypeLastCell) col will be included in the Max formula

' Change Log:
'       3/6/2022:  Initial Creation, based on fx_Find_LastCol

' ****************************************************************************
  
' -----------------
' Declare Variables
' -----------------

On Error Resume Next

    Dim int_LastCol_1st As Long
        int_LastCol_1st = wsTarget.Cells(1, wsTarget.Columns.Count).End(xlToLeft).Column
        
    Dim int_LastCol_2nd As Long
    If bolIncludeSpecialCells = True Then
        int_LastCol_2nd = wsTarget.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Column
    End If
    
    Dim int_LastCol_3rd As Long
    If bolIncludeUsedRange = True Then
        int_LastCol_3rd = wsTarget.UsedRange.Columns(wsTarget.UsedRange.Columns.Count).Column
    End If

    Dim int_LastCol_Max As Long

On Error Resume Next

' ---------------------------------
' Determine which int_LastCol to use
' ---------------------------------

    int_LastCol_Max = WorksheetFunction.Max(int_LastCol_1st, int_LastCol_2nd, int_LastCol_3rd)
        If int_LastCol_Max = 0 Or int_LastCol_Max = 1 Then int_LastCol_Max = 2 ' Don't pass 0 or 1
        fx_Find_LastColumn = int_LastCol_Max
        
End Function
Function fx_Find_Row(ws As Worksheet, strTarget As String, Optional strTargetFieldName As String, Optional strTargetCol As String) As Long

' Purpose: To find the target value in the passed column for the passed worksheet.  Replaces the Find function, to account for hidden rows.
' Trigger: Called
' Updated: 3/6/2022

' Use Example: _
    Call fx_Find_Row( _
        ws:=ThisWorkbook.Sheets("Projects"), _
        strTargetFieldName:="Project", _
        strTarget:="P.343 - Migrate to Win10")

' Use Example 2: Passing the Target Field Name _
    intRowCurProject = fx_Find_Row(ws:=wsProjects, strTarget:=strProjName, strTargetFieldName:="Project")

' Use Example 3: Passing the Target Column letter reference _
    intRowCurProject = fx_Find_Row(ws:=wsProjects, strTarget:=strProjName, strTargetCol:="B")

' Change Log:
'       12/26/2021: Initial Creation
'       12/27/2021: Made the int_LastRow more dynamic, and added the 1 to capture a blank row
'       1/19/2022:  Added the code to allow strTargetCol to be passed
'       3/6/2022:   Added Error Handling around the Dictionary to allow duplicates
'                   Updated to handle situations where strTargetCol AND strTargetFieldName are not passed

' ****************************************************************************
  
' -----------------
' Declare Variables
' -----------------

    ' Declare Header Variables
    
    Dim arry_Header_Data() As Variant
        arry_Header_Data = Application.Transpose(ws.Range(ws.Cells(1, 1), ws.Cells(1, 99)))
        
    Dim col_Target As Long
        If strTargetCol <> "" Then
            col_Target = ws.Range(strTargetCol & "1").Column
        ElseIf strTargetFieldName <> "" Then
            col_Target = fx_Create_Headers_v2(strTargetFieldName, arry_Header_Data)
        Else
            col_Target = 1
        End If
    
    ' Declare Other Variables

    Dim int_LastRow As Long
        int_LastRow = WorksheetFunction.Max( _
        ws.Cells(Rows.Count, col_Target).End(xlUp).Row, _
        ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row) + 1
        
    Dim arryData() As Variant
        arryData = ws.Range(ws.Cells(1, col_Target), ws.Cells(int_LastRow, col_Target))

    Dim dictData As New Scripting.Dictionary
        dictData.CompareMode = TextCompare
        
    Dim i As Long
        
' -------------------
' Fill the Dictionary
' -------------------
    
On Error Resume Next
    
    For i = 1 To UBound(arryData)
        dictData.Add Key:=arryData(i, 1), Item:=i
    Next i
    
On Error GoTo 0
    
' --------------------
' Find the Current Row
' --------------------
    
    fx_Find_Row = dictData(strTarget)

End Function
Public Function fx_File_Exists(strFullPath As String) As Boolean

' Purpose: This function will determine if a file exists already.
' Trigger: Called
' Updated: 8/19/2021

' Change Log:
'       8/19/2021: Initial Creation

' ********************************************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Dim Objects
    Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
' ---------------------------
' Check if passed file exists
' ---------------------------
    
    If objFSO.FileExists(strFullPath) = True Then fx_File_Exists = True

    'Release the Object
    Set objFSO = Nothing

End Function
Sub fx_Export_Data_for_RiskTrend_Database(wsTarget As Worksheet, strFileName As String, dtRunDate As Date)

' Purpose: To export the passed worksheet to be used in the Risk Trend Database.
' Trigger: Called
' Updated: 7/15/2022

' Change Log:
'       6/24/2022:  Initial Creation, based on 'u_Prep_File_For_Email' from my To Do macros
'       7/6/2022:   Added the code for the two date columns
'       7/15/2022:  Updated the file path to point to Axcel's folder, and added the MsgBox to confirm the files there
'                   Switched from Hyperlink => Shell to avoid the hyperlink security warning

' ********************************************************************************************************************************************************

' USE EXAMPLE: _
    Call fx_Export_Data_for_RiskTrend_Database( _
        wsTarget:=wsData, _
        strFileName:="POLICYEXCEPTIONS_" & "APR22", _
        dtRundate = "TEST")

' LEGEND MANDATORY:
'   wsTarget: The worksheet that will be exported for the process to work
'   strFileName: The name of the file to be used (ex. 'POLICYEXCEPTIONS_APR22')

' LEGEND OPTIONAL:
'   int_SourceHeaderRow: TBD

' ********************************************************************************************************************************************************

' -----------------------------------
' Determine if the process should run
' -----------------------------------

' Disabled on 7/15/2022, not necessary if I am already having Megan click a button to trigger
'Dim intRunProcess As Long
'    intRunProcess = MsgBox( _
'        Prompt:="Would you like to prep the '" & ActiveWorkbook.Name & "' workbook for the RT Database?", _
'        Title:="Prep File?", _
'        Buttons:=vbQuestion + vbYesNo)
'
'        If intRunProcess = 7 Then Exit Sub 'Abort if cancel was pushed

myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

    ' Declare Worksheets
    
    Dim wbSelected As Workbook
    Set wbSelected = ActiveWorkbook
    
    Dim ws As Worksheet

    ' Declare Integers
       
    Dim int_LastCol As Long
        int_LastCol = wsTarget.Cells(1, wsTarget.Columns.Count).End(xlToLeft).Column
    
    Dim int_LastRow As Long
        int_LastRow = WorksheetFunction.Max(wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row, _
                         wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row, _
                         wsTarget.Cells(wsTarget.Rows.Count, "C").End(xlUp).Row)
                         
    Dim i As Integer
    
    ' Declare Strings

    Dim strNewFileName As String
        strNewFileName = strFileName & ".xlsx"
        
    Dim strNewFilePath As String
        strNewFilePath = "\\Hfd-data001\data\GRP\NHGRP\CREDAD\SecureCreditMIS\Combined Data Sets\Extracts"
    
    Dim strNewFileFullPath As String
        strNewFileFullPath = strNewFilePath & "\" & strNewFileName
    
    ' Declare Ranges
    
    Dim cell As Range
    
' --------------------------------------
' Copy back values only for the wsTarget
' --------------------------------------

    If wsTarget.AutoFilterMode = True Then wsTarget.AutoFilter.ShowAllData
        
    For Each cell In wsTarget.UsedRange
        If cell.HasFormula = True Then
            cell.Value2 = cell.Value2
        End If
    Next cell
            
' ---------------------------
' Delete all but the wsTarget
' ---------------------------

Application.DisplayAlerts = False

    For Each ws In wbSelected.Worksheets
        If ws.Name <> wsTarget.Name Then
            ws.Delete
        End If
    Next ws

Application.DisplayAlerts = True

' -----------------------
' Add the two date columns
' -----------------------

With wsTarget

    ' Month Year
    .Cells(1, int_LastCol + 1).Value = "Month Year"
        .Range(.Cells(2, int_LastCol + 1), .Cells(int_LastRow, int_LastCol + 1)).Value = CStr(Format(dtRunDate, "MMMM YYYY"))
        .Range(.Cells(2, int_LastCol + 1), .Cells(int_LastRow, int_LastCol + 1)).NumberFormat = "MMMM YYYY"
    
    ' Date Exported
    .Cells(1, int_LastCol + 2).Value = "Date Exported"
        .Range(.Cells(2, int_LastCol + 2), .Cells(int_LastRow, int_LastCol + 2)).Value = CStr(Date)

End With

' -----------------
' Save the workbook
' -----------------

Application.DisplayAlerts = False

    wbSelected.SaveAs Filename:=strNewFileFullPath, FileFormat:=xlOpenXMLWorkbook
    
    'ThisWorkbook.FollowHyperlink (strFilePath)
    Call Shell("explorer.exe" & " " & strNewFilePath, vbNormalFocus)
    myPrivateMacros.DisableForEfficiencyOff
    
' ------------------------------------------------
' Confirm the extract file was created to the user
' ------------------------------------------------
    
    If fx_File_Exists(strNewFileFullPath) = True Then
    
        MsgBox Title:="Extract File Created", _
        Buttons:=vbOKOnly, _
        Prompt:="The extract file was created and can be found here:" & Chr(10) & Chr(10) _
        & strNewFileFullPath
           
    Else
    
        MsgBox Title:="Extract File Failed", _
        Buttons:=vbCritical, _
        Prompt:="The extract file failed to be created.  Please see James for a fix, or create the file manually."
    
    End If

    
Application.DisplayAlerts = True
    
    wbSelected.Close savechanges:=False

End Sub
Function fx_Convert_to_Values(ws_Target As Worksheet, str_TargetField_Name As String, Optional int_FirstRowofData As Long)
    
' Purpose: To convert the passed range from text to values.
' Trigger: Called
' Updated: 7/3/2023

' Change Log:
'       7/3/2023:   Initial Creation, based on 'u_Convert_to_Values' from my To Do
'                   Overhauled to be more dynamic / resiliant

' ***********************************************************************************************************************************

' USE EXAMPLE: _
    Call fx_Convert_to_Values(ws_Target:=ThisWorkbook.Worksheets("Data"), str_TargetField_Name:="Account Number / Loan Number")

' LEGEND MANDATORY:
'   ws_Target: The worksheet where the data to be converted resides.
'   str_TargetField_Name: The name of the field where the target data to be convereted resides.

' LEGEND OPTIONAL:
'   int_FirstRowofData: If passed then use this as the int_FirstRow instead of 2.

' Note:
'       7/3/2023: Using an Array to load / read the data took 1/2 the time as updating the range directly

' ********************************************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

With ws_Target

    ' Declare "Ranges" / Cell References
    
    Dim int_LastCol As Integer
        int_LastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
        If int_LastCol = 1 Then int_LastCol = 2
    
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol)))

    Dim col_Target As Integer
        col_Target = fx_Create_Headers_v2(str_TargetField_Name, arry_Header)
           
    Dim int_FirstRow As Long
        If int_FirstRowofData > 0 Then
            int_FirstRow = int_FirstRowofData
        Else
            int_FirstRow = 2
        End If
        
    Dim int_LastRow As Long
        int_LastRow = .Cells(Rows.Count, col_Target).End(xlUp).Row
        If int_LastRow = 1 Then int_LastRow = 2
    
    ' Declare Loop Variables
    
    Dim rng_TargetData As Range
    Set rng_TargetData = .Range(.Cells(int_FirstRow, col_Target), .Cells(int_LastRow, col_Target))
    
    Dim arry_TargetData() As Variant
        arry_TargetData = WorksheetFunction.Transpose(rng_TargetData.Value)
    
    Dim i As Long
    
End With

' -----------------------------------
' Attempt entire selection conversion
' -----------------------------------

On Error GoTo IndividualProcess

    rng_TargetData.NumberFormat = "0"
    rng_TargetData.Value = rng_TargetData.Value
    
Exit Function
    
' ---------------------------------------
' Do conversion at invidiual record level
' ---------------------------------------
    
IndividualProcess:

On Error Resume Next

    Debug.Print "There was an error in the 'fx_Convert_to_Values' function at " & Now & Chr(10) _
                & "As a result the conversion had to be done at the individual record level."
    
    For i = LBound(arry_TargetData) To UBound(arry_TargetData)
        If IsNumeric(arry_TargetData(i)) = True Then
            rng_TargetData(i, 1).Value = val(arry_TargetData(i))
        End If
    Next i
    
End Function
