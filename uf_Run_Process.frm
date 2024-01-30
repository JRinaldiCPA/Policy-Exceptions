VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Run_Process 
   Caption         =   "Sageworks Dashboard - Admin User"
   ClientHeight    =   8676.001
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   11832
   OleObjectBlob   =   "uf_Run_Process.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Run_Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declare Worksheets
    Dim wsData As Worksheet
    Dim wsDetailData As Worksheet
    Dim wsPrior As Worksheet
    
    Dim wsLists As Worksheet
    Dim wsArrays As Worksheet
    Dim wsChangeLog As Worksheet
    Dim wsUpdates As Worksheet
    Dim wsChecklist As Worksheet
    Dim wsValidation As Worksheet
    Dim wsFormulas As Worksheet
    Dim wsPivot As Worksheet

' Declare Strings
    Dim strCustomer As String
    Dim strNewFileFullPath As String 'Used by o_51_Create_a_XLSX_Copy to create the XLSX to attach to the email
    Dim strLastCol_wsData As String

' Declare Integers
    Dim intLastRow As Long
    Dim intLastRow_wsArrays As Long
    
    Dim intLastCol As Integer
    Dim intLastCol_wsLists As Integer
    
' Declare Data "Ranges"
    Dim col_Customer As Integer
    Dim col_CustID As Integer
    Dim col_LOB As Integer
    Dim col_Region As Integer
    Dim col_PM_Name As Integer
    Dim col_AM_Name As Integer
    
    Dim col_BRG As Integer
    Dim col_FRG As Integer
    Dim col_CCRP As Integer
    
    Dim col_Exposure As Integer
    Dim col_Outstanding As Integer
    
    Dim col_PM_Attest As Integer
    Dim col_PM_Attest_Exp As Integer
    Dim col_CovCompl As Integer
    Dim col_CovCompl_Exp As Integer
    Dim col_ChangeFlag As Integer
    
' Declare List "Ranges"

    Dim col_LOB1_Lists As Integer
    Dim col_Cust1_Lists As Integer
    Dim col_PM1_Lists As Integer
    
' Declare Arrays / Other
    
    Dim arryHeader_Data() As Variant
    Dim arryHeader_Lists() As Variant

    Dim ary_Customers
    Dim ary_PM
    
    Dim ary_Lists_PMLookup

' Declare Dictionaries
    Dim dict_PMs As Scripting.Dictionary

' Declare "Booleans"
    Dim bolPrivilegedUser As Boolean
    
    Dim bol_wsDetailData_Exists As Boolean
    Dim bol_wsPrior_Exists As Boolean
    
    Dim bol_AttestationStatus As String
    Dim bol_Edit_Filter As Boolean
    Dim bol_QC_Flags As Boolean
    Dim bol_CovCompliance As String
    
Option Explicit

Private Sub UserForm_Initialize()
 
' Purpose:  To initialize the userform, including adding in the data from the arrays.
' Trigger:  Workbook Open
' Updated:  11/16/2020
' Author:   James Rinaldi

' Change Log:
'       3/23/2020: Initial Creation
'       8/19/2020: Added the logic to exclude the exempt customers
'       11/16/2020: I added the autofilter to hide the CRE and ABL LOBs
'       12/29/2020: Moved the CRE and ABL filtering to o_64_Update_Workbook_for_Admins

' ****************************************************************************

'Call Me.o_02_Declare_Global_Variables

'Call Me.o_03_Declare_Global_Arrays

    Call myPublicVariables.o_1_Assign_Global_Variables

' -----------
' Initialize the initial values
' -----------
    Me.StartUpPosition = 0 'Allow you to set the position
        Me.Top = Application.Top + (Application.UsableHeight / 1.5) - (Me.Height / 2) 'Open near the bottom of the screen
        Me.Left = Application.Left + (Application.UsableWidth / 2) - (Me.Width / 2)

End Sub
Private Sub cmd_Import_CDS_Data_Click()

    Call o1_Import_Comb_Data_Set_Data.o_01_MAIN_PROCEDURE

End Sub

Private Sub cmd_Import_Chained_Report_Click()
    
    Call o2_Import_Sageworks_Data.o_01_MAIN_PROCEDURE
    
End Sub
Private Sub cmd_Import_BB_Exceptions_Click()

    Call o3_Import_BB_Data.o_01_MAIN_PROCEDURE

    End Sub
Private Sub cmd_Manipulate_Data_Sources_Click()

    Call o4_Manipulate_Data_Source.o_01_MAIN_PROCEDURE

End Sub
Private Sub cmd_Merge_Exception_Data_Click()

    Call o5_Merge_Exception_Data.o_01_MAIN_PROCEDURE
    
End Sub
Private Sub cmd_Create_Policy_Exceptions_Click()

    Call o7_Create_Policy_Exceptions_ws.o_01_MAIN_PROCEDURE
    
End Sub
Private Sub cmd_Export_Data_Click()
    
    Call Macros.Export_Data_for_Combined_Data_Set
    
End Sub
Private Sub cmd_Cancel_Click()

    Unload Me

End Sub
