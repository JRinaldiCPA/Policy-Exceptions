VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_wsSelector 
   Caption         =   "Select the Worksheet to import"
   ClientHeight    =   2388
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4908
   OleObjectBlob   =   "uf_wsSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_wsSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lst_Worksheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    ' Pass the selected value to the Public Variable
    str_wsSelected = lst_Worksheets.Value
    Unload Me

End Sub
Private Sub UserForm_Initialize()
    
' Purpose: To create a list of worksheets for the user to choose from.
' Trigger: Called
' Updated: 7/2/2021

' Change Log:
'       7/2/2021: Initial Creation
'       7/2/2021: Added the code to only be applicable to visible sheets
'       7/4/2021: Added code to resize depending on the count of sheets

' ****************************************************************************
    
' -----------------
' Declare Variables
' -----------------
    
    Dim ws As Worksheet
    
' ----------------------------------
' Add the worksheets to the list box
' ----------------------------------
    
    For Each ws In ActiveWorkbook.Sheets
    'For Each ws In wbSelectWorksheet.Sheets
        If ws.Visible = xlSheetVisible Then
            uf_wsSelector.lst_Worksheets.AddItem ws.Name
        End If
    Next ws
    
' ----------------------------------
' Add the worksheets to the list box
' ----------------------------------
    
    uf_wsSelector.Height = (20 * ActiveWorkbook.Sheets.Count) + 50
    uf_wsSelector.lst_Worksheets.Height = 20 * ActiveWorkbook.Sheets.Count
    uf_wsSelector.lbl_Message.Top = (20 * ActiveWorkbook.Sheets.Count) + 10
    
End Sub
Private Sub cmd_Cancel_Click()
    
    Unload Me

End Sub

