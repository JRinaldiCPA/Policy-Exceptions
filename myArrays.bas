Attribute VB_Name = "myArrays"
Dim arryCurExceptions(1 To 6) As String

Option Explicit
Public Function Get_arryCurExceptions() As Variant
    
    ' Create an array of all of the Current Policy Exceptions
    
    arryCurExceptions(1) = "1 - Collateral & Other Support"
    arryCurExceptions(2) = "2 - Debt Repayment Capacity / Liquidity"
    arryCurExceptions(3) = "3 - Financial Covenants"
    arryCurExceptions(4) = "4 - Maximum Amortization / Ability to Amortize"
    arryCurExceptions(5) = "5 - Maximum Tenor"
    arryCurExceptions(6) = "House Limit Exposure (HLE) Breach"

    Get_arryCurExceptions = arryCurExceptions
    
End Function


