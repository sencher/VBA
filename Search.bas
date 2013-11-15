Attribute VB_Name = "Module1"

Private Const max_rows As Integer = 200
Private Const start_row As Integer = 5
Private Const end_row As Integer = 120
Private Const stats_col As Integer = 3
Private Const start_col As Integer = 18

Private Const end_col As Integer = start_col + 6
Private searchText As Variant
Private excludeText As String

Sub Search()

    Dim arr(max_rows) As Boolean
    searchText = Cells(3, end_col + 1)
    excludeText = Cells(3, end_col + 2)
    excludeEnabled = Len(excludeText) > 0
    For Row = start_row To end_row
        'If Cells(Row, 3).Value = "" Then Exit For

        Rows(Row).Hidden = False
        For Each cell In Range(Cells(Row, start_col), Cells(Row, end_col))

            'exclude
            If excludeEnabled And InStr(1, cell, excludeText) <> 0 Then
                GoTo ContinueLoop
            End If

            If InStr(1, cell, searchText) <> 0 Then arr(Row) = True
            If arr(Row) = True Then Exit For
ContinueLoop:
        Next
        If arr(Row) = False Then Rows(Row).Hidden = True
    Next
End Sub

Sub ShowAll()
    For Row = start_row To end_row
        Rows(Row).Hidden = False
    Next
End Sub
Function ShowStats(value As Boolean)
    For col = stats_col To start_col - 1
        Columns(col).Hidden = Not value
    Next
End Function

Function CalcSum(ByVal FirstArg As Integer, ParamArray OtherArgs())
Dim ReturnValue
' If the function is invoked as follows:

' Local variables are assigned the following values: FirstArg = 4,
' OtherArgs(1) = 3, OtherArgs(2) = 2, and so on, assuming default
' lower bound for arrays = 1.
End Function
Sub Test()
'some = CalcSum(4, 3, 2, 1)
Debug.Print searchText
'Dim MyVar
'MyVar = "Come see me in the Immediate pane."
'Debug.Print MyVar

End Sub
Function trace(ParamArray items() As Variant)
    Dim value As String
    For Each i In items
        value = value + " " + i
    Next
    MsgBox value
    
End Function

