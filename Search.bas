Attribute VB_Name = "Module1"
Private Const notEmptySymbol As String = "+"
Private Const max_rows As Integer = 200
Private Const start_row As Integer = 5
Private Const end_row As Integer = 120
Private Const stats_col_begin As Integer = 3
Private Const bonuses_col_begin As Integer = 18
Private Const bonuses_col_end As Integer = bonuses_col_begin + 6
Private searchText As Variant
Private excludeText As String


Sub SearchAdvanced()
    Dim hidden As Boolean
    excludeText = Cells(3, bonuses_col_end + 2)
    excludeEnabled = Len(excludeText) > 0
    For Row = start_row To end_row
        hidden = False
        For Each cell In Range(Cells(Row, bonuses_col_begin), Cells(Row, bonuses_col_end))
            searchText = Cells(3, cell.Column)
            If Len(searchText) < 1 Then GoTo ContinueLoop
                        
'            If (searchText = notEmptySymbol And cell.value = "") Or InStr(1, cell, searchText) = 0 Or excludeEnabled And InStr(1, cell, excludeText) <> 0 Then
'                hidden = True
'                Exit For
'            End If
            
            'empty check
            If searchText = notEmptySymbol Then
                If cell.value = "" Then
                    hidden = True
                    Exit For
                End If
            'not found
            ElseIf InStr(1, cell, searchText) = 0 Then
                hidden = True
                Exit For
            'found, check exclude
            ElseIf excludeEnabled And InStr(1, cell, excludeText) <> 0 Then
                hidden = True
                Exit For
            End If


ContinueLoop:
        Next
        'MsgBox hidden
        Rows(Row).hidden = hidden
    Next
End Sub

Sub Search()

    Dim arr(max_rows) As Boolean
    searchText = Cells(3, bonuses_col_end + 1)
    excludeText = Cells(3, bonuses_col_end + 2)
    excludeEnabled = Len(excludeText) > 0
    For Row = start_row To end_row

        Rows(Row).hidden = False
        For Each cell In Range(Cells(Row, bonuses_col_begin), Cells(Row, bonuses_col_end))

            'exclude
            If excludeEnabled And InStr(1, cell, excludeText) <> 0 Then
                GoTo ContinueLoop
            End If

            If InStr(1, cell, searchText) <> 0 Then arr(Row) = True
            If arr(Row) = True Then Exit For
ContinueLoop:
        Next
        If arr(Row) = False Then Rows(Row).hidden = True
    Next
End Sub
Sub ClearAdvanced()
    For Each cell In Range(Cells(3, bonuses_col_begin), Cells(3, bonuses_col_end))
        cell.value = ""
    Next
End Sub

Sub ShowAll()
    For Row = start_row To end_row
        Rows(Row).hidden = False
    Next
End Sub

Function ShowStats(value As Boolean)
    For col = stats_col_begin To bonuses_col_begin - 1
        Columns(col).hidden = Not value
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

