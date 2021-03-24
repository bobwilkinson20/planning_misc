Sub select_team()
'
' select_team Macro
' scope down to a particular team
'
' Keyboard Shortcut: Ctrl+t
'
    Dim activeRow, ActiveCol As Integer
    activeRow = ActiveCell.Row
    activeColumn = ActiveCell.Column

    Dim teamName As String
    teamName = Cells(2, activeColumn).Value

    Const numTeams As Integer = 11
    Const firstTeamColumn As Integer = 8

    Dim i As Integer
    For i = 0 To numTeams
        If Not (activeColumn - firstTeamColumn) = i Then
            Columns(firstTeamColumn + i).Select
            Selection.EntireColumn.Hidden = True
            Columns(firstTeamColumn + numTeams + i).Select
            Selection.EntireColumn.Hidden = True
        End If
    Next i

  ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=7, Criteria1:="*" & teamName & "*"
  Cells(3, 1).Select
  Cells(3, activeColumn).Select


End Sub
Sub clear_team()
'
' clear_team Macro
'
' Keyboard Shortcut: Ctrl+n
'
    Dim activeRow, ActiveCol As Integer
    activeRow = ActiveCell.Row
    activeColumn = ActiveCell.Column

    Const numTeams As Integer = 11
    Const firstTeamColumn As Integer = 8

    Dim i As Integer
    For i = 0 To numTeams
        Columns(firstTeamColumn + i).Select
        Selection.EntireColumn.Hidden = False
        Columns(firstTeamColumn + numTeams + i).Select
        Selection.EntireColumn.Hidden = False
    Next i

  ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=7
  Cells(3, 1).Select
  Cells(3, activeColumn).Select

End Sub
