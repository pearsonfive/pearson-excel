Attribute VB_Name = "MContextMenu"
Option Explicit

Private Const TAG_FORMAT_TABLE_CONTEXT_BUTTON As String = "format_table_context_button"

Public Sub AddToContextMenu()
Dim contextMenu As CommandBar

    '--- remove first
    Call removeExistingEntries

    Set contextMenu = Application.CommandBars("Cell")

    With contextMenu.Controls.Add(Type:=msoControlButton, before:=5, temporary:=True)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatTableNoTitle"
        .FaceId = 634
        .Caption = "Format table (no title)"
        .Tag = TAG_FORMAT_TABLE_CONTEXT_BUTTON
    End With
    With contextMenu.Controls.Add(Type:=msoControlButton, before:=5, temporary:=True)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatTableWithTitle"
        .FaceId = 498
        .Caption = "Format table (with title)"
        .Tag = TAG_FORMAT_TABLE_CONTEXT_BUTTON
        .BeginGroup = True
    End With
    
End Sub

Private Sub removeExistingEntries()
Dim contextMenu As CommandBar
Dim ctrl As CommandBarControl

    Set contextMenu = Application.CommandBars("Cell")

    For Each ctrl In contextMenu.Controls
        If ctrl.Tag = TAG_FORMAT_TABLE_CONTEXT_BUTTON Then
            ctrl.Delete
        End If
    Next ctrl

End Sub
