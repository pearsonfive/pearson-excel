Attribute VB_Name = "MFormatTable"
Option Explicit

Public Sub FormatTableWithTitle()
    formatTable True
End Sub
Public Sub FormatTableNoTitle()
    formatTable False
End Sub

Private Sub formatTable(Optional hasTitle As Boolean = True)
Dim rTitle As Range, rHeader As Range, rBody As Range

    '--- need a minimum number of rows
    If Selection.Rows.Count <= IIf(hasTitle, 3, 2) Then Exit Sub

    If hasTitle Then
        Set rTitle = Selection.Rows(1)
    End If

    Set rHeader = Selection.Rows(IIf(hasTitle, 2, 1))
    Set rBody = Selection.Offset(IIf(hasTitle, 2, 1)).Resize(Selection.Rows.Count - IIf(hasTitle, 2, 1))

    '--- outer borders
    If hasTitle Then
        rTitle.BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    End If
    rHeader.BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    rBody.BorderAround xlContinuous, xlThin, xlColorIndexAutomatic

    '--- fill colours
    If hasTitle Then
        With rTitle.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With
    End If
    With rHeader.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    With rBody.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .PatternTintAndShade = 0
    End With

    '--- interior borders
    With rHeader.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .ThemeColor = 1
        .TintAndShade = -0.14996795556505
        .Weight = xlThin
    End With
    With rBody.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .ThemeColor = 1
        .TintAndShade = -0.14996795556505
        .Weight = xlThin
    End With
    With rBody.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .ThemeColor = 1
        .TintAndShade = -0.14996795556505
        .Weight = xlThin
    End With
    
    '--- alignments
    If hasTitle Then
        rTitle.HorizontalAlignment = xlCenterAcrossSelection
    End If
    rHeader.HorizontalAlignment = xlCenter
    rBody.HorizontalAlignment = xlCenter

End Sub

