Attribute VB_Name = "NewMacros"
Sub ReformatRTF()
'
' ReformatRTF Macro
'
'

Dim IdColumn1
Set IdColumn1 = ActiveDocument.Tables(1).Columns(1)
Dim SourceColumn2
Set SourceColumn2 = ActiveDocument.Tables(1).Columns(2)
Dim TargetColumn3
Set TargetColumn3 = ActiveDocument.Tables(1).Columns(3)
Dim CommentsColumn4
Set CommentsColumn4 = ActiveDocument.Tables(1).Columns(4)
Dim StatusColumn5
Set StatusColumn5 = ActiveDocument.Tables(1).Columns(5)
Dim FilenameColumn6
Set FilenameColumn6 = ActiveDocument.Tables(1).Columns(6)
Dim HeaderRow
Set HeaderRow = ActiveDocument.Tables(1).Rows(1)

    With Selection.PageSetup
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(1)
        .RightMargin = CentimetersToPoints(1)
    End With

    IdColumn1.SetWidth ColumnWidth:=Application.CentimetersToPoints(1), RulerStyle:= _
        wdAdjustFirstColumn
    IdColumn1.Select
        Selection.Font.Size = 8
        Selection.Shading.BackgroundPatternColor = -603917569
    SourceColumn2.SetWidth ColumnWidth:=Application.CentimetersToPoints(5), RulerStyle:= _
        wdAdjustFirstColumn
    SourceColumn2.Select
        Selection.Font.Size = 9
    TargetColumn3.SetWidth ColumnWidth:=Application.CentimetersToPoints(5), RulerStyle:= _
        wdAdjustFirstColumn
    TargetColumn3.Select
        Selection.Font.Size = 9
    CommentsColumn4.SetWidth ColumnWidth:=Application.CentimetersToPoints(4), RulerStyle:= _
        wdAdjustFirstColumn
    CommentsColumn4.Select
        Selection.Font.Size = 8
    StatusColumn5.SetWidth ColumnWidth:=Application.CentimetersToPoints(1.5), RulerStyle:= _
        wdAdjustFirstColumn
    StatusColumn5.Select
        Selection.Font.Size = 8
    FilenameColumn6.SetWidth ColumnWidth:=Application.CentimetersToPoints(2.5), RulerStyle:= _
        wdAdjustFirstColumn
    FilenameColumn6.Select
        Selection.Font.Size = 8
        Selection.Shading.BackgroundPatternColor = -603917569
    HeaderRow.Select
        Selection.Font.Size = 9
        Selection.Font.Bold = True
        Selection.Shading.BackgroundPatternColor = -603923969

End Sub
