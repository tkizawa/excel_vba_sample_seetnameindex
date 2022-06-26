Attribute VB_Name = "Module1"
Option Explicit

Sub 目次作成()
    Dim sheet1 As Object
    Dim count As Integer
    
    With Worksheets("目次")
        .Cells(1, 1).Value = "目次"
        count = 2
        For Each sheet1 In Worksheets
            ' 目次シートを目次から除外する
            If sheet1.Name <> "目次" Then
                .Hyperlinks.Add Anchor:=.Cells(count, 1), _
                Address:="", _
                SubAddress:="'" + sheet1.Name + "'" + "!A1", _
                TextToDisplay:=sheet1.Name
                count = count + 1
            End If
        Next sheet1
    End With
End Sub
