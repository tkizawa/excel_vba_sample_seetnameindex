Attribute VB_Name = "Module1"
Option Explicit

Sub �ڎ��쐬()
    Dim sheet1 As Object
    Dim count As Integer
    
    With Worksheets("�ڎ�")
        .Cells(1, 1).Value = "�ڎ�"
        count = 2
        For Each sheet1 In Worksheets
            ' �ڎ��V�[�g��ڎ����珜�O����
            If sheet1.Name <> "�ڎ�" Then
                .Hyperlinks.Add Anchor:=.Cells(count, 1), _
                Address:="", _
                SubAddress:="'" + sheet1.Name + "'" + "!A1", _
                TextToDisplay:=sheet1.Name
                count = count + 1
            End If
        Next sheet1
    End With
End Sub
