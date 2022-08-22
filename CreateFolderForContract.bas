Attribute VB_Name = "CreateFolderForContract"
Option Explicit

Sub main()
    
End Sub

Sub createfolder()
    Dim obj As Object
    Dim i As Integer
    Dim pathToFile As String
    Dim arrayFolder() As Variant
    Dim item As Variant
    Dim initem As Variant
    Dim nameFold As String
    Dim objFolder As Object
    Dim OldNameFolder As String
    Dim workOBJ As Range
    
    nameFold = Cells(ActiveCell.Row, "O").Value
    
    Set workOBJ = Cells(ActiveCell.Row, "P")

    OldNameFolder = Cells(ActiveCell.Row, "R").Value
    
    arrayFolder = Array("����������", "����������", "������������", "���������� �������")

    pathToFile = ThisWorkbook.Worksheets("settings").Range("AddressToFiles")

    Set obj = CreateObject("Scripting.FileSystemObject")

    With obj
        If obj.FolderExists(pathToFile & OldNameFolder) Then
            If nameFold <> OldNameFolder Then
                On Error GoTo Balvak
                    Name pathToFile & OldNameFolder As pathToFile & nameFold
'                    Call create_Hyperlinks(workOBJ, pathToFile & nameFold)
                    MsgBox "Rename folder"
                On Error GoTo 0
            End If

        Else
            On Error GoTo Balvak
                .createfolder (pathToFile & nameFold & "\")
                 For Each item In arrayFolder
                    .createfolder (pathToFile & nameFold & "\" & item)
                    If item = "���������� �������" Then
                        For Each initem In Array("01_��", "02_������_����", "03_��", "04_���", "05_�����������", "06_��", "07_����_������������", "08_������_�� ������")
                            .createfolder (pathToFile & nameFold & "\" & item & "\" & initem)
                        Next initem
                    End If
                 Next item
                 Call create_Hyperlinks(workOBJ, pathToFile & nameFold)
                 MsgBox "Create new folder"
'                 Exit Sub
            On Error GoTo 0
Balvak:
                MsgBox Err.Description, vbCritical
            
        End If
        
' TODO ����������� ���������� ����� ����� ��� �������
' ������� ��������������� ����� � ����� ������
' �������� ����� ����� ������� ��� �������� �������� ����� (1 ��� ������������, 2 ��� �������) � ����������� �� �������

    End With
    workOBJ.Offset(0, 2).Value = nameFold
    Call create_Hyperlinks(workOBJ, pathToFile & nameFold)
    Call set_info_of_availability(obj, workOBJ, pathToFile & nameFold & "\" & "���������� �������" & "\" & "08_������_�� ������" & "\")
End Sub

Sub create_Hyperlinks(ByVal h_Cell As Range, ByVal h_Address As String)

    ThisWorkbook.Worksheets("main").Hyperlinks.Add Anchor:=h_Cell, Address:=h_Address, TextToDisplay:="Clik!"

End Sub


Sub set_info_of_availability(ByVal FSO As Object, ByVal av_Cell As Range, ByVal pathToDirect As String)
    
    If FSO.GetFolder(pathToDirect).Files.Count > 0 Then
        Cells(av_Cell.Row, "Q").Value = "+"
    Else
        Cells(av_Cell.Row, "Q").Value = "-"
    End If
End Sub


