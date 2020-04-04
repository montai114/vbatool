VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValueCheckForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ValueCheckForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ValueCheckForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SelectFolderButton_Click()
    ValueCheckForm.FolderPath.Caption = selectFolder()
End Sub

Private Sub CheckButton_Click()
    Dim errorMsg As String
    errorMsg = checkForm()
    If errorMsg <> "" Then
        MsgBox errorMsg
    Else
        Call checkExcel
    End If
End Sub

Function selectFolder() As Variant
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> 0 Then
            selectFolder = .SelectedItems(1)
        End If
    End With
End Function

Function checkExcel()
    Dim buf As String, cnt As Long
    Dim myFolder As Variant
    'myFoler example:D:\vba\test 最後に\つかない
    myFolder = ValueCheckForm.FolderPath.Caption & "\"
    buf = Dir(myFolder & "*.xls*")
    Do While buf <> ""
        Workbooks.Open Filename:=(myFolder + buf)
        cnt = cnt + 1
        Sheet1.Cells(cnt, 1) = buf
        Sheet1.Cells(cnt, 2) = Workbooks(buf).Worksheets(1).Cells(CInt(ValueCheckForm.RowTextBox.Text), CInt(ValueCheckForm.ColumnTextBox.Text))
        Workbooks(buf).Close
        buf = Dir()
    Loop
    Columns("A:B").AutoFit
End Function

Function checkForm() As String
    'Dim errorMsg As String
    If ValueCheckForm.FolderPath.Caption = "" Or ValueCheckForm.FolderPath.Caption = "フォルダを指定してください。" Then
        checkForm = addVbCrLf("フォルダを指定してください。")
    End If
    If Not IsNumeric(ValueCheckForm.RowTextBox.Text) Then
        checkForm = checkForm & addVbCrLf("行は数値を入力してください。")
    End If
    If Not IsNumeric(ValueCheckForm.ColumnTextBox.Text) Then
        checkForm = checkForm & "列は数値を入力してください。"
    End If
End Function

Function addVbCrLf(ByVal msg As String) As String
    addVbCrLf = msg & vbCrLf
End Function
