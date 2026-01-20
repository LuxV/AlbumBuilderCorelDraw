VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGenAlbum 
   Caption         =   "Создание АИ"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5730
   OleObjectBlob   =   "frmGenAlbum.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGenAlbum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Sub btnBrowse_Click()
    Dim sh As Object
    Set sh = CreateObject("Shell.Application")

    Dim f As Object
    Set f = sh.BrowseForFolder(0, "Выберите папку", 0)

    If Not f Is Nothing Then
        txtPath.Text = f.Self.Path
    End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnStart_Click()
    ' Проверка пути
    If Trim(txtPath.Text) = "" Then
        MsgBox "Пожалуйста, укажите путь к корневой папке.", vbExclamation
        txtPath.SetFocus
        Exit Sub
    End If
    
    If Dir(Trim(txtPath.Text), vbDirectory) = "" Then
        MsgBox "Указанная папка не существует.", vbExclamation
        txtPath.SetFocus
        Exit Sub
    End If
    
    ' Проверка имени объекта
    If Trim(txtObjectName.Text) = "" Then
        MsgBox "Пожалуйста, введите название объекта.", vbExclamation
        txtObjectName.SetFocus
        Exit Sub
    End If

    
    Dim StartIndex As Integer
    StartIndex = CLng(Val(txtStartIndex.Text))
    If StartIndex < 1 Then StartIndex = 1
    
    ' Всё проверено, скрываем форму и запускаем обработку
    Me.Hide
    BuildAlbum Trim(txtPath.Text), Trim(txtObjectName.Text), chkOnlyPhotos.Value, StartIndex
    Unload Me
End Sub

Private Sub chkOnlyPhotos_Change()

    lbStartIndex.Visible = chkOnlyPhotos.Value
    txtStartIndex.Visible = chkOnlyPhotos.Value

End Sub


Private Sub UserForm_Initialize()

    lbStartIndex.Visible = False
    txtStartIndex.Visible = False

End Sub
