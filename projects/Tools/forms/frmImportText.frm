VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportText 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5100
   OleObjectBlob   =   "frmImportText.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImportText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBrowse_Click()
    Dim sh As Object
    Set sh = CreateObject("Shell.Application")

    Dim f As Object
    Set f = sh.BrowseForFolder(0, "Выберите папку с TXT-файлами", 0)

    If f Is Nothing Then Exit Sub

    txtFolder.Text = f.Self.path
    FillFileList f.Self.path
End Sub


Private Sub FillFileList(ByVal folderPath As String)
    Dim fn As String
    lstFiles.Clear

    fn = Dir(folderPath & "\*.txt")
    Do While fn <> ""
        lstFiles.AddItem fn
        fn = Dir
    Loop
End Sub


Private Sub btnRun_Click()
    If txtFolder.Text = "" Then Exit Sub
    If lstFiles.ListIndex = -1 Then Exit Sub

    Dim fullPath As String
    fullPath = txtFolder.Text & "\" & lstFiles.Value

    BuildParagraphTextFromFile fullPath
End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub
