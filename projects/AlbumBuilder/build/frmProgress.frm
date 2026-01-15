VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8310.001
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cancelled As Boolean

Public Sub Init(ByVal maxFolders As Long)
    lblFolderText.caption = "0 / " & maxFolders
    lblFileText.caption = "0 / 0"
    lblFolderProgress.Width = 0
    lblFileProgress.Width = 0
    lblStatus.caption = "Запуск..."
    Cancelled = False
    Me.Show vbModeless
    DoEvents
End Sub

Public Sub UpdateFolder(ByVal current As Long, ByVal total As Long)
    lblFolderText.caption = current & " / " & total
    lblFolderProgress.Width = (fraFolders.Width - 50) * current / total
    DoEvents
    'If current Mod 2 = 0 Then DoEvents  ' вызываем DoEvents не на каждой итерации, а через каждые 2
End Sub

Public Sub UpdateFiles(ByVal current As Long, ByVal total As Long)
    lblFileText.caption = current & " / " & total
    lblFileProgress.Width = (fraFiles.Width - 50) * current / total
    DoEvents
    'If current Mod 5 = 0 Then DoEvents  ' экономно, каждые 5 файлов
End Sub

Public Sub SetStatus(ByVal txt As String)
    lblStatus.caption = txt
    DoEvents
End Sub

Private Sub btnCancel_Click()

    Cancelled = True
    lblStatus.caption = "Отмена..."
    DoEvents
    
End Sub
