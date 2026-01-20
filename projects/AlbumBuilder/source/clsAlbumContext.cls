VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAlbumContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IllNumber As Integer
Public ObjectName As String

Public doc As Document
Public currentPage As Page

Public StructureType As String
Public OnlyPhotos As Boolean

Public Sub Init(ByVal doc As Document, _
                ByVal objectName As String, _
                ByVal startIndex As Long, _
                ByVal structureType As String, _
                ByVal onlyPhotos As Boolean)

    Set Me.doc = doc
    Me.objectName = objectName
    Me.IllNumber = startIndex
    Me.structureType = structureType
    Me.onlyPhotos = onlyPhotos
End Sub