VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hsRetVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mName As String
Private mContent As String
Public Property Let ValueName(nName As String)
mName = nName
End Property

Public Property Get ValueName() As String
ValueName = mName
End Property

Public Property Let ValueContent(nContent As String)
mContent = nContent
End Property

Public Property Get ValueContent() As String
ValueContent = mContent
End Property

