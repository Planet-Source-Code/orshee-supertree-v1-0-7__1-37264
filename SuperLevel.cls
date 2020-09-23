VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"SuperNode"
Attribute VB_Ext_KEY = "Member0" ,"SuperNode"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarNodesCollection As Collection

Public Sub AddNode(New_Node As SuperNode, Optional Before)
    If IsMissing(Before) Then
        mvarNodesCollection.Add New_Node
    Else
        mvarNodesCollection.Add New_Node, , Before
    End If
End Sub

Public Property Get Nodes(Index As Variant) As SuperNode
Attribute Nodes.VB_UserMemId = 0
  Set Nodes = mvarNodesCollection(Index)
End Property

Public Property Get NodesCount() As Long
    NodesCount = mvarNodesCollection.Count
End Property

Public Sub RemoveNode(Index As Variant)
    mvarNodesCollection.Remove Index
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mvarNodesCollection.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mvarNodesCollection = New Collection
End Sub

Private Sub Class_Terminate()
    Set mvarNodesCollection = Nothing
End Sub

