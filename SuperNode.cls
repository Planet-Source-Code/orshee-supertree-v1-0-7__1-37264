VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim mvarCaption As String
Dim mvarUniqueName As String
Dim mvarValue As Double
Dim mvarPercentage As Double

Dim mvarLeft As Long
Dim mvarTop As Long
Dim mvarRight As Long
Dim mvarBottom As Long
Dim mvarCenter As Long 'Used for conection lines

Dim mvarSelected As Boolean
Dim mvarParentNode As SuperNode

Public Property Let Caption(Value As String)
    mvarCaption = Value
End Property

Public Property Get Caption() As String
    Caption = mvarCaption
End Property

Public Property Let UniqueName(Value As String)
    mvarUniqueName = Value
End Property

Public Property Get UniqueName() As String
    UniqueName = mvarUniqueName
End Property

Public Property Let Value(New_Value As Double)
    mvarValue = New_Value
End Property

Public Property Get Value() As Double
    Value = mvarValue
End Property

Public Property Let Percentage(New_Value As Double)
    mvarPercentage = New_Value
End Property

Public Property Get Percentage() As Double
    Percentage = mvarPercentage
End Property

Public Property Let Selected(Value As Boolean)
    mvarSelected = Value
End Property

Public Property Get Selected() As Boolean
    Selected = mvarSelected
End Property

Public Property Let ParentNode(Node As SuperNode)
    Set mvarParentNode = Node
End Property

Public Property Get ParentNode() As SuperNode
    Set ParentNode = mvarParentNode
End Property

'***[Positioning properties]****************************************************************************************
'Not intended to be used by user but by mighty engine :)

Public Property Let Left(Value As Long)
    mvarLeft = Value
End Property

Public Property Get Left() As Long
    Left = mvarLeft
End Property

Public Property Let Top(Value As Long)
    mvarTop = Value
End Property

Public Property Get Top() As Long
    Top = mvarTop
End Property

Public Property Let Right(Value As Long)
    mvarRight = Value
End Property

Public Property Get Right() As Long
    Right = mvarRight
End Property

Public Property Let Bottom(Value As Long)
    mvarBottom = Value
End Property

Public Property Get Bottom() As Long
    Bottom = mvarBottom
End Property

Public Property Let Center(Value As Long)
    mvarCenter = Value
End Property

Public Property Get Center() As Long
    Center = mvarCenter
End Property

