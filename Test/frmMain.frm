VERSION 5.00
Object = "*\A..\SPT_SuperTree.vbp"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin SPT_SuperTree.SuperTree SuperTree1 
      Align           =   1  'Align Top
      Height          =   5325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10665
      _ExtentX        =   11377
      _ExtentY        =   4101
      BackColor       =   -2147483643
      ConnectorType   =   1
      NodeSize        =   0
      ToolbarState    =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim myNode As SuperNode
    Set myNode = New SuperNode
    Dim myLevel As SuperLevel
    Set myLevel = New SuperLevel
    
    myNode.Caption = "Root"
    myNode.Value = 10000

    myLevel.AddNode myNode
    SuperTree1.AddLevel myLevel
    
    SuperTree1.ConnectorColor = RGB(30, 30, 130)
    SuperTree1.Refresh
End Sub

Private Sub Form_Resize()
    SuperTree1.Height = ScaleHeight
End Sub

Private Sub SuperTree1_NodeClick(LevelIndex As Long, NodeIndex As Long)
    Dim myNode(5) As New SuperNode
    Dim myLevel As New SuperLevel
    Dim i As Long
    
    'First check if there are sub levels and remove them
    For i = LevelIndex + 1 To SuperTree1.LevelsCount
        SuperTree1.RemoveLevel (LevelIndex + 1)
    Next i
    
    For i = 0 To 4
        myNode(i).Caption = "Node " & CStr(i + 1)
        myNode(i).Value = (i + 1) * 1000
        myNode(i).Percentage = (i + 1) * 10
        myNode(i).ParentNode = SuperTree1.SelectedNode
        myLevel.AddNode myNode(i)
    Next i
    SuperTree1.AddLevel myLevel
    SuperTree1.Refresh
End Sub
