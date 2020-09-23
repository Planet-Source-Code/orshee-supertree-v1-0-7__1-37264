VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl SuperTree 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   ScaleHeight     =   4635
   ScaleWidth      =   7605
   ToolboxBitmap   =   "SuperTree.ctx":0000
   Begin SPT_SuperTree.SuperViewPort SuperViewPort 
      Align           =   1  'Align Top
      Height          =   3675
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   6482
      BorderStyle     =   0
      ViewPortWidth   =   30000
      Begin VB.PictureBox PB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3285
         Left            =   1560
         ScaleHeight     =   3285
         ScaleWidth      =   4995
         TabIndex        =   2
         Top             =   180
         Width           =   5000
      End
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   5100
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SuperTree.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SuperTree.ctx":0697
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SuperTree.ctx":0A1D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnConnector"
            Object.ToolTipText     =   "Connector type"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "conStraight"
                  Text            =   "Straight Line"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "conStep"
                  Text            =   "Step Line"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "conSpray"
                  Text            =   "Spray"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnZoom"
            Object.ToolTipText     =   "Node size"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "zoomSmall"
                  Text            =   "Small"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "zoomMedium"
                  Text            =   "Medium"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "zoomLarge"
                  Text            =   "Large"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnGrowth"
            Object.ToolTipText     =   "Growth type"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "growLeft"
                  Text            =   "Grow from left"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "growSelected"
                  Text            =   "Grow from selected"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "SuperTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***[Node Settings]******************************************************************************************************
Dim NODE_FONT As String
Dim NODE_FONT_SIZE As Long
Dim NODE_HEIGHT As Long
Dim NODE_WIDTH As Long
Dim NODE_SHADOW_DEPTH As Long

Dim NODE_VSPACE As Long
Dim NODE_HSPACE As Long


'***[Enumerations]***************************************************************************************************
Public Enum enuAppearance
    apFlat = 0
    ap3D = 1
End Enum

Public Enum enuBorderStyle
    bsNoBorder = 0
    bsFixedSingle = 1
End Enum

Public Enum enuConnectorType
    ctStraightLine = 0
    ctStepLine = 1
    ctSpray = 2
End Enum

Public Enum enuTreeGrowth
    tgFromLeft = 0
    tgFromSelected = 1
End Enum

Public Enum enuNodeSize
    tzSmall = 0
    tzMedium = 1
    tzLarge = 2
End Enum

Public Enum enuToolbarState
    tsDisabled = 0
    tsActive = 1
End Enum

'***[Variables]******************************************************************************************************
Dim mvarLevels As Collection
Dim mvarSelectedNode As SuperNode
Dim mvarConnectorColor As Long 'Connector color
Dim mvarConnectorType As enuConnectorType 'Connector type : either Line or Spray
Dim mvarTreeGrowth As enuTreeGrowth 'Determines how tree nodes will be placed
Dim mvarNodeSize As enuNodeSize 'Determines size of node and its font
Dim mvarToolbarState As enuToolbarState  'Determines visibility of toolbar

Dim mvarStartX As Long 'Start coordinates for SuperTree
Dim mvarStartY As Long

Dim mvarMaxWidth 'Maximal width of tree


'***[Events]*********************************************************************************************************
Public Event NodeClick(LevelIndex As Long, NodeIndex As Long)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'***[Properties]*****************************************************************************************************
Public Property Get Appearance() As enuAppearance
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal Value As enuAppearance)
    UserControl.Appearance() = Value
    PropertyChanged "Appearance"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = PB.BackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    PB.BackColor = Value
    SuperViewPort.BackColor = Value
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As enuBorderStyle
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As enuBorderStyle)
    UserControl.BorderStyle() = Value
    PropertyChanged "BorderStyle"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    UserControl.Enabled() = Value
    PropertyChanged "Enabled"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    DrawTree
End Sub

Public Property Let SelectedNode(Node As SuperNode)
    Set mvarSelectedNode = Node
End Property

Public Property Get SelectedNode() As SuperNode
    Set SelectedNode = mvarSelectedNode
End Property

Public Property Let ConnectorColor(Value As Long)
    mvarConnectorColor = Value
    PropertyChanged "ConnectorColor"
End Property

Public Property Get ConnectorColor() As Long
    ConnectorColor = mvarConnectorColor
End Property

Public Property Let ConnectorType(Value As enuConnectorType)
    mvarConnectorType = Value
    PropertyChanged "ConnectorType"
End Property

Public Property Get ConnectorType() As enuConnectorType
    ConnectorType = mvarConnectorType
End Property

Public Property Let TreeGrowth(Value As enuTreeGrowth)
    mvarTreeGrowth = Value
    PropertyChanged "TreeGrowth"
End Property

Public Property Get TreeGrowth() As enuTreeGrowth
    TreeGrowth = mvarTreeGrowth
End Property

Public Property Let NodeSize(Value As enuNodeSize)
    mvarNodeSize = Value
    PropertyChanged "NodeSize"
End Property

Public Property Get NodeSize() As enuNodeSize
    NodeSize = mvarNodeSize
End Property

Public Property Let ToolbarState(Value As enuToolbarState)
    mvarToolbarState = Value
    PropertyChanged "ToolbarState"
End Property

Public Property Get ToolbarState() As enuToolbarState
    ToolbarState = mvarToolbarState
End Property

Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "BTNCONNECTOR"
        Select Case ConnectorType
        Case ctStraightLine
            ConnectorType = ctStepLine
            Refresh
        Case ctStepLine
            ConnectorType = ctSpray
            Refresh
        Case ctSpray
            ConnectorType = ctStraightLine
            Refresh
        End Select
    Case "BTNZOOM"
        Select Case NodeSize
        Case tzSmall
            NodeSize = tzMedium
            InitSettings
            Refresh
        Case tzMedium
            NodeSize = tzLarge
            InitSettings
            Refresh
        Case tzLarge
            NodeSize = tzSmall
            InitSettings
            Refresh
        End Select
    Case "BTNGROWTH"
        Select Case TreeGrowth
        Case tgFromLeft
            TreeGrowth = tgFromSelected
            Refresh
        Case tgFromSelected
            TreeGrowth = tgFromLeft
            Refresh
        End Select
    End Select
    
End Sub

Private Sub ToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case UCase(ButtonMenu.Key)
    'Set Connector Type
    Case "CONSTRAIGHT"
        ConnectorType = ctStraightLine
        Refresh
    Case "CONSTEP"
        ConnectorType = ctStepLine
        Refresh
    Case "CONSPRAY"
        ConnectorType = ctSpray
        Refresh
    'Set Zoom level
    Case "ZOOMSMALL"
        NodeSize = tzSmall
        InitSettings
        Refresh
    Case "ZOOMMEDIUM"
        NodeSize = tzMedium
        InitSettings
        Refresh
    Case "ZOOMLARGE"
        NodeSize = tzLarge
        InitSettings
        Refresh
    'Set Growth type
    Case "GROWLEFT"
        TreeGrowth = tgFromLeft
        Refresh
    Case "GROWSELECTED"
        TreeGrowth = tgFromSelected
        Refresh
    End Select
End Sub

'***[Life Control]***************************************************************************************************
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    PB.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    SuperViewPort.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    ConnectorColor = PropBag.ReadProperty("ConnectorColor", 0)
    ConnectorType = PropBag.ReadProperty("ConnectorType", ctStraightLine)
    TreeGrowth = PropBag.ReadProperty("TreeGrowth", tgFromLeft)
    NodeSize = PropBag.ReadProperty("NodeSize", tzMedium)
    ToolbarState = PropBag.ReadProperty("ToolbarState", tsDisabled)
    
    InitSettings
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", PB.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("ConnectorColor", ConnectorColor, 0)
    Call PropBag.WriteProperty("ConnectorType", ConnectorType, ctStraightLine)
    Call PropBag.WriteProperty("TreeGrowth", TreeGrowth, tgFromLeft)
    Call PropBag.WriteProperty("NodeSize", NodeSize, tzMedium)
    Call PropBag.WriteProperty("ToolbarState", ToolbarState, tsDisabled)
End Sub

Private Sub InitSettings()
    NODE_FONT = "Arial CE"
    
    Select Case NodeSize
    Case tzSmall
        NODE_FONT_SIZE = 7
        NODE_HEIGHT = 450
        NODE_WIDTH = 1800
        NODE_SHADOW_DEPTH = 15
        
        NODE_VSPACE = 900
        NODE_HSPACE = 2000
    Case tzMedium
        NODE_FONT_SIZE = 8
        NODE_HEIGHT = 500
        NODE_WIDTH = 2000
        NODE_SHADOW_DEPTH = 30
        
        NODE_VSPACE = 1000
        NODE_HSPACE = 2200
    Case tzLarge
        NODE_FONT_SIZE = 9
        NODE_HEIGHT = 550
        NODE_WIDTH = 2200
        NODE_SHADOW_DEPTH = 45
        
        NODE_VSPACE = 1100
        NODE_HSPACE = 2400
    End Select

    PB.FontName = NODE_FONT
    PB.FontSize = NODE_FONT_SIZE
    PB.Move 0, 0
    
    If ToolbarState = tsActive Then
        ToolBar.Visible = True
    Else
        ToolBar.Visible = False
    End If
End Sub

Private Sub UserControl_Resize()
    Dim myTmp As Long
    If ToolbarState = tsActive Then
        If ScaleHeight > ToolBar.Height Then
            SuperViewPort.Height = ScaleHeight - ToolBar.Height
        End If
    Else
        SuperViewPort.Height = ScaleHeight
    End If
    
    myTmp = PB.Width - ScaleWidth + SuperViewPort.VScrollWidth
    If myTmp < 0 Then myTmp = 0
    SuperViewPort.ViewPortWidth = myTmp
    
    myTmp = PB.Height - ScaleHeight + SuperViewPort.HScrollHeight
    If myTmp < 0 Then myTmp = 0
    SuperViewPort.ViewPortHeight = myTmp
    'Refresh
End Sub

Private Sub UserControl_Initialize()
    Set mvarLevels = New Collection
    mvarStartX = 100
    mvarStartY = 100
End Sub

Private Sub UserControl_Terminate()
    Set mvarLevels = Nothing
End Sub

Public Property Get Hdc() As Long
    Hdc = UserControl.Hdc
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub PB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub PB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If LevelsCount > 0 Then
        Dim i As Long
        Dim j As Long
        Dim myNode As New SuperNode
        For i = 1 To LevelsCount
            For j = 1 To Levels(i).NodesCount
                Set myNode = Levels(i).Nodes(j)
                If X >= myNode.Left And X <= myNode.Right And Y >= myNode.Top And Y <= myNode.Bottom Then
                    If Not SelectedNode Is Nothing Then
                        SelectedNode.Selected = False
                    End If
                    myNode.Selected = True
                    SelectedNode = myNode
                    RaiseEvent NodeClick(i, j)
                    Exit Sub
                End If
            Next j
        Next i
    End If
End Sub

Private Sub PB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'***[Methods]********************************************************************************************************
Public Sub AddLevel(New_Level As SuperLevel)
    mvarLevels.Add New_Level
End Sub

Public Sub RemoveLevel(Index As Variant)
    mvarLevels.Remove Index
End Sub

Public Property Get Levels(Index As Variant) As SuperLevel
    Set Levels = mvarLevels(Index)
End Property

Public Property Get LevelsCount() As Long
    LevelsCount = mvarLevels.Count
End Property

'***[Visualisation]**************************************************************************************************
Private Sub DrawTree()
    Dim i As Long
    Dim j As Long
    Dim X As Long
    Dim Y As Long
    Dim myNode As SuperNode
    Dim myToolbarHeight As Long
    Dim myTmp As Long
    
    If ToolbarState = tsActive Then
        myToolbarHeight = ToolBar.Height
    Else
        myToolbarHeight = 0
    End If
    PB.Cls
    mvarMaxWidth = 0
    If LevelsCount > 0 Then
        For i = 1 To LevelsCount
            For j = 1 To Levels(i).NodesCount
                Set myNode = Levels(i).Nodes(j)
                
                X = mvarStartX + ((j - 1) * NODE_HSPACE)
                Y = mvarStartY + ((i - 1) * NODE_VSPACE)
                
                With myNode
                    Select Case mvarTreeGrowth
                    Case tgFromLeft
                        .Left = X
                        .Right = X + NODE_WIDTH + NODE_SHADOW_DEPTH
                        .Center = X + (NODE_WIDTH / 2)
                    Case tgFromSelected
                        If Not .ParentNode Is Nothing Then
                            .Left = X + .ParentNode.Left - mvarStartX
                            .Right = X + .ParentNode.Right - mvarStartX
                            .Center = X + .ParentNode.Center - mvarStartX
                        Else
                            .Left = X
                            .Right = X + NODE_WIDTH + NODE_SHADOW_DEPTH
                            .Center = X + (NODE_WIDTH / 2)
                        End If
                    End Select
                    .Top = Y
                    .Bottom = Y + NODE_HEIGHT + NODE_SHADOW_DEPTH
                End With
                
                
                If mvarMaxWidth < myNode.Right Then
                    mvarMaxWidth = myNode.Right + mvarStartX
                End If
                PB.Width = mvarMaxWidth
                PB.Height = (2 * mvarStartY) + (LevelsCount * NODE_VSPACE) '- (NODE_VSPACE - NODE_HEIGHT)
                DrawNode myNode
            
                myTmp = PB.Width - ScaleWidth + SuperViewPort.VScrollWidth
                If myTmp < 0 Then myTmp = 0
                SuperViewPort.ViewPortWidth = myTmp

                myTmp = PB.Height - ScaleHeight + SuperViewPort.HScrollHeight
                If myTmp < 0 Then myTmp = 0
                SuperViewPort.ViewPortHeight = myTmp
            Next j
        Next i
        
        
        'Draw Connector line from selected node to the root
        If Not SelectedNode Is Nothing Then
            Set myNode = SelectedNode
            Do While Not myNode.ParentNode Is Nothing
                Select Case ConnectorType
                Case ctStraightLine
                    PB.DrawWidth = 3
                    PB.Line (myNode.Center, myNode.Top - 15)-(myNode.ParentNode.Center, myNode.ParentNode.Bottom + 15), ConnectorColor
                    PB.DrawWidth = 1
                Case ctStepLine
                    PB.DrawWidth = 3
                    PB.Line (myNode.ParentNode.Center, myNode.ParentNode.Bottom)-(myNode.ParentNode.Center, myNode.ParentNode.Bottom + (myNode.Top - myNode.ParentNode.Bottom) / 2), ConnectorColor
                    PB.Line (myNode.ParentNode.Center, myNode.ParentNode.Bottom + (myNode.Top - myNode.ParentNode.Bottom) / 2)-(myNode.Center, myNode.ParentNode.Bottom + (myNode.Top - myNode.ParentNode.Bottom) / 2), ConnectorColor
                    PB.Line (myNode.Center, myNode.ParentNode.Bottom + (myNode.Top - myNode.ParentNode.Bottom) / 2)-(myNode.Center, myNode.Top), ConnectorColor
                    PB.DrawWidth = 1
                End Select
                Set myNode = myNode.ParentNode
            Loop
        End If
    End If
End Sub

'Private Sub DrawNode(Node As SuperNode, X As Long, Y As Long)
Private Sub DrawNode(Node As SuperNode)
    Dim i As Long
    
    PB.ForeColor = RGB(50, 50, 50)
    PB.Line (Node.Left + NODE_SHADOW_DEPTH, Node.Top + NODE_SHADOW_DEPTH)-(Node.Right, Node.Bottom), , BF     'Node Shadow
    If Node.Selected = True Then
        PB.ForeColor = RGB(200, 230, 200)
    Else
        PB.ForeColor = RGB(240, 230, 200)
    End If
    PB.Line (Node.Left, Node.Top)-(Node.Right - NODE_SHADOW_DEPTH, Node.Bottom - NODE_SHADOW_DEPTH), , BF 'Node Background
    PB.ForeColor = RGB(70, 70, 70)
    PB.Line (Node.Left, Node.Top)-(Node.Right - NODE_SHADOW_DEPTH, Node.Bottom - NODE_SHADOW_DEPTH), , B 'Node Border
    PB.Line (Node.Left, Node.Top + NODE_HEIGHT / 2)-(Node.Right - NODE_SHADOW_DEPTH, Node.Top + NODE_HEIGHT / 2), , B 'Horizontal Split Line
    PB.Line (Node.Left + NODE_WIDTH / 3 * 2, Node.Top + NODE_HEIGHT / 2)-(Node.Left + NODE_WIDTH / 3 * 2, Node.Top + NODE_HEIGHT), , B 'Vertical Split Line
    
    'Draw all connection lines, (path to selected node will be made later in main DrawTree procedure)
    If Not Node.ParentNode Is Nothing Then
        Select Case ConnectorType
        Case ctStraightLine
            PB.Line (Node.Center, Node.Top - 15)-(Node.ParentNode.Center, Node.ParentNode.Bottom + 15), ConnectorColor
        Case ctStepLine
            PB.Line (Node.ParentNode.Center, Node.ParentNode.Bottom)-(Node.ParentNode.Center, Node.ParentNode.Bottom + (Node.Top - Node.ParentNode.Bottom) / 2), ConnectorColor
            PB.Line (Node.ParentNode.Center, Node.ParentNode.Bottom + (Node.Top - Node.ParentNode.Bottom) / 2)-(Node.Center, Node.ParentNode.Bottom + (Node.Top - Node.ParentNode.Bottom) / 2), ConnectorColor
            PB.Line (Node.Center, Node.ParentNode.Bottom + (Node.Top - Node.ParentNode.Bottom) / 2)-(Node.Center, Node.Top), ConnectorColor
        Case ctSpray
            'This is slappy gradient code that i'll change
            Dim myStep As Long
            Dim myColorStep As Long
            Dim myColor As Long
            Dim myR As Long
            Dim myG As Long
            Dim myB As Long
            Dim myDirection As Integer
            myR = 255: myG = 255: myB = 255
            myDirection = 0
            myStep = 15
            myColorStep = NODE_WIDTH / 510
            For i = Node.Left + 60 To Node.Right - 60 Step myStep
                myColor = RGB(myR, myG, myB)
                PB.Line (i, Node.Top - 45)-(Node.ParentNode.Center, Node.ParentNode.Bottom + 15), myColor
                If myDirection = 0 Then
                    myR = myR - myColorStep
                    myG = myG - myColorStep
                    myB = myB - myColorStep
                    If myR < 0 Then
                        myDirection = 1
                        myR = 0
                        myG = 0
                        myB = 0
                    End If
                Else
                    myR = myR + myColorStep
                    myG = myG + myColorStep
                    myB = myB + myColorStep
                End If
            Next i
        End Select
    End If
    
    'Draw String values
    DrawText Node.Caption, Node.Left + 60, Node.Top + 30, RGB(0, 0, 0), True
    DrawText Node.Value, Node.Left + 60, Node.Top + NODE_HEIGHT / 2 + 30, RGB(0, 0, 0)
    DrawText Node.Percentage & " %", Node.Left + NODE_WIDTH / 3 * 2 + 60, Node.Top + NODE_HEIGHT / 2 + 30, RGB(255, 0, 0)
End Sub

Private Sub DrawText(Text As String, X As Long, Y As Long, Color As Long, Optional Bold As Boolean = False)
    PB.CurrentX = X
    PB.CurrentY = Y
    PB.ForeColor = Color
    PB.FontBold = Bold
    PB.Print Text
End Sub


