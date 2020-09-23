VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl SuperViewPort 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ControlContainer=   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   4770
   Begin VB.PictureBox PB 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4500
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   2220
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.FlatScrollBar VScroll 
      Height          =   2205
      Left            =   4500
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3889
      _Version        =   393216
      Orientation     =   1179648
   End
   Begin MSComCtl2.FlatScrollBar HScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2190
      Visible         =   0   'False
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1179649
   End
End
Attribute VB_Name = "SuperViewPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***[Enumerations]***************************************************************************************************
Public Enum svpAppearance
    apFlat = 0
    ap3D = 1
End Enum

Public Enum svpBorderStyle
    bsNoBorder = 0
    bsFixedSingle = 1
End Enum

Const mvar_def_ViewPortHeight As Long = 6400
Const mvar_def_ViewPortWidth As Long = 9600
Const mvar_def_SmallChangeH As Long = 10
Const mvar_def_SmallChangeV As Long = 10
Const mvar_def_LargeChangeH As Long = 100
Const mvar_def_LargeChangeV As Long = 100

Dim mvarViewPortHeight As Long
Dim mvarViewPortWidth As Long
Dim mvarSmallChangeH As Long
Dim mvarSmallChangeV As Long
Dim mvarLargeChangeH As Long
Dim mvarLargeChangeV As Long

Dim mvarVScrollOldVal As Long
Dim mvarHScrollOldVal As Long

'***[Properties]*****************************************************************************************************
Public Property Get Appearance() As svpAppearance
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal Value As svpAppearance)
    UserControl.Appearance() = Value
    PropertyChanged "Appearance"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    UserControl.BackColor() = Value
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As svpBorderStyle
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As svpBorderStyle)
    UserControl.BorderStyle() = Value
    PropertyChanged "BorderStyle"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    UserControl.Enabled() = Value
    PropertyChanged "Enabled"
End Property

Public Property Let ViewPortHeight(Value As Long)
    mvarViewPortHeight = Value
    VScroll.Max = Value
    PropertyChanged "ViewPortHeight"
End Property

Public Property Get ViewPortHeight() As Long
    ViewPortHeight = mvarViewPortHeight
End Property

Public Property Let ViewPortWidth(Value As Long)
    mvarViewPortWidth = Value
    HScroll.Max = Value
    PropertyChanged "ViewPortWidth"
End Property

Public Property Get ViewPortWidth() As Long
    ViewPortWidth = mvarViewPortWidth
End Property

Public Property Let SmallChangeH(Value As Long)
    mvarSmallChangeH = Value
    HScroll.SmallChange = Value
    PropertyChanged "SmallChangeH"
End Property

Public Property Get SmallChangeH() As Long
    SmallChangeH = mvarSmallChangeH
End Property

Public Property Let SmallChangeV(Value As Long)
    mvarSmallChangeV = Value
    VScroll.SmallChange = Value
    PropertyChanged "SmallChangeV"
End Property

Public Property Get SmallChangeV() As Long
    SmallChangeV = mvarSmallChangeV
End Property

Public Property Let LargeChangeH(Value As Long)
    mvarLargeChangeH = Value
    HScroll.LargeChange = LargeChangeH
    PropertyChanged "LargeChangeH"
End Property

Public Property Get LargeChangeH() As Long
    LargeChangeH = mvarLargeChangeH
End Property

Public Property Let LargeChangeV(Value As Long)
    mvarLargeChangeV = Value
    VScroll.LargeChange = Value
    PropertyChanged "LargeChangeV"
End Property

Public Property Get LargeChangeV() As Long
    LargeChangeV = mvarLargeChangeV
End Property


Public Property Get VScrollWidth() As Long
    If VScroll.Visible = True Then
        VScrollWidth = VScroll.Width
    Else
        VScrollWidth = 0
    End If
End Property

Public Property Get HScrollHeight() As Long
    If HScroll.Visible = True Then
        HScrollHeight = HScroll.Height
    Else
        HScrollHeight = 0
    End If
End Property

'***[Life Control]***************************************************************************************************
'Initialize with default values
Private Sub UserControl_InitProperties()
    mvarViewPortHeight = mvar_def_ViewPortHeight
    mvarViewPortWidth = mvar_def_ViewPortWidth
    mvarSmallChangeH = mvar_def_SmallChangeH
    mvarSmallChangeV = mvar_def_SmallChangeV
    mvarLargeChangeH = mvar_def_LargeChangeH
    mvarLargeChangeV = mvar_def_LargeChangeV
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", bsFixedSingle)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", ap3D)
    ViewPortHeight = PropBag.ReadProperty("ViewPortHeight", mvar_def_ViewPortHeight)
    ViewPortWidth = PropBag.ReadProperty("ViewPortWidth", mvar_def_ViewPortWidth)
    SmallChangeH = PropBag.ReadProperty("SmallChangeH", mvar_def_SmallChangeH)
    SmallChangeV = PropBag.ReadProperty("SmallChangeV", mvar_def_SmallChangeV)
    LargeChangeH = PropBag.ReadProperty("LargeChangeH", mvar_def_LargeChangeH)
    LargeChangeV = PropBag.ReadProperty("LargeChangeV", mvar_def_LargeChangeV)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, bsFixedSingle)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, ap3D)
    Call PropBag.WriteProperty("ViewPortHeight", ViewPortHeight, mvar_def_ViewPortHeight)
    Call PropBag.WriteProperty("ViewPortWidth", ViewPortWidth, mvar_def_ViewPortWidth)
    Call PropBag.WriteProperty("SmallChangeH", SmallChangeH, mvar_def_SmallChangeH)
    Call PropBag.WriteProperty("SmallChangeV", SmallChangeV, mvar_def_SmallChangeV)
    Call PropBag.WriteProperty("LargeChangeH", LargeChangeH, mvar_def_LargeChangeH)
    Call PropBag.WriteProperty("LargeChangeV", LargeChangeV, mvar_def_LargeChangeV)
End Sub

Private Sub UserControl_Initialize()
    mvarVScrollOldVal = 0
    mvarHScrollOldVal = 0
End Sub


Private Sub UserControl_Resize()
    VScroll.Value = 0
    HScroll.Value = 0
    RenderViewPort
End Sub

Private Sub UserControl_Paint()
    RenderViewPort
End Sub


Private Sub RenderViewPort()
    If Ambient.UserMode = True Then
        
        VScroll.Visible = True
        HScroll.Visible = True
        
        
        HScroll.Move 0, ScaleHeight - HScroll.Height, ScaleWidth - VScrollWidth
        If ScaleHeight > HScrollHeight Then
            VScroll.Move ScaleWidth - VScroll.Width, 0, VScroll.Width, ScaleHeight - HScrollHeight
        End If
        
        'Visibility and position of picture box in the lower-right corner
        PB.Visible = True
        PB.Move HScroll.Width, VScroll.Height, VScroll.Width, HScroll.Height
    
        PB.ZOrder vbBringToFront
        VScroll.ZOrder vbBringToFront
        HScroll.ZOrder vbBringToFront
    End If
End Sub

Public Sub Refresh()
    RenderViewPort
End Sub

Private Sub VScroll_Change()
    Dim myCtl As Control
    For Each myCtl In UserControl.ContainedControls
        If mvarVScrollOldVal <> VScroll.Value Then
            myCtl.Top = myCtl.Top + (mvarVScrollOldVal - VScroll.Value)
        End If
    Next
    mvarVScrollOldVal = VScroll.Value
End Sub

Private Sub VScroll_Scroll()
    Dim myCtl As Control
    For Each myCtl In UserControl.ContainedControls
        If mvarVScrollOldVal <> VScroll.Value Then
            myCtl.Top = myCtl.Top + (mvarVScrollOldVal - VScroll.Value)
        End If
    Next
    mvarVScrollOldVal = VScroll.Value
End Sub

Private Sub HScroll_Change()
    Dim myCtl As Control
    For Each myCtl In UserControl.ContainedControls
        If mvarHScrollOldVal <> HScroll.Value Then
            myCtl.Left = myCtl.Left + (mvarHScrollOldVal - HScroll.Value)
        End If
    Next
    mvarHScrollOldVal = HScroll.Value
End Sub

Private Sub HScroll_Scroll()
    Dim myCtl As Control
    For Each myCtl In UserControl.ContainedControls
        If mvarHScrollOldVal <> HScroll.Value Then
            myCtl.Left = myCtl.Left + (mvarHScrollOldVal - HScroll.Value)
        End If
    Next
    mvarHScrollOldVal = HScroll.Value
End Sub


