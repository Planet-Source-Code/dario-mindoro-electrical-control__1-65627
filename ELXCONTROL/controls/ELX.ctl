VERSION 5.00
Begin VB.UserControl ELX 
   BackColor       =   &H00939393&
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LockControls    =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   5940
   ToolboxBitmap   =   "ELX.ctx":0000
   Begin VB.PictureBox fanON 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2100
      Picture         =   "ELX.ctx":0312
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   5
      Top             =   2145
      Width           =   1050
   End
   Begin VB.PictureBox fanOFF 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2115
      Picture         =   "ELX.ctx":3A04
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   4
      Top             =   1380
      Width           =   1050
   End
   Begin VB.PictureBox lightON 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   3315
      Picture         =   "ELX.ctx":70F6
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   3
      Top             =   2130
      Width           =   1050
   End
   Begin VB.PictureBox lightOFF 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   3345
      Picture         =   "ELX.ctx":A7E8
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   2
      Top             =   1350
      Width           =   1050
   End
   Begin VB.PictureBox imgOff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      Picture         =   "ELX.ctx":DEDA
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   1
      Top             =   270
      Width           =   1050
   End
   Begin VB.PictureBox imgON 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      Picture         =   "ELX.ctx":115CC
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   0
      Top             =   0
      Width           =   1050
   End
   Begin VB.Timer tmrAnim 
      Enabled         =   0   'False
      Left            =   390
      Top             =   1950
   End
   Begin VB.Image imgAnim 
      Height          =   750
      Index           =   2
      Left            =   300
      Picture         =   "ELX.ctx":14CBE
      Top             =   4080
      Width           =   1050
   End
   Begin VB.Image imgAnim 
      Height          =   750
      Index           =   1
      Left            =   300
      Picture         =   "ELX.ctx":183B0
      Top             =   3285
      Width           =   1050
   End
   Begin VB.Image imgAnim 
      Height          =   750
      Index           =   0
      Left            =   300
      Picture         =   "ELX.ctx":1BAA2
      Top             =   2520
      Width           =   1050
   End
End
Attribute VB_Name = "ELX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Electronic Switch : Dario Mindoro - www.mindworksoft.com

Option Explicit


Private Type POINTAPI
        X As Long
        Y As Long
End Type

Public Enum ItemType
    FAN = 0
    LIGHT = 1
End Enum

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private mpoiCursorPos As POINTAPI

Public Event Click()
Public Event DblClick()

Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private LedStat As Boolean
Private ctrlID As Integer
Private animCount As Integer
Private ctrlNum As Integer

Dim m_ItemType As ItemType



Private Sub imgOff_Click()
    RaiseEvent Click
End Sub

Private Sub imgOff_DblClick()
    LedStat = True
    Status = LedStat
    UserControl.PropertyChanged "Status"
    RaiseEvent DblClick
End Sub

Private Sub imgOff_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgOff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub imgON_Click()
    RaiseEvent Click
End Sub

Private Sub imgON_DblClick()
    LedStat = False
    Status = LedStat
    UserControl.PropertyChanged "Status"
    RaiseEvent DblClick
End Sub


Private Sub imgOff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgON_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgON_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgON_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub tmrAnim_Timer()
    If m_ItemType <> FAN Then
        tmrAnim.Enabled = False
        Exit Sub
    End If
    
    animCount = animCount + 1
    If animCount > imgAnim.UBound Then
        animCount = 0
    End If
    imgON.Cls
    imgON.Picture = imgAnim(animCount).Picture
End Sub

Private Sub UserControl_Initialize()
    imgON.Top = 0
    imgON.Left = 0
    imgOff.Top = 0
    imgOff.Left = 0
End Sub

Private Sub UserControl_InitProperties()
    Status = False
    controltype = FAN
End Sub


Public Property Get controltype() As ItemType
    controltype = m_ItemType
End Property
Public Property Let controltype(ByVal sNewValue As ItemType)
    m_ItemType = sNewValue
    UserControl.PropertyChanged "ControlType"
    
    Select Case m_ItemType
    Case Is = 0
        
        imgOff.Picture = fanOFF.Picture
        imgOff.Width = fanOFF.Width
        imgOff.Height = fanOFF.Height
        
        imgON.Picture = fanON.Picture
        imgON.Width = fanON.Width
        imgON.Height = fanON.Height
        
        If LedStat = True Then
            tmrAnim.Interval = 50
            tmrAnim.Enabled = True
            animCount = 0
        Else
            'turn off the animation for fan
            tmrAnim.Interval = 0
            tmrAnim.Enabled = False
        End If
        
    Case Is = 1
    
        imgOff.Picture = lightOFF.Picture
        imgOff.Width = lightOFF.Width
        imgOff.Height = lightOFF.Height
        
        imgON.Picture = lightON.Picture
        imgON.Width = lightON.Width
        imgON.Height = lightON.Height
    End Select
    
    Call UserControl_Resize
End Property


Public Property Get ControlID() As Integer
    ControlID = ctrlID
End Property
Public Property Let ControlID(ByVal sNewValue As Integer)
    ctrlID = sNewValue
    UserControl.PropertyChanged "ControlID"
End Property


Public Property Get Status() As Boolean
    Status = LedStat
End Property

Public Property Let Status(ByVal sNewValue As Boolean)
    LedStat = sNewValue
    UserControl.PropertyChanged "Status"
    If LedStat = True Then
        imgON.Visible = True
        imgOff.Visible = False
        
        If m_ItemType = FAN Then
            tmrAnim.Interval = 20
            tmrAnim.Enabled = True
            animCount = 0
        End If
        
    Else
        imgON.Visible = False
        imgOff.Visible = True
        
        If m_ItemType = FAN Then
            tmrAnim.Interval = 0
            tmrAnim.Enabled = False
        End If
        
    End If
End Property


Private Sub UserControl_Resize()
    UserControl.Width = imgON.Width
    UserControl.Height = imgON.Height
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Status", LedStat, False
    .WriteProperty "ControlID", ctrlID, 0
    .WriteProperty "ControlNumber", ctrlNum, 0
    .WriteProperty "ControlType", m_ItemType, 0
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    Status = .ReadProperty("Status", False)
    ControlID = .ReadProperty("ControlID", 0)
    ctrlNum = .ReadProperty("ControlNumber", 0)
    m_ItemType = .ReadProperty("ControlType", 0)
End With
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

