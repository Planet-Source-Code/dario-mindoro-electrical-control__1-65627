VERSION 5.00
Begin VB.UserControl ROOM 
   BackColor       =   &H00404040&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2355
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MouseIcon       =   "ROOM.ctx":0000
   MousePointer    =   99  'Custom
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   157
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ROOM"
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   945
      TabIndex        =   0
      Top             =   975
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   375
      Top             =   315
      Width           =   1575
   End
End
Attribute VB_Name = "ROOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Room control : Dario Mindoro - www.mindworksoft.com

Public Enum RoomStatus
    LoadOFF = 0
    LoadON = 1
End Enum


Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ButtonUp(Button As Integer, Shift As Integer)

Public Event DblClick()


Dim mMouseDown As Boolean
Dim mRoomID As Integer
Dim mStat As RoomStatus
Dim tmpcolor As Variant

Public Property Get Color() As OLE_COLOR
  Color = Shape1.FillColor
End Property

Public Property Let Color(ByVal c As OLE_COLOR)
  Shape1.FillColor = c
End Property

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
  tmpcolor = Shape1.BorderColor
  Shape1.BorderColor = &HFFFFFF
End Sub

Private Sub UserControl_ExitFocus()
  Shape1.BorderColor = tmpcolor
End Sub

Private Sub UserControl_InitProperties()
    Caption = Ambient.DisplayName
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button = 1 Then
        mMouseDown = True
    End If
    
    If Button <> 1 And mMouseDown = True Then
        mMouseDown = False
        RaiseEvent ButtonUp(Button, Shift)
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent ButtonUp(Button, Shift)
End Sub

Private Sub UserControl_Resize()
  d = Shape1.BorderWidth / 2
  Shape1.Left = d
  Shape1.Top = d
  Shape1.Width = ScaleWidth - d * 2
  Shape1.Height = ScaleHeight - d * 2
  
  lblcaption.Left = 5
  lblcaption.Top = 5

End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Caption", Caption, Ambient.DisplayName
    .WriteProperty "RoomID", Caption, mRoomID
    .WriteProperty "RoomStatus", mStat, 0
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    Caption = .ReadProperty("Caption", Ambient.DisplayName)
    mRoomID = .ReadProperty(RoomID, 0)
    mStat = .ReadProperty(RoomStatus, 0)
End With
End Sub
Public Property Get Caption() As String
Caption = lblcaption.Caption
End Property

Public Property Let Caption(ByVal sNewValue As String)
    lblcaption.Caption = sNewValue
    UserControl.PropertyChanged "Caption"
End Property
Public Property Get RoomID() As Integer
    RoomID = mRoomID
End Property
Public Property Let RoomID(ByVal sNewValue As Integer)
    mRoomID = sNewValue
    UserControl.PropertyChanged "RoomID"
End Property

Public Property Get RoomStatus() As RoomStatus
    RoomStatus = mStat
End Property
Public Property Let RoomStatus(ByVal sNewValue As RoomStatus)
    mStat = sNewValue
    If mStat = LoadOFF Then
        Shape1.BorderColor = &HC00000
        Shape1.FillColor = &H800000
    Else
        Shape1.BorderColor = &HC0C0&
        Shape1.FillColor = &H8080&
    End If
    
    UserControl.PropertyChanged "RoomStatus"
    
End Property

