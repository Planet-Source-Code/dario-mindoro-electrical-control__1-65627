VERSION 5.00
Begin VB.UserControl LED 
   BackColor       =   &H00939393&
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2370
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
   ScaleHeight     =   765
   ScaleWidth      =   2370
   ToolboxBitmap   =   "LED.ctx":0000
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   225
      Left            =   465
      TabIndex        =   0
      Top             =   90
      Width           =   1785
   End
   Begin VB.Image imgOff 
      Height          =   390
      Left            =   0
      Picture         =   "LED.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   390
   End
   Begin VB.Image imgON 
      Height          =   390
      Left            =   0
      Picture         =   "LED.ctx":0AC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   390
   End
End
Attribute VB_Name = "LED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Led Control : Dario Mindoro - www.mindworksoft.com


Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private mpoiCursorPos As POINTAPI

Private LedStat As Boolean

Public Event MouseOver()

Private Sub UserControl_InitProperties()
    Status = False
    Caption = Ambient.DisplayName
End Sub

Private Sub UserControl_Resize()
    lblcaption.Left = imgOff.Width + 100
End Sub

Public Property Get Status() As Boolean
    Status = LedStat
End Property

Public Property Let Status(ByVal sNewValue As Boolean)
    LedStat = sNewValue
    UserControl.PropertyChanged "Status"
    If LedStat = True Then
        imgON.Visible = True
        imgOff.Visible = False
    Else
        imgON.Visible = False
        imgOff.Visible = True
    End If
End Property


Public Property Get Caption() As String
Caption = lblcaption.Caption
End Property

Public Property Let Caption(ByVal sNewValue As String)
lblcaption.Caption = sNewValue
UserControl.PropertyChanged "Caption"
End Property

Public Property Get ForeColor() As Long
ForeColor = lblcaption.ForeColor
End Property

Public Property Let ForeColor(ByVal lNewValue As Long)
lblcaption.ForeColor = lNewValue
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get Bold() As Boolean
Bold = lblcaption.FontBold
End Property

Public Property Let Bold(ByVal bNewValue As Boolean)
lblcaption.FontBold = bNewValue
UserControl.PropertyChanged "Bold"
End Property



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Status", LedStat, False
    .WriteProperty "Caption", Caption, Ambient.DisplayName
    .WriteProperty "ForeColor", ForeColor, 0
    .WriteProperty "Bold", Bold, False
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    Status = .ReadProperty("Status", False)
    Caption = .ReadProperty("Caption", Ambient.DisplayName)
    ForeColor = .ReadProperty("ForeColor", 0)
    Bold = .ReadProperty("Bold", False)
End With
End Sub
