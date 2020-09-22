VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dards Electrical controls"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   555
      Left            =   90
      TabIndex        =   13
      Top             =   6600
      Width           =   1665
   End
   Begin Project1.box3d2 box3d23 
      Height          =   1035
      Left            =   90
      TabIndex        =   12
      Top             =   4440
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1826
      Begin VB.Label Label3 
         Caption         =   "Double click the Fan or the lamp to change its status ON/OFF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   825
         Left            =   180
         TabIndex        =   14
         Top             =   120
         Width           =   1275
      End
   End
   Begin Project1.box3d2 box3d22 
      Height          =   2115
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   3731
      Begin Project1.ROOM ROOM1 
         Height          =   1185
         Left            =   240
         TabIndex        =   9
         Top             =   660
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   2090
         RoomID          =   "ROOM1"
      End
      Begin VB.Label Label1 
         Caption         =   "Selected Room"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   270
         TabIndex        =   10
         Top             =   330
         Width           =   1155
      End
   End
   Begin Project1.box3d2 box3d21 
      Height          =   2085
      Left            =   90
      TabIndex        =   1
      Top             =   2280
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   3678
      Begin Project1.LED LED1 
         Height          =   435
         Left            =   210
         TabIndex        =   2
         Top             =   570
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   767
         Caption         =   "FAN"
         ForeColor       =   -2147483630
      End
      Begin Project1.LED LED2 
         Height          =   435
         Left            =   210
         TabIndex        =   3
         Top             =   990
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   767
         Caption         =   "LAMP1"
         ForeColor       =   -2147483630
      End
      Begin Project1.LED LED3 
         Height          =   435
         Left            =   210
         TabIndex        =   4
         Top             =   1410
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   767
         Caption         =   "LAMP2"
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label2 
         Caption         =   "Circuit Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   270
         TabIndex        =   11
         Top             =   150
         Width           =   1095
      End
   End
   Begin VB.PictureBox picMap 
      BackColor       =   &H00404040&
      Height          =   7080
      Left            =   1860
      Picture         =   "Form1.frx":000C
      ScaleHeight     =   468
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   538
      TabIndex        =   0
      Top             =   90
      Width           =   8130
      Begin Project1.box3d2 box3d24 
         Height          =   435
         Left            =   60
         TabIndex        =   15
         Top             =   6510
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   767
         Begin VB.Label Label4 
            Caption         =   "I made this controls for my BMS project but I want to share this to PSC just vote it if its useful to you"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   7605
         End
      End
      Begin Project1.ELX ELX2 
         Height          =   750
         Left            =   3150
         TabIndex        =   5
         Top             =   630
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1323
      End
      Begin Project1.ELX ELX1 
         Height          =   750
         Left            =   240
         TabIndex        =   6
         Top             =   2460
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1323
         ControlType     =   1
      End
      Begin Project1.ELX ELX3 
         Height          =   750
         Left            =   6630
         TabIndex        =   7
         Top             =   1950
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1323
         ControlType     =   1
      End
   End
   Begin Project1.box3d2 box3d25 
      Height          =   975
      Left            =   90
      TabIndex        =   17
      Top             =   5550
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1720
      Begin VB.Label Label5 
         Caption         =   "By: Dario Mindoro MindWorkSoft.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   465
         Left            =   180
         TabIndex        =   18
         Top             =   270
         Width           =   1365
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub ELX1_DblClick()
    'turn on/off the control
    LED2.Status = ELX1.Status
    CheckRoomStatus
End Sub


Private Sub ELX2_DblClick()
    'turn on/off the control
    LED1.Status = ELX2.Status
    CheckRoomStatus
End Sub


Private Sub ELX3_DblClick()
    'turn on/off the control
    LED3.Status = ELX3.Status
    CheckRoomStatus
End Sub

Private Sub Form_Load()
    'set all control type
    ELX1.controltype = LIGHT
    ELX3.controltype = LIGHT
    ELX2.controltype = FAN
End Sub

Private Sub CheckRoomStatus()
    'check if there are any lamp or fans turned off or On
    If ELX1.Status = False And ELX2.Status = False And ELX3.Status = False Then
        ROOM1.RoomStatus = LoadOFF
    Else
        ROOM1.RoomStatus = LoadON
    End If
End Sub
