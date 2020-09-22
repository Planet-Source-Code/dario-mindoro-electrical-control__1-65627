VERSION 5.00
Begin VB.UserControl box3d2 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "box3d2.ctx":0000
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   24
      X2              =   290
      Y1              =   226
      Y2              =   226
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   306
      X2              =   306
      Y1              =   28
      Y2              =   224
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   30
      X2              =   304
      Y1              =   14
      Y2              =   14
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   14
      X2              =   14
      Y1              =   18
      Y2              =   220
   End
End
Attribute VB_Name = "box3d2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'3d frame : Dario Mindoro - www.mindworksoft.com


Private Sub UserControl_Initialize()
UserControl_Resize
End Sub

Private Sub UserControl_Resize()
  Line1.X1 = 0
  Line1.X2 = 0
  Line1.Y1 = 0
  Line1.Y2 = UserControl.ScaleHeight
  
  Line2.X1 = 0
  Line2.X2 = UserControl.ScaleWidth
  Line2.Y1 = 0
  Line2.Y2 = 0
  
  Line3.X1 = UserControl.ScaleWidth - 1
  Line3.X2 = UserControl.ScaleWidth - 1
  Line3.Y1 = 0
  Line3.Y2 = UserControl.ScaleHeight
  
  Line4.X1 = 0
  Line4.X2 = UserControl.ScaleWidth
  Line4.Y1 = UserControl.ScaleHeight - 1
  Line4.Y2 = UserControl.ScaleHeight - 1
 
End Sub
