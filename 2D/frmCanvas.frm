VERSION 5.00
Begin VB.Form frmCanvas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "MIDAR's Matrix Lessons using Asteroids - http://www.midar.com/vblessons/"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   Icon            =   "frmCanvas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_DoAnimation 
      Interval        =   25
      Left            =   105
      Top             =   60
   End
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_DblClick()

    If Me.WindowState <> vbNormal Then Exit Sub
    Me.Width = Me.Height
    
End Sub

Private Sub Form_Load()
    
    Call Form_DblClick
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Me.Timer_DoAnimation.Enabled = False
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Me.Timer_DoAnimation.Enabled = True

End Sub

Private Sub Form_Resize()

    Call Init_ViewMapping
    
End Sub

Private Sub Timer_DoAnimation_Timer()
    
    Call Main
    
End Sub
