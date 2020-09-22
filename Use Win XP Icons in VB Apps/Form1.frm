VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Use Win XP Icons in VB Apps"
   ClientHeight    =   300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call mod32BitIcon.SetIcon(Me.hWnd, "AAA")
End Sub
