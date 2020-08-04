VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Konfirmasi Keluar Program"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Jawab As Integer
   Jawab = MsgBox("Anda yakin akan keluar dari program?", vbQuestion + vbYesNo, "Konfirmasi Keluar")
   If Jawab = vbNo Then Cancel = -1
End Sub

