VERSION 5.00
Begin VB.Form Speak 
   Caption         =   "Speak"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "Speak.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Speak"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
End
Attribute VB_Name = "Speak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Main.BonziFalar Text1.Text
    Text1.Text = ""
    Unload Speak
End Sub

