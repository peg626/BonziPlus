VERSION 5.00
Begin VB.Form frmName 
   Caption         =   "Name"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "Name.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Qual é seu nome?"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If Not Text1.Text = "" Then
        Main.BonziAni "Greet"
        Main.BonziFalar ("Prazer em te conhecer," & Text1.Text & ".")
        Main.SaveName Text1.Text
        Main.SaveFirstStart False
        Me.Hide
    Else
        MsgBox "Error:Esta vazio o nome.", vbCritical
    End If
End Sub

