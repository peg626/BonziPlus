VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Search"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Text"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim Pesquisa As String
    Pesquisa = Text1.Text
    Main.BonziStop
    Set Sea2Req = Main.RBonziAni("SearchingReturn")
    
    Set SeaReq = Main.RBonziFalar("Eu ja pesquisei para voce!")
    
    
End Sub


Private Sub Form_Load()
    Main.BonziFalar "Eu vou pesquisar alguma coisa para voce!"
    Main.BonziAni "Searching"
End Sub

Private Sub Form_Terminate()
    Main.BonziStop
    Main.BonziAni "SearchingReturn"
    Main.BonziFalar "Ok, acho que agora ja deu"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Main.BonziStop
    Main.BonziAni "SearchingReturn"
    Main.BonziFalar "Ok, acho que agora ja deu"
End Sub
