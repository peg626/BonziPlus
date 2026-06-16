VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Begin VB.Form Main 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   3000
      Top             =   2160
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RanReq As Object
Dim NameReq As Object
Private Sub Agent1_RequestComplete(ByVal Request As Object)


    If Request Is RanReq Then
        Agent1.Characters("Bonzi").Stop RanReq
    ElseIf Request Is NameReq Then
        frmName.Show
    ElseIf Request Is SeaReq Then
        Shell "cmd /c start https://www.google.com/search?q=" & Pesquisa, vbHide
        frmSearch.Hide
    
    End If
End Sub

Private Sub Form_Load()
DataPasta = Environ$("APPDATA") & "\BonziPlus"
If Dir(DataPasta, vbDirectory) = "" Then
    MkDir DataPasta
End If
Agent1.Characters.Load "Bonzi", "Bonzi.acs"
Agent1.Characters("Bonzi").Show
' Começamos com os comandos!
With Agent1.Characters("Bonzi")
.Show

.Commands.RemoveAll

.Commands.Add "agntwelc", "Bem-vindo!", "O Bonzi vai comprimentar."
.Commands.Add "unload", "Unload", "Desligar"
.Commands.Add "speak", "Falar", "Falar alguma coisa"
.Commands.Add "random", "Uma palavra sabia aleatoria"
.Commands.Add "search", "Pesquisar"
End With

Main.Hide
If LoadFirstStart() Then
    
    
    Agent1.Characters("Bonzi").Play "Greet"
    BonziFalar "Welcome to BonziPlus!"
    Agent1.Characters("Bonzi").Play "Explain"
    BonziFalar "Eu sou um remake do BonziBuddy, mas sem virus!"
    BonziFalar "Nesse momento, eu sou muito limitado ainda."
    Set NameReq = Agent1.Characters("Bonzi").Speak("Eu quero te conhecer melhor... qual é seu nome?")
    
Else
    BonziFalar "Ola," & LoadName() & "!"
End If
End Sub
Public Sub BonziAni(ByVal Animationa As String)
    Agent1.Characters("Bonzi").Play Animationa
End Sub
Public Sub SaveName(ByVal Nome As String)

    Dim Arq As Integer

    Arq = FreeFile

    Open DataPasta & "\name.txt" For Output As #Arq

    Print #Arq, Nome

    Close #Arq

End Sub
Public Sub BonziStop()
    Agent1.Characters("Bonzi").Stop
End Sub
Public Function LoadName() As String

    Dim Arq As Integer

    If Dir(DataPasta & "\name.txt") = "" Then
        LoadName = ""
        Exit Function
    End If

    Arq = FreeFile

    Open DataPasta & "\name.txt" For Input As #Arq

    If Not EOF(Arq) Then
        Line Input #Arq, LoadName
    End If

    Close #Arq

End Function
Public Function LoadFirstStart() As Boolean

    Dim Arq As Integer
    Dim Valor As String

    If Dir(DataPasta & "\firststart.txt") = "" Then

        SaveFirstStart True
        LoadFirstStart = True
        Exit Function

    End If

    Arq = FreeFile

    Open DataPasta & "\firststart.txt" For Input As #Arq

    If EOF(Arq) Then

        Close #Arq

        SaveFirstStart True
        LoadFirstStart = True
        Exit Function

    End If

    Line Input #Arq, Valor

    Close #Arq

    LoadFirstStart = (LCase$(Trim$(Valor)) = "true")

End Function
Public Sub SaveFirstStart(ByVal Valor As Boolean)

    Dim Arq As Integer

    Arq = FreeFile

    Open DataPasta & "\firststart.txt" For Output As #Arq

    If Valor Then
        Print #Arq, "true"
    Else
        Print #Arq, "false"
    End If

    Close #Arq

End Sub
Public Function RBonziAni(Animationa As String) As Object
    Set RBonziAni = Agent1.Characters("Bonzi").Play(Animationa)

End Function
Public Function RBonziFalar(Texto As String) As Object
    Set RBonziFalar = Agent1.Characters("Bonzi").Speak(Texto)
End Function
Public Sub BonziFalar(ByVal Texto As String)
    Agent1.Characters("Bonzi").Speak Texto
End Sub

Private Sub Agent1_Command(ByVal UserInput As Object)
    Select Case UserInput.Name

        Case "agntwelc"
            Dim Req As Object
            Agent1.Characters("Bonzi").Play "Greet"
            Set Req = Agent1.Characters("Bonzi").Speak("Ola," & LoadName() & "!")
            
            
        
        Case "unload"
            Agent1.Characters.Unload "Bonzi"
            Unload Me
        
        Case "speak"
            Speak.Show
        Case "random"
            Agent1.Characters("Bonzi").Play "Explain"
            FalarAleatorio
        Case "search"
            frmSearch.Show
    End Select
End Sub

Private Sub FalarAleatorio()
    
    
    Dim Frases(4) As String

    Frases(0) = "Um sabio é sempre sabio."
    Frases(1) = "1+1=10"
    Frases(2) = "6 7"
    Frases(3) = "4 2"
    Frases(4) = "Penso, logo existo"

    Randomize
    Set RanReq = Agent1.Characters("Bonzi").Speak(Frases(Int(Rnd * 5)))
    
    
End Sub

