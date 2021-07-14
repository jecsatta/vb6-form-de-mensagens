VERSION 5.00
Begin VB.Form Principal 
   Caption         =   "Principal"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Abrir"
      Height          =   795
      Left            =   2400
      TabIndex        =   1
      Top             =   1440
      Width           =   1492
   End
   Begin VB.CommandButton Botoes 
      Caption         =   "Abrir"
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   1500
      Width           =   1492
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Botoes_Click()
    Dim a As Integer
    With New FormMensagem
        .Titulo "Titulo"
        .Altura 3000 'optional
        .Largura 5000 'optional
        .Mensagem "Escolha uma Opcao"
        .AdicionaBotao "Teste"
        .AdicionaBotao "Teste2"
        a = .Mostra
    End With
    MsgBox a
End Sub

Private Sub Command1_Click()
    Dim a As Integer
    
    Dim s As FormMensagem
    
    Set s = New FormMensagem
    a = s.Titulo("Titulo") _
        .Mensagem("Escolha uma Opcao") _
        .AdicionaBotao("Teste", 500) _
        .AdicionaBotao("Teste2", 3000) _
        .Mostra()
    Set s = Nothing
    MsgBox a
End Sub
