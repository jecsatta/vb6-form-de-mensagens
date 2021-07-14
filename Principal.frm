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
   Begin VB.CommandButton Botoes 
      Caption         =   "Abrir"
      Height          =   795
      Left            =   2160
      TabIndex        =   0
      Top             =   960
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
        .AdicionaBotao "Teste"
        .AdicionaBotao "Teste2"
        a = .Mostra
    End With
    MsgBox a
End Sub

