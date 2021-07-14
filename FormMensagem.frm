VERSION 5.00
Begin VB.Form FormMensagem 
   BorderStyle     =   0  'None
   Caption         =   "Mensagem"
   ClientHeight    =   3015
   ClientLeft      =   2715
   ClientTop       =   3360
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   0
      ScaleHeight     =   2985
      ScaleMode       =   0  'User
      ScaleWidth      =   5533.746
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C88F71&
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   0
         ScaleHeight     =   420
         ScaleWidth      =   5640
         TabIndex        =   1
         Top             =   0
         Width           =   5640
         Begin VB.Label labelTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Teste"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   0
            TabIndex        =   2
            Top             =   60
            Width           =   5625
         End
      End
      Begin VB.CommandButton Botoes 
         Caption         =   "Esse botão não é exibido em tela"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   2520
         Visible         =   0   'False
         Width           =   3765
      End
      Begin VB.Label labelMensagem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem"
         Height          =   735
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   5415
      End
   End
End
Attribute VB_Name = "FormMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vOpcao As Integer
Private vTopPadrao As Integer
Const C_Espacamento = 10

Public Property Get Opcao() As Integer
    Opcao = vOpcao
End Property

Public Function AdicionaBotao(vCaption As String, Optional vLeft As Integer = -1, Optional vTop As Integer = -1, Optional vWidth As Integer = 1500, Optional vHeight = 375) As FormMensagem
    Dim Botao           As CommandButton
    Dim vIndex          As Integer
    Dim vTamanhoTotal   As Integer
    Dim i               As Integer
    
    Set AdicionaBotao = Me
    vIndex = Botoes.Count
    Load Botoes(vIndex)
    
    Botoes(vIndex).Width = vWidth
    
    If vTop = -1 Then
        vTop = vTopPadrao
    Else
        vTopPadrao = vTop
    End If
    
    vTamanhoTotal = 0
    
    For i = 1 To Botoes.Count - 2
        vTamanhoTotal = vTamanhoTotal + Botoes(i).Width
    Next

    
    'Left inicial + todos os botoes + espaçamento
    If vLeft = -1 Then vLeft = 100 + vTamanhoTotal + vIndex * C_Espacamento

    Botoes(vIndex).Move vLeft, vTop, vWidth, 375
    Botoes(vIndex).Caption = vCaption
    Botoes(vIndex).Visible = True
End Function

Public Function Altura(vAltura As Integer) As FormMensagem
Attribute Altura.VB_UserMemId = 1610809349
    Set Altura = Me
    Me.Height = vAltura
    Picture1.Height = Me.Height
End Function

Private Sub Botoes_Click(Index As Integer)
    vOpcao = Index
    Me.Visible = False
End Sub

Private Sub Form_Load()
    vOpcao = -1
    vTopPadrao = 2500
End Sub

Public Function Largura(vLargura As Integer) As FormMensagem
    Set Largura = Me
    Me.Width = vLargura
    Picture1.Width = vLargura
    Picture2.Width = vLargura
    labelMensagem.Width = vLargura
    labelTitulo.Width = vLargura
End Function

Public Function Mensagem(vMensagem As String) As FormMensagem
Attribute Mensagem.VB_UserMemId = 1610809350
    Set Mensagem = Me
    labelMensagem.Caption = vMensagem
End Function

Public Function Mostra() As Integer
Attribute Mostra.VB_UserMemId = 1610809347
    Me.Show 1
    Mostra = vOpcao
    Unload Me
End Function

Public Function Titulo(vTitulo As String) As FormMensagem
Attribute Titulo.VB_UserMemId = 1610809351
    Set Titulo = Me
    Me.Caption = vTitulo
    labelTitulo = vTitulo

End Function

