VERSION 5.00
Begin VB.Form FormMensagem 
   BorderStyle     =   0  'None
   Caption         =   "Mensagem"
   ClientHeight    =   3165
   ClientLeft      =   2715
   ClientTop       =   3360
   ClientWidth     =   5685
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
   ScaleHeight     =   3165
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3105
      ScaleMode       =   0  'User
      ScaleWidth      =   5654.373
      TabIndex        =   0
      Top             =   0
      Width           =   5655
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
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label labelMensagem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MEnsagem"
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
Public Property Get Opcao() As Integer
    Opcao = vOpcao
End Property

Public Function AdicionaBotao(vCaption As String, vOpcao As Integer, vLeft As Integer, vTop As Integer, vWidth As Integer) As FormMensagem
    Dim Botao As CommandButton
    Set AdicionaBotao = Me
    Load Botoes(vOpcao)
    Botoes(vOpcao).Move vLeft, vTop, vWidth, 375
    Botoes(vOpcao).Caption = vCaption
    Botoes(vOpcao).Visible = True
End Function

Private Sub Botoes_Click(Index As Integer)
    vOpcao = Index
    Me.Visible = False
End Sub

Public Function Mostra() As Integer
    Me.Show 1
    Mostra = vOpcao
    Unload Me
End Function
Public Function Largura(vLargura As Integer) As FormMensagem
    Set Largura = Me
    Me.Width = vLargura
    
End Function
Public Function Altura(vAltura As Integer) As FormMensagem
    Set Altura = Me
    Me.Width = vAltura
End Function
Public Function Mensagem(vMensagem As String) As FormMensagem
    Set Mensagem = Me
    Me.Caption = vMensagem
End Function
Public Function Titulo(vTitulo As String) As FormMensagem
    Set Titulo = Me
    Me.Caption = vTitulo
End Function

Private Sub Form_Load()
    vOpcao = -1
End Sub
