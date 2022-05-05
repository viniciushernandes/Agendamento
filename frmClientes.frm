VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmClientes 
   Caption         =   "Clientes"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Novo"
            Object.ToolTipText     =   "Cadastrar novo cliente"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Alterar"
            Object.ToolTipText     =   "Alterar cliente"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir cliente"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Gravar"
            Object.ToolTipText     =   "Gravar"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Extrato"
            Object.ToolTipText     =   "Extrato do cliente"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   4200
      Width           =   7455
      Begin VB.Label Label1 
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   480
         Left            =   2520
         TabIndex        =   24
         Top             =   195
         Width           =   2400
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Cliente"
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   7455
      Begin VB.CheckBox chkAtivo 
         Caption         =   "Cliente ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5880
         TabIndex        =   10
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtObservações 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   2880
         Width           =   4215
      End
      Begin VB.TextBox txtCelular 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5760
         TabIndex        =   7
         Top             =   2160
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskNascimento 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtTelefone 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtEstado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3720
         TabIndex        =   5
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtCidade 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtBairro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtEndereço 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   3855
      End
      Begin VB.TextBox txtNome 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   5655
      End
      Begin VB.CommandButton cmdPes 
         Caption         =   "..."
         Height          =   315
         Left            =   960
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtCódigo 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblObservações 
         AutoSize        =   -1  'True
         Caption         =   "Observações"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1440
         TabIndex        =   22
         Top             =   2640
         Width           =   945
      End
      Begin VB.Label lblCelular 
         AutoSize        =   -1  'True
         Caption         =   "Celular"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5760
         TabIndex        =   21
         Top             =   1920
         Width           =   480
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   3840
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   11
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":031A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":0B34
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":0E4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":1168
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":1342
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":151C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":1836
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":1B50
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":1E6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClientes.frx":2044
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblNascimento 
         AutoSize        =   -1  'True
         Caption         =   "Nascimento"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   840
      End
      Begin VB.Label lblTelefone 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4200
         TabIndex        =   18
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label lblCidade 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblBairro 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4200
         TabIndex        =   16
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label lblEndereço 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblCódigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Novo As Currency

Private Sub cmdPes_Click()
    C_Cliente = ""
    frmClientesPes.Show vbModal
    If Trim(C_Cliente) <> "" Then
        txtCódigo.Text = C_Cliente
        txtCódigo.SetFocus
    End If
End Sub

Private Sub Form_Load()
    InicializaCampos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "Novo"
            InicializaCampos
            LiberaCampos
            lblCódigo.Enabled = False
            txtCódigo.Enabled = False
            cmdPes.Enabled = False
            Toolbar1.Buttons.Item(1).Enabled = False
            Toolbar1.Buttons.Item(2).Enabled = False
            Toolbar1.Buttons.Item(3).Enabled = False
            Toolbar1.Buttons.Item(5).Enabled = True
            txtNome.SetFocus
        Case "Alterar"
            If Trim(txtCódigo.Text) = "" Then
                txtCódigo.SetFocus
                Exit Sub
            End If
            Clientes.Index = "Chave1"
            Clientes.Seek "=", txtCódigo.Text
            If Clientes.NoMatch Then
                txtCódigo.Text = ""
                txtCódigo.SetFocus
                Exit Sub
            End If
            LiberaCampos
            lblCódigo.Enabled = False
            txtCódigo.Enabled = False
            cmdPes.Enabled = False
            Toolbar1.Buttons.Item(1).Enabled = False
            Toolbar1.Buttons.Item(3).Enabled = False
            Toolbar1.Buttons.Item(5).Enabled = True
            txtNome.SetFocus
        Case "Excluir"
            If Trim(txtCódigo.Text) = "" Then
                txtCódigo.SetFocus
                Exit Sub
            End If
            Clientes.Index = "Chave1"
            Clientes.Seek "=", txtCódigo.Text
            If Clientes.NoMatch Then
                txtCódigo.Text = ""
                txtCódigo.SetFocus
                Exit Sub
            End If
            If MsgBox("Deseja realmente excluir este cliente?", vbInformation + vbYesNo) = vbYes Then
                Clientes.Delete
                InicializaCampos
                txtCódigo.SetFocus
            End If
        Case "Gravar"
            If Trim(txtNome.Text) = "" Then
                MsgBox "Informe o nome do cliente!", vbInformation
                txtNome.SetFocus
                Exit Sub
            End If
            Clientes.Index = "Chave1"
            Clientes.Seek "=", txtCódigo.Text
            If Clientes.NoMatch Then
                Disponível
                Clientes.AddNew
                Clientes("Código") = Novo
            Else
                Clientes.Edit
            End If
            Entradados
            Clientes.Update
            InicializaCampos
            BloqueiaCampos
            txtCódigo.SetFocus
        Case "Cancelar"
            InicializaCampos
            BloqueiaCampos
            txtCódigo.SetFocus
        Case "Sair"
            Unload Me
        Case "Extrato"
            C_Cliente = txtCódigo.Text
            frmExtratoCliente.Show
        End Select
End Sub

Private Sub Disponível()
    Novo = 1
    Clientes.Index = "Chave1"
    Clientes.Seek "=", Novo
    If Clientes.NoMatch = False Then
        Do Until Clientes.NoMatch
            Novo = Novo + 1
            Clientes.Seek "=", Novo
        Loop
    End If
End Sub

Private Sub InicializaCampos()
    txtCódigo.Text = ""
    txtCódigo.MaxLength = 10
    lblCódigo.Enabled = True
    txtCódigo.Enabled = True
    cmdPes.Enabled = True
    txtNome.Text = ""
    txtNome.MaxLength = 50
    txtEndereço.Text = ""
    txtEndereço.MaxLength = 50
    txtBairro.Text = ""
    txtBairro.MaxLength = 50
    txtCidade.Text = ""
    txtCidade.MaxLength = 50
    txtEstado.Text = ""
    txtEstado.MaxLength = 2
    txtTelefone.Text = ""
    txtTelefone.MaxLength = 30
    txtCelular.Text = ""
    txtCelular.MaxLength = 30
    mskNascimento.Mask = ""
    mskNascimento.Text = ""
    mskNascimento.Mask = "##/##/####"
    txtObservações.Text = ""
    txtObservações.MaxLength = 50
    chkAtivo.Value = 0
    
    Toolbar1.Buttons.Item(1).Enabled = True
    Toolbar1.Buttons.Item(2).Enabled = True
    Toolbar1.Buttons.Item(3).Enabled = True
    Toolbar1.Buttons.Item(4).Enabled = True
    Toolbar1.Buttons.Item(5).Enabled = False
    Toolbar1.Buttons.Item(6).Enabled = True
    Toolbar1.Buttons.Item(8).Enabled = False
End Sub

Private Sub BloqueiaCampos()
    lblNome.Enabled = False
    txtNome.Enabled = False
    lblEndereço.Enabled = False
    txtEndereço.Enabled = False
    lblBairro.Enabled = False
    txtBairro.Enabled = False
    lblCidade.Enabled = False
    txtCidade.Enabled = False
    txtEstado.Enabled = False
    lblTelefone.Enabled = False
    txtTelefone.Enabled = False
    lblCelular.Enabled = False
    txtCelular.Enabled = False
    lblNascimento.Enabled = False
    mskNascimento.Enabled = False
    lblObservações.Enabled = False
    txtObservações.Enabled = False
    chkAtivo.Enabled = False
End Sub

Private Sub LiberaCampos()
    lblNome.Enabled = True
    txtNome.Enabled = True
    lblEndereço.Enabled = True
    txtEndereço.Enabled = True
    lblBairro.Enabled = True
    txtBairro.Enabled = True
    lblCidade.Enabled = True
    txtCidade.Enabled = True
    txtEstado.Enabled = True
    lblTelefone.Enabled = True
    txtTelefone.Enabled = True
    lblCelular.Enabled = True
    txtCelular.Enabled = True
    lblNascimento.Enabled = True
    mskNascimento.Enabled = True
    lblObservações.Enabled = True
    txtObservações.Enabled = True
    chkAtivo.Enabled = True
End Sub

Private Sub Entradados()
    Clientes("Nome") = txtNome.Text
    Clientes("Endereço") = txtEndereço.Text
    Clientes("Bairro") = txtBairro.Text
    Clientes("Cidade") = txtCidade.Text
    Clientes("Estado") = txtEstado.Text
    Clientes("Telefone") = txtTelefone.Text
    Clientes("Celular") = txtCelular.Text
    If Not IsDate(mskNascimento.Text) Then
        Clientes("Nascimento") = 0
    Else
        Clientes("Nascimento") = mskNascimento.Text
    End If
    Clientes("Observações") = txtObservações.Text
    If chkAtivo.Value = 0 Then
        Clientes("Ativo") = "Não"
    Else
        Clientes("Ativo") = "Sim"
    End If
End Sub

Private Sub Mostradados()
    txtCódigo.Text = Clientes("Código")
    txtNome.Text = Clientes("Nome")
    txtEndereço.Text = Clientes("Endereço")
    txtBairro.Text = Clientes("Bairro")
    txtCidade.Text = Clientes("Cidade")
    txtEstado.Text = Clientes("Estado")
    txtTelefone.Text = Clientes("Telefone")
    txtCelular.Text = Clientes("Celular")
    If Clientes("Nascimento") = 0 Then
        mskNascimento.Text = "__/__/____"
    Else
        mskNascimento.Text = Clientes("Nascimento")
    End If
    txtObservações.Text = Clientes("Observações")
    If Clientes("Ativo") = "Não" Then
        chkAtivo.Value = 0
    Else
        chkAtivo.Value = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtCódigo_Change()
    If Trim(txtCódigo.Text) = "" Then
        InicializaCampos
        txtCódigo.SetFocus
        Exit Sub
    End If
    Clientes.Index = "Chave1"
    Clientes.Seek "=", txtCódigo.Text
    If Clientes.NoMatch Then
        txtCódigo.Text = ""
        txtCódigo.SetFocus
    Else
        Mostradados
        Toolbar1.Buttons.Item(8).Enabled = True
    End If
End Sub
