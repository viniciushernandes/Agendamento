VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmServiços 
   Caption         =   "Serviços"
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
      TabIndex        =   5
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
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Novo"
            Object.ToolTipText     =   "Cadastrar novo serviço"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Alterar"
            Object.ToolTipText     =   "Alterar serviço"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir serviço"
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
      TabIndex        =   9
      Top             =   4200
      Width           =   7455
      Begin VB.Label Label1 
         Caption         =   "Serviços"
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
         TabIndex        =   10
         Top             =   200
         Width           =   2400
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Serviço"
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   7455
      Begin VB.TextBox txtObservações 
         Enabled         =   0   'False
         Height          =   1695
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1560
         Width           =   5655
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtCódigo 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdPes 
         Caption         =   "..."
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtNome 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label lblObservações 
         AutoSize        =   -1  'True
         Caption         =   "Informações sobre este serviço"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   1320
         Width           =   2205
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Enabled         =   0   'False
         Height          =   195
         Left            =   960
         TabIndex        =   11
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label lblCódigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   720
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
            NumListImages   =   9
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmServiços.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmServiços.frx":031A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmServiços.frx":0B34
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmServiços.frx":0E4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmServiços.frx":1168
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmServiços.frx":1342
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmServiços.frx":151C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmServiços.frx":1836
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmServiços.frx":1B50
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmServiços"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Novo As Currency

Private Sub cmdPes_Click()
    C_Serviço = ""
    frmServiçosPes.Show vbModal
    If Trim(C_Serviço) <> "" Then
        txtCódigo.Text = C_Serviço
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
            Serviços.Index = "Chave1"
            Serviços.Seek "=", txtCódigo.Text
            If Serviços.NoMatch Then
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
            Serviços.Index = "Chave1"
            Serviços.Seek "=", txtCódigo.Text
            If Serviços.NoMatch Then
                txtCódigo.Text = ""
                txtCódigo.SetFocus
                Exit Sub
            End If
            If MsgBox("Deseja realmente excluir este serviço?", vbInformation + vbYesNo) = vbYes Then
                Serviços.Delete
                InicializaCampos
                txtCódigo.SetFocus
            End If
        Case "Gravar"
            If Trim(txtNome.Text) = "" Then
                MsgBox "Informe a descrição do serviço!", vbInformation
                txtNome.SetFocus
                Exit Sub
            End If
            Serviços.Index = "Chave1"
            Serviços.Seek "=", txtCódigo.Text
            If Serviços.NoMatch Then
                Disponível
                Serviços.AddNew
                Serviços("Código") = Novo
            Else
                Serviços.Edit
            End If
            Entradados
            Serviços.Update
            InicializaCampos
            BloqueiaCampos
            txtCódigo.SetFocus
        Case "Cancelar"
            InicializaCampos
            BloqueiaCampos
            txtCódigo.SetFocus
        Case "Sair"
            Unload Me
        End Select
End Sub

Private Sub Disponível()
    Novo = 1
    Serviços.Index = "Chave1"
    Serviços.Seek "=", Novo
    If Serviços.NoMatch = False Then
        Do Until Serviços.NoMatch
            Novo = Novo + 1
            Serviços.Seek "=", Novo
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
    txtValor.Text = 0
    txtValor.Text = Format(txtValor.Text, "0.00")
    txtObservações.Text = ""
    
    Toolbar1.Buttons.Item(1).Enabled = True
    Toolbar1.Buttons.Item(2).Enabled = True
    Toolbar1.Buttons.Item(3).Enabled = True
    Toolbar1.Buttons.Item(4).Enabled = True
    Toolbar1.Buttons.Item(5).Enabled = False
    Toolbar1.Buttons.Item(6).Enabled = True
End Sub

Private Sub BloqueiaCampos()
    lblNome.Enabled = False
    txtNome.Enabled = False
    lblValor.Enabled = False
    txtValor.Enabled = False
    lblObservações.Enabled = False
    txtObservações.Enabled = False
End Sub

Private Sub LiberaCampos()
    lblNome.Enabled = True
    txtNome.Enabled = True
    lblValor.Enabled = True
    txtValor.Enabled = True
    lblObservações.Enabled = True
    txtObservações.Enabled = True
End Sub

Private Sub Entradados()
    Serviços("Nome") = txtNome.Text
    Serviços("Valor") = txtValor.Text
    If Trim(txtObservações.Text) = "" Then
        Serviços("Observações") = "vazio"
    Else
        Serviços("Observações").AppendChunk (txtObservações.Text)
    End If
End Sub

Private Sub Mostradados()
    txtCódigo.Text = Serviços("Código")
    txtNome.Text = Serviços("Nome")
    txtValor.Text = Format(Serviços("Valor"), "##,##0.00")
    txtObservações.Text = Serviços("Observações").GetChunk(0, 32768)
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
    Serviços.Index = "Chave1"
    Serviços.Seek "=", txtCódigo.Text
    If Serviços.NoMatch Then
        txtCódigo.Text = ""
        txtCódigo.SetFocus
    Else
        Mostradados
    End If
End Sub

Private Sub txtValor_GotFocus()
    txtValor.SelStart = 0
    txtValor.SelLength = Len(txtValor.Text)
End Sub

Private Sub txtValor_LostFocus()
    If Trim(txtValor.Text) = "" Then
        txtValor.Text = 0
    End If
    txtValor.Text = Format(txtValor.Text, "##,##0.00")
End Sub
