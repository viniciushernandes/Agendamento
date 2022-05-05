VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAgendamento 
   Caption         =   "Agendamento"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Gravar"
            Object.ToolTipText     =   "Gravar agendamento"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   10215
      Begin MSMask.MaskEdBox mskHora 
         Height          =   315
         Left            =   8880
         TabIndex        =   6
         Top             =   300
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "hs."
         Height          =   195
         Left            =   9600
         TabIndex        =   22
         Top             =   360
         Width           =   210
      End
      Begin VB.Label lblAgenda 
         Alignment       =   1  'Right Justify
         Caption         =   "Agendando para: 5 de novembro de 2004 �s"
         Height          =   195
         Left            =   4800
         TabIndex        =   20
         Top             =   360
         Width           =   4020
      End
      Begin VB.Label Label5 
         Caption         =   "Agendamento"
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
         Left            =   240
         TabIndex        =   19
         Top             =   195
         Width           =   3480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Agendamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   7080
      TabIndex        =   16
      Top             =   720
      Width           =   3255
      Begin MSACAL.Calendar Calendar1 
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3015
         _Version        =   524288
         _ExtentX        =   5318
         _ExtentY        =   5741
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2004
         Month           =   10
         Day             =   18
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agendar para:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Servi�os"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   6855
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1455
         Left            =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483624
         GridColor       =   0
         FocusRect       =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "|<Descri��o                                                                                        |>Valor              "
      End
      Begin VB.CommandButton cmdAddServi�o 
         Height          =   315
         Left            =   6240
         Picture         =   "frmAgendamento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Adicionar servi�o no agendamento do cliente"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtDescri��o 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   4575
      End
      Begin VB.CommandButton cmdPesServi�os 
         Caption         =   "..."
         Height          =   315
         Left            =   960
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtC�dServi�o 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   4800
         TabIndex        =   24
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   4080
         TabIndex        =   23
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descri��o"
         Height          =   195
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C�digo"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   6855
      Begin VB.TextBox txtC�dCliente 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdPesClientes 
         Caption         =   "..."
         Height          =   315
         Left            =   960
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtNomeCliente 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   480
         Width           =   5055
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   6240
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
               Picture         =   "frmAgendamento.frx":014A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgendamento.frx":0464
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgendamento.frx":0C7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgendamento.frx":0F98
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgendamento.frx":12B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgendamento.frx":148C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgendamento.frx":1666
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgendamento.frx":1980
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgendamento.frx":1C9A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblC�digo 
         AutoSize        =   -1  'True
         Caption         =   "C�digo"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmAgendamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ValorServi�o As Currency
Dim Mes As String
Dim DataAg As Date
Dim HoraAg As Date
Dim NLivre As String
Dim TL As Integer
Dim Cont As Integer
Dim NomeServi�o As String


Private Sub Calendar1_Click()
    
    If Calendar1.Month = 1 Then
        Mes = "Janeiro"
    ElseIf Calendar1.Month = 2 Then
        Mes = "Fevereiro"
    ElseIf Calendar1.Month = 3 Then
        Mes = "Mar�o"
    ElseIf Calendar1.Month = 4 Then
        Mes = "Abril"
    ElseIf Calendar1.Month = 5 Then
        Mes = "Maio"
    ElseIf Calendar1.Month = 6 Then
        Mes = "Junho"
    ElseIf Calendar1.Month = 7 Then
        Mes = "Julho"
    ElseIf Calendar1.Month = 8 Then
        Mes = "Agosto"
    ElseIf Calendar1.Month = 9 Then
        Mes = "Setembro"
    ElseIf Calendar1.Month = 10 Then
        Mes = "Outubro"
    ElseIf Calendar1.Month = 11 Then
        Mes = "Novembro"
    ElseIf Calendar1.Month = 12 Then
        Mes = "Dezembro"
    End If
    lblAgenda.Caption = "Agendando para: " & Calendar1.Day & " de " & Mes & " de " & Calendar1.Year & " �s "
End Sub

Private Sub cmdAddServi�o_Click()
    If Trim(txtC�dServi�o.Text) = "" Then
        txtC�dServi�o.SetFocus
        Exit Sub
    End If
    TL = MSFlexGrid1.Rows
    Cont = 1
    Do Until Cont = TL
        MSFlexGrid1.Row = Cont
        MSFlexGrid1.Col = 0
        If Trim(MSFlexGrid1.Clip) = Trim(txtC�dServi�o.Text) Then
            MsgBox "Servi�o j� incluso!", vbInformation
            txtC�dServi�o.Text = ""
            txtDescri��o.Text = ""
            txtC�dServi�o.SetFocus
            Exit Sub
        End If
        Cont = Cont + 1
    Loop
    MSFlexGrid1.AddItem txtC�dServi�o.Text & Chr(9) & txtDescri��o.Text & Chr(9) & Format(ValorServi�o, "##,##0.00")
    lblTotal = Format(CCur(lblTotal) + ValorServi�o, "##,##0.00")
    txtC�dServi�o.Text = ""
    txtDescri��o.Text = ""
    txtC�dServi�o.SetFocus
End Sub

Private Sub cmdPesClientes_Click()
    C_Cliente = ""
    frmClientesPes.Show vbModal
    If Trim(C_Cliente) <> "" Then
        txtC�dCliente.Text = C_Cliente
        txtC�dServi�o.SetFocus
    End If
End Sub

Private Sub cmdPesServi�os_Click()
    C_Servi�o = ""
    frmServi�osPes.Show vbModal
    If Trim(C_Servi�o) <> "" Then
        txtC�dServi�o.Text = C_Servi�o
        cmdAddServi�o.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Calendar1.Today
    
    If Calendar1.Month = 1 Then
        Mes = "Janeiro"
    ElseIf Calendar1.Month = 2 Then
        Mes = "Fevereiro"
    ElseIf Calendar1.Month = 3 Then
        Mes = "Mar�o"
    ElseIf Calendar1.Month = 4 Then
        Mes = "Abril"
    ElseIf Calendar1.Month = 5 Then
        Mes = "Maio"
    ElseIf Calendar1.Month = 6 Then
        Mes = "Junho"
    ElseIf Calendar1.Month = 7 Then
        Mes = "Julho"
    ElseIf Calendar1.Month = 8 Then
        Mes = "Agosto"
    ElseIf Calendar1.Month = 9 Then
        Mes = "Setembro"
    ElseIf Calendar1.Month = 10 Then
        Mes = "Outubro"
    ElseIf Calendar1.Month = 11 Then
        Mes = "Novembro"
    ElseIf Calendar1.Month = 12 Then
        Mes = "Dezembro"
    End If
    lblAgenda.Caption = "Agendando para: " & Calendar1.Day & " de " & Mes & " de " & Calendar1.Year & " �s "

    If AlteraAgenda = True Then
        NLivre = N_Agenda
        Toolbar1.Buttons.Item(2).Enabled = False
        Agenda.Index = "Chave1"
        Agenda.Seek "=", N_Agenda
        If Agenda.NoMatch Then
            MsgBox "Agendamento n�o encontrado!", vbInformation
            Unload Me
            Exit Sub
        Else
            Clientes.Index = "Chave1"
            Clientes.Seek "=", Agenda("C�dCliente")
            If Clientes.NoMatch Then
                Clientes.AddNew
                Clientes("Nome") = "Cliente n�o cadastrado!"
            End If
            txtC�dCliente.Text = Agenda("C�dCliente")
            txtNomeCliente.Text = Clientes("Nome")
            txtC�dCliente.Enabled = False
            txtNomeCliente.Enabled = False
            cmdPesClientes.Enabled = False
            Calendar1.Day = Mid(Agenda("Data"), 1, 2)
            Calendar1.Month = Mid(Agenda("Data"), 4, 2)
            Calendar1.Year = Mid(Agenda("Data"), 7, 4)
            mskHora.Text = Mid(Agenda("Hora"), 1, 5)
            
            lblTotal = "0,00"
            SAgenda.Index = "Chave2"
            SAgenda.Seek "=", N_Agenda
            If SAgenda.NoMatch = False Then
                Servi�os.Index = "Chave1"
                While Not SAgenda.EOF
                    If Trim(SAgenda("N�mero")) <> Trim(N_Agenda) Then
                        SAgenda.MoveLast
                    Else
                        Servi�os.Seek "=", SAgenda("C�dServi�o")
                        If Servi�os.NoMatch Then
                            NomeServi�o = "N�o cadastrado"
                        Else
                            NomeServi�o = Servi�os("Nome")
                        End If
                        MSFlexGrid1.AddItem SAgenda("C�dServi�o") & Chr(9) & NomeServi�o & Chr(9) & Format(SAgenda("Valor"), "##,##0.00")
                        lblTotal = Format(CCur(lblTotal) + ValorServi�o, "##,##0.00")
                    End If
                    SAgenda.MoveNext
                Wend
            End If
        End If
    End If
End Sub

Private Sub MSFlexGrid1_Click()
    Dim Item As Integer
    Dim Valor As Currency
    
    If MSFlexGrid1.Rows = 1 Then
        txtC�dServi�o.SetFocus
        Exit Sub
    End If
    
    Item = MSFlexGrid1.RowSel
    
    If MsgBox("Deseja excluir este servi�o?", vbInformation + vbYesNo) = vbYes Then
        MSFlexGrid1.Col = 2
        Valor = MSFlexGrid1.Clip
        lblTotal = Format(CCur(lblTotal) - Valor, "##,##0.00")
        MSFlexGrid1.Col = 0
        If MSFlexGrid1.Rows > 2 Then
            MSFlexGrid1.RemoveItem Item
        Else
            MSFlexGrid1.Clear
            MSFlexGrid1.Rows = 1
            MSFlexGrid1.FormatString = "|<Descri��o                                                                                        |>Valor              "
        End If
    End If
    txtC�dServi�o.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "Gravar"
            If Trim(txtC�dCliente.Text) = "" Then
                MsgBox "Informe um cliente!", vbInformation
                txtC�dCliente.SetFocus
                Exit Sub
            End If
            If mskHora.Text = "__:__" Then
                MsgBox "Informe a hora do cliente!", vbInformation
                mskHora.SetFocus
                Exit Sub
            End If
            DataAg = Calendar1.Day & "/" & Calendar1.Month & "/" & Calendar1.Year
            HoraAg = mskHora.Text & ":" & "00"
            If DataAg < Date Then
                If MsgBox("Este agendamento est� no passado. Deseja agendar mesmo assim?", vbInformation + vbYesNo) = vbNo Then
                    Calendar1.SetFocus
                    Exit Sub
                End If
            End If
            If AlteraAgenda = False Then
                PegaLivre
            End If
            Agenda.Index = "Chave2"
            Agenda.Seek "=", DataAg, HoraAg
            If Agenda.NoMatch = False Then
                If Trim(txtC�dCliente.Text) <> Trim(Agenda("C�dCliente")) Then
                    Clientes.Index = "Chave1"
                    Clientes.Seek "=", Agenda("C�dCliente")
                    If Clientes.NoMatch Then
                        Clientes.AddNew
                        Clientes("Nome") = "N�o cadastrado"
                    End If
                    MsgBox "Hor�rio j� agendado para " & Clientes("Nome") & "!" & vbCrLf & "Por favor escolha outro hor�rio.", vbInformation
                    mskHora.SetFocus
                    Exit Sub
                End If
            End If
            Agenda.Index = "Chave1"
            Agenda.Seek "=", NLivre
            If Agenda.NoMatch Then
                Agenda.AddNew
            Else
                Agenda.Edit
            End If
            Entradados
            Agenda.Update
            If AlteraAgenda = False Then
                InicializaCampos
                txtC�dCliente.SetFocus
            Else
                Unload Me
                Exit Sub
            End If
        Case "Cancelar"
            InicializaCampos
            txtC�dCliente.SetFocus
        Case "Sair"
            Unload Me
        End Select
End Sub

Private Sub Entradados()
    Dim C�dServi�o As String
    Dim Valor As Currency
    
    Agenda("N�mero") = NLivre
    Agenda("C�dCliente") = txtC�dCliente.Text
    Agenda("Data") = DataAg
    Agenda("Hora") = HoraAg
    Agenda("Confirmado") = "N�o"
    Agenda("DataPagamento") = 0
    Agenda("ValorPago") = 0
    
    SAgenda.Index = "Chave2"
    SAgenda.Seek "=", NLivre
    If SAgenda.NoMatch = False Then
        While Not SAgenda.EOF
            If Trim(SAgenda("N�mero")) = Trim(NLivre) Then
                SAgenda.Delete
            Else
                SAgenda.MoveLast
            End If
            SAgenda.MoveNext
        Wend
    End If
    
    TL = MSFlexGrid1.Rows
    Cont = 1
    Do Until Cont = TL
        MSFlexGrid1.Row = Cont
        MSFlexGrid1.Col = 0
        C�dServi�o = MSFlexGrid1.Clip
        MSFlexGrid1.Col = 2
        Valor = MSFlexGrid1.Clip
        SAgenda.AddNew
        SAgenda("N�mero") = NLivre
        SAgenda("C�dServi�o") = C�dServi�o
        SAgenda("Valor") = Valor
        SAgenda.Update
        Cont = Cont + 1
    Loop
    
End Sub

Private Sub txtC�dCliente_Change()
    If Trim(txtC�dCliente.Text) = "" Then
        txtC�dCliente.Text = ""
        txtNomeCliente.Text = ""
        txtC�dCliente.SetFocus
        Exit Sub
    End If
    Clientes.Index = "Chave1"
    Clientes.Seek "=", txtC�dCliente.Text
    If Clientes.NoMatch Then
        txtC�dCliente.Text = ""
        txtNomeCliente.Text = ""
        txtC�dCliente.SetFocus
    Else
        txtNomeCliente.Text = Clientes("Nome")
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtC�dCliente_LostFocus()
    If Trim(txtC�dCliente.Text) <> "" Then
        If Clientes("Ativo") = "N�o" Then
            MsgBox "Cliente inativo!", vbCritical
            txtC�dCliente.Text = ""
            txtNomeCliente.Text = ""
            txtC�dCliente.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtC�dServi�o_Change()
    If Trim(txtC�dServi�o.Text) = "" Then
        txtC�dServi�o.Text = ""
        txtDescri��o.Text = ""
        txtC�dServi�o.SetFocus
        Exit Sub
    End If
    Servi�os.Index = "Chave1"
    Servi�os.Seek "=", txtC�dServi�o.Text
    If Servi�os.NoMatch Then
        txtC�dServi�o.Text = ""
        txtDescri��o.Text = ""
        txtC�dServi�o.SetFocus
    Else
        txtDescri��o.Text = Servi�os("Nome")
        ValorServi�o = Servi�os("Valor")
    End If
End Sub

Private Sub InicializaCampos()
    txtC�dCliente.Text = ""
    txtC�dCliente.MaxLength = 10
    txtNomeCliente.Text = ""
    txtC�dServi�o.Text = ""
    txtC�dServi�o.MaxLength = 10
    txtDescri��o.Text = ""
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 1
    MSFlexGrid1.FormatString = "|<Descri��o                                                                                        |>Valor              "
    lblTotal = "0,00"
    Calendar1.Today
    mskHora.Mask = ""
    mskHora.Text = ""
    mskHora.Mask = "##:##"

    If Calendar1.Month = 1 Then
        Mes = "Janeiro"
    ElseIf Calendar1.Month = 2 Then
        Mes = "Fevereiro"
    ElseIf Calendar1.Month = 3 Then
        Mes = "Mar�o"
    ElseIf Calendar1.Month = 4 Then
        Mes = "Abril"
    ElseIf Calendar1.Month = 5 Then
        Mes = "Maio"
    ElseIf Calendar1.Month = 6 Then
        Mes = "Junho"
    ElseIf Calendar1.Month = 7 Then
        Mes = "Julho"
    ElseIf Calendar1.Month = 8 Then
        Mes = "Agosto"
    ElseIf Calendar1.Month = 9 Then
        Mes = "Setembro"
    ElseIf Calendar1.Month = 10 Then
        Mes = "Outubro"
    ElseIf Calendar1.Month = 11 Then
        Mes = "Novembro"
    ElseIf Calendar1.Month = 12 Then
        Mes = "Dezembro"
    End If
    lblAgenda.Caption = "Agendando para: " & Calendar1.Day & " de " & Mes & " de " & Calendar1.Year & " �s "
End Sub

Private Sub PegaLivre()
    Livre.MoveFirst
    NLivre = CCur(Livre("Agenda")) + 1
    Livre.Edit
    Livre("Agenda") = NLivre
    Livre.Update
End Sub
