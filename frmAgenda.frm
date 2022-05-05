VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAgenda 
   Caption         =   "Agenda"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Alterar Agendamento"
            Object.ToolTipText     =   "Alterar agendamento"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Excluir Agendamento"
            Object.ToolTipText     =   "Excluir agendamento"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Confirmar"
            Object.ToolTipText     =   "Confirmar serviço(s)"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Serviços a serem feitos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   5760
      TabIndex        =   13
      Top             =   720
      Width           =   5655
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4335
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   840
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483624
         GridColor       =   0
         FocusRect       =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "|<Descrição                                                                 |>Valor                "
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
         Left            =   3000
         TabIndex        =   18
         Top             =   5280
         Width           =   615
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
         Left            =   3720
         TabIndex        =   17
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   75
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mostrando serviços a serem feitos no cliente:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   3180
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   6480
      Width           =   11295
      Begin VB.Label Label3 
         Caption         =   "Clientes Agendados"
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
         Left            =   120
         TabIndex        =   12
         Top             =   195
         Width           =   5160
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clientes agendados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5535
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   3615
      End
      Begin VB.CommandButton cmdPesClientes 
         Caption         =   "..."
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtCódCliente 
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar clientes"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   5295
      End
      Begin MSComCtl2.DTPicker cboDataInicial 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   38280
      End
      Begin MSComCtl2.DTPicker cboDataFinal 
         Height          =   315
         Left            =   4080
         TabIndex        =   5
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   38280
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   -2147483624
         GridColor       =   0
         FocusRect       =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "|<Cliente                                                         |^Data             |^Hora     "
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "até"
         Height          =   195
         Left            =   3480
         TabIndex        =   4
         Top             =   480
         Width           =   225
      End
      Begin VB.Label Label1 
         Caption         =   "Mostrar clientes agendados para"
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   4560
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   12
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":031A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":0B34
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":0E4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":1168
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":1342
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":151C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":1836
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":1B50
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":1E6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":2684
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmAgenda.frx":285E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DataI As Date
Dim DataF As Date
Dim NomeCliente As String
Dim NAg As String
Dim NAtual As String
Dim ContaAg As Integer

Private Sub cmdMostrar_Click()
    Dim Udata As Date
    
    DataI = cboDataInicial.Day & "/" & cboDataInicial.Month & "/" & cboDataInicial.Year
    DataF = cboDataFinal.Day & "/" & cboDataFinal.Month & "/" & cboDataFinal.Year
    
    If DataF < DataI Then
        MsgBox "A data final está menor que a data inicial!", vbInformation
        cboDataInicial.SetFocus
        Exit Sub
    End If
    
    MSFlexGrid1.Clear
    MSFlexGrid1.FormatString = "|<Cliente                                                         |^Data             |^Hora     "
    MSFlexGrid1.Rows = 1
    
    MSFlexGrid2.Clear
    MSFlexGrid2.FormatString = "|<Descrição                                                                 |>Valor                "
    MSFlexGrid2.Rows = 1
    
    lblCliente = ""
    NAtual = ""
    lblTotal = "0,00"
    Toolbar1.Buttons.Item(1).Enabled = False
    Toolbar1.Buttons.Item(2).Enabled = False
    Toolbar1.Buttons.Item(4).Enabled = False
    
    Clientes.Index = "Chave1"
    If ContaAg > 0 Then
        Agenda.Index = "Chave2"
        Agenda.Seek "<", DataI, 0
        Agenda.MoveNext
        If Agenda.EOF = False Then
            If Agenda("Confirmado") = "Sim" Then
                Do Until Agenda("Confirmado") = "Não"
                    Agenda.MoveNext
                    If Agenda.EOF Then
                        Exit Sub
                    End If
                Loop
                Udata = Agenda("Data")
            Else
                Udata = Agenda("Data")
            End If
        End If
        While Not Agenda.EOF
            If Agenda("Data") > DataF Then
                Agenda.MoveLast
            Else
                If Agenda("Confirmado") = "Não" Then
                    If Trim(Agenda("CódCliente")) = Trim(txtCódCliente.Text) Or Trim(txtCódCliente.Text) = "" Then
                        If Udata <> Agenda("Data") Then
                            MSFlexGrid1.AddItem Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                            Udata = Agenda("Data")
                        End If
                        Clientes.Seek "=", Agenda("CódCliente")
                        If Clientes.NoMatch Then
                            NomeCliente = "Não cadastrado"
                        Else
                            NomeCliente = Clientes("Nome")
                        End If
                        MSFlexGrid1.AddItem Agenda("Número") & Chr(9) & NomeCliente & Chr(9) & Agenda("Data") & Chr(9) & Mid(Agenda("Hora"), 1, 5)
                    End If
                End If
            End If
            Agenda.MoveNext
        Wend
    End If
End Sub

Private Sub cmdPesClientes_Click()
    C_Cliente = ""
    frmClientesPes.Show vbModal
    If Trim(C_Cliente) <> "" Then
        txtCódCliente.Text = C_Cliente
        cmdMostrar.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    cmdMostrar_Click
    txtCódCliente.SetFocus
End Sub

Private Sub Form_Load()
    NAtual = ""
    cboDataInicial.Day = Mid(Date, 1, 2)
    cboDataInicial.Month = Mid(Date, 4, 2)
    cboDataInicial.Year = Mid(Date, 7, 4)
    cboDataFinal.Day = Mid(Date, 1, 2)
    cboDataFinal.Month = Mid(Date, 4, 2)
    cboDataFinal.Year = Mid(Date, 7, 4)
    ContaAg = 0
    ContaAgenda
End Sub

Private Sub ContaAgenda()
    On Error Resume Next
    Agenda.MoveFirst
    Do Until Agenda.EOF Or ContaAg = 1
        Agenda.MoveNext
        ContaAg = ContaAg + 1
    Loop
End Sub

Private Sub MSFlexGrid1_Click()
    Dim NomeServiço As String
    Dim NCLiente As String
    
    If MSFlexGrid1.Rows = 1 Then
        Exit Sub
    End If
    
    MSFlexGrid1.Col = 0
    NAg = MSFlexGrid1.Clip
    MSFlexGrid1.Col = 1
    NCLiente = MSFlexGrid1.Clip
    MSFlexGrid1.Col = 0
    
    If NAg = "" Then
        MSFlexGrid2.Clear
        MSFlexGrid2.FormatString = "|<Descrição                                                                 |>Valor                "
        MSFlexGrid2.Rows = 1
        lblCliente = ""
        NAtual = ""
        lblTotal = "0,00"
        Toolbar1.Buttons.Item(1).Enabled = False
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(4).Enabled = False
        Exit Sub
    End If
    
    NAtual = NAg
    Toolbar1.Buttons.Item(1).Enabled = True
    Toolbar1.Buttons.Item(2).Enabled = True
    Toolbar1.Buttons.Item(4).Enabled = True
    
    MSFlexGrid2.Clear
    MSFlexGrid2.FormatString = "|<Descrição                                                                 |>Valor                "
    MSFlexGrid2.Rows = 1
    
    lblTotal = "0,00"
    SAgenda.Index = "Chave2"
    SAgenda.Seek "=", NAg
    If SAgenda.NoMatch Then
        MSFlexGrid2.AddItem Chr(9) & "Não há serviços a serem feitos." & Chr(9) & "0,00"
    Else
        Serviços.Index = "Chave1"
        While Not SAgenda.EOF
            If Trim(SAgenda("Número")) <> Trim(NAg) Then
                SAgenda.MoveLast
            Else
                Serviços.Seek "=", SAgenda("CódServiço")
                If Serviços.NoMatch Then
                    NomeServiço = "Não cadastrado"
                Else
                    NomeServiço = Serviços("Nome")
                End If
                MSFlexGrid2.AddItem Chr(9) & NomeServiço & Chr(9) & Format(Serviços("Valor"), "##,##0.00")
                lblTotal = Format(CCur(lblTotal) + Serviços("Valor"), "##,##0.00")
            End If
            SAgenda.MoveNext
        Wend
        lblCliente = NCLiente
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "Alterar Agendamento"
            If NAtual = "" Then
                MsgBox "Selecione um agendamento para alterar!", vbInformation
                Exit Sub
            End If
            AlteraAgenda = True
            N_Agenda = NAtual
            frmAgendamento.Show vbModal
        Case "Excluir Agendamento"
            If NAtual = "" Then
                MsgBox "Selecione um agendamento para excluir!", vbInformation
                Exit Sub
            End If
            If MsgBox("Deseja realmente excluir este agendamento de " & lblCliente & "?", vbInformation + vbYesNo) = vbYes Then
                Agenda.Index = "Chave1"
                Agenda.Seek "=", NAtual
                If Agenda.NoMatch = False Then
                    Agenda.Delete
                End If
                
                SAgenda.Index = "Chave2"
                SAgenda.Seek "=", NAtual
                If SAgenda.NoMatch = False Then
                    While Not SAgenda.EOF
                        If Trim(SAgenda("Número")) <> Trim(NAtual) Then
                            SAgenda.MoveLast
                        Else
                            SAgenda.Delete
                        End If
                        SAgenda.MoveNext
                    Wend
                End If
                cmdMostrar_Click
            End If
        Case "Confirmar"
            If NAtual = "" Then
                MsgBox "Selecione um agendamento para confirmar!", vbInformation
                Exit Sub
            End If
            Agenda.Index = "Chave1"
            Agenda.Seek "=", NAtual
            If Agenda.NoMatch Then
                MsgBox "Problemas com este agendamento!", vbInformation
            Else
                CancelaConfirmação = True
                ValorServiço = lblTotal
                C_ClientePagto = lblCliente
                frmPagando.Show vbModal
                If CancelaConfirmação = True Then
                    MsgBox "Operação cancelada!", vbInformation
                    Exit Sub
                Else
                    Agenda.Edit
                    Agenda("Confirmado") = "Sim"
                    Agenda("ValorPago") = ValorPago
                    Agenda("DataPagamento") = Date
                    Agenda.Update
                End If
            End If
        Case "Sair"
            Unload Me
        End Select
End Sub

Private Sub txtCódCliente_Change()
    If Trim(txtCódCliente.Text) = "" Then
        txtNomeCliente.Text = ""
        txtCódCliente.SetFocus
        Exit Sub
    End If
    Clientes.Index = "Chave1"
    Clientes.Seek "=", txtCódCliente.Text
    If Clientes.NoMatch Then
        txtNomeCliente.Text = ""
        txtCódCliente.Text = ""
        txtCódCliente.SetFocus
    Else
        txtNomeCliente.Text = Clientes("Nome")
    End If
End Sub
