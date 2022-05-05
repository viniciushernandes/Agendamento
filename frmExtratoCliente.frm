VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExtratoCliente 
   Caption         =   "Extrato de Clientes"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Pagamento"
            Object.ToolTipText     =   "Pagamento"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Agendamentos"
            Object.ToolTipText     =   "Mostrar todos os agendamentos confirmados"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin VB.TextBox txtGeralCliente 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Valor já pago pelo cliente:"
         Top             =   240
         Visible         =   0   'False
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   8895
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6840
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total a pagar"
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
         Left            =   5640
         TabIndex        =   7
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Extrato de clientes"
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
         TabIndex        =   6
         Top             =   195
         Width           =   5160
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8895
      Begin ComctlLib.ImageList ImageList1 
         Left            =   4200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   3
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmExtratoCliente.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmExtratoCliente.frx":031A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmExtratoCliente.frx":0634
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         Caption         =   "Clientes Agendados"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3570
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   -2147483624
      GridColor       =   0
      FocusRect       =   2
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "|^Data             |^Hora       |>Total             |>Valor Pago    "
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   4095
      Left            =   4560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   -2147483624
      GridColor       =   0
      FocusRect       =   2
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "|<Serviço(s)                                                    |>Valor          "
   End
End
Attribute VB_Name = "frmExtratoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Total As Currency
Dim NAg As String
Dim TotalS As Currency
Dim TotalP As Currency
Dim Geral As Currency


Private Sub Form_Activate()
    
    MSFlexGrid1.Clear
    MSFlexGrid1.FormatString = "|^Data             |^Hora       |>Total             |>Valor Pago    "
    MSFlexGrid1.Rows = 1
    
    MSFlexGrid2.Clear
    MSFlexGrid2.FormatString = "|<Serviço(s)                                                    |>Valor          "
    MSFlexGrid2.Rows = 1
    
    TotalP = 0
    
    Agenda.Index = "Chave3"
    Agenda.Seek "<", C_Cliente, 0, 0
    Agenda.MoveNext
    If Agenda.EOF Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Exit Sub
    End If
    If Trim(Agenda("CódCliente")) = Trim(C_Cliente) Then
        SAgenda.Index = "Chave2"
        While Not Agenda.EOF
            Total = 0
            If Trim(Agenda("CódCliente")) <> Trim(C_Cliente) Then
                Agenda.MoveLast
            Else
                If Trim(Agenda("Confirmado")) = "Sim" Then
                    SAgenda.Seek "=", Agenda("Número")
                    If SAgenda.NoMatch = False Then
                        While Not SAgenda.EOF
                            If Trim(SAgenda("Número")) <> Trim(Agenda("Número")) Then
                                SAgenda.MoveLast
                            Else
                                Total = Total + SAgenda("Valor")
                            End If
                            SAgenda.MoveNext
                        Wend
                        If CCur(Agenda("ValorPago")) < CCur(Total) Then
                            MSFlexGrid1.AddItem Agenda("Número") & Chr(9) & Agenda("Data") & Chr(9) & Mid(Agenda("Hora"), 1, 5) & Chr(9) & Format(Total, "##,##0.00") & Chr(9) & Format(Agenda("ValorPago"), "##,##0.00")
                            TotalP = TotalP + (CCur(Total) - CCur(Agenda("ValorPago")))
                        End If
                    End If
                End If
            End If
            Agenda.MoveNext
        Wend
        lblTotal = Format(TotalP, "##,##0.00")
    Else
        Toolbar1.Buttons.Item(2).Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Clientes.Index = "Chave1"
    Clientes.Seek "=", C_Cliente
    If Clientes.NoMatch Then
        MsgBox "Problemas com este cliente!", vbInformation
        Unload Me
        Exit Sub
    End If
    lblCliente = Clientes("Nome") & " "
    NAg = ""
End Sub

Private Sub MSFlexGrid1_Click()
    Dim NomeServiço As String
    Dim NCLiente As String
    
    If MSFlexGrid1.Rows = 1 Then
        Exit Sub
    End If
    
    MSFlexGrid1.Col = 0
    NAg = MSFlexGrid1.Clip
            
    MSFlexGrid2.Clear
    MSFlexGrid2.FormatString = "|<Serviço(s)                                                    |>Valor          "
    MSFlexGrid2.Rows = 1
    
    Agenda.Index = "Chave1"
    Agenda.Seek "=", NAg
    If Agenda.NoMatch Then
        MsgBox "Problemas com este agendamento!", vbInformation
        Exit Sub
    End If
    
    TotalS = 0
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
                MSFlexGrid2.AddItem Chr(9) & NomeServiço & Chr(9) & Format(SAgenda("Valor"), "##,##0.00")
                TotalS = Format(CCur(TotalS) + Serviços("Valor"), "##,##0.00")
            End If
            SAgenda.MoveNext
        Wend
    End If
    If Agenda("ValorPago") < TotalS Then
        Toolbar1.Buttons.Item(1).Enabled = True
    Else
        Toolbar1.Buttons.Item(1).Enabled = False
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "Pagamento"
            If NAg = "" Then
                MsgBox "Selecione uma conta para pagar!", vbInformation
                Exit Sub
            End If
            Agenda.Index = "Chave1"
            Agenda.Seek "=", NAg
            If Agenda.NoMatch Then
                MsgBox "Problemas com este agendamento!", vbInformation
            Else
                CancelaConfirmação = True
                ValorServiço = TotalS - Agenda("ValorPago")
                C_ClientePagto = lblCliente
                frmPagando.Show vbModal
                If CancelaConfirmação = True Then
                    MsgBox "Operação cancelada!", vbInformation
                Else
                    Agenda.Edit
                    Agenda("Confirmado") = "Sim"
                    Agenda("ValorPago") = Agenda("ValorPago") + ValorPago
                    Agenda("DataPagamento") = Date
                    Agenda.Update
                    Form_Activate
                End If
                Toolbar1.Buttons.Item(1).Enabled = False
                txtGeralCliente.Visible = False
            End If
        Case "Agendamentos"
            TodosAgendamentos
        Case "Sair"
            Unload Me
        End Select
End Sub

Private Sub TodosAgendamentos()
    
    MSFlexGrid1.Clear
    MSFlexGrid1.FormatString = "|^Data             |^Hora       |>Total             |>Valor Pago    "
    MSFlexGrid1.Rows = 1
    
    MSFlexGrid2.Clear
    MSFlexGrid2.FormatString = "|<Serviço(s)                                                    |>Valor          "
    MSFlexGrid2.Rows = 1
    
    TotalP = 0
    Geral = 0
    
    Agenda.Index = "Chave3"
    Agenda.Seek "<", C_Cliente, 0, 0
    Agenda.MoveNext
    If Agenda.EOF Then
        Exit Sub
    End If
    If Trim(Agenda("CódCliente")) = Trim(C_Cliente) Then
        SAgenda.Index = "Chave2"
        While Not Agenda.EOF
            Total = 0
            If Trim(Agenda("CódCliente")) <> Trim(C_Cliente) Then
                Agenda.MoveLast
            Else
                If Trim(Agenda("Confirmado")) = "Sim" Then
                    SAgenda.Seek "=", Agenda("Número")
                    If SAgenda.NoMatch = False Then
                        While Not SAgenda.EOF
                            If Trim(SAgenda("Número")) <> Trim(Agenda("Número")) Then
                                SAgenda.MoveLast
                            Else
                                Total = Total + SAgenda("Valor")
                            End If
                            SAgenda.MoveNext
                        Wend
                        MSFlexGrid1.AddItem Agenda("Número") & Chr(9) & Agenda("Data") & Chr(9) & Mid(Agenda("Hora"), 1, 5) & Chr(9) & Format(Total, "##,##0.00") & Chr(9) & Format(Agenda("ValorPago"), "##,##0.00")
                        Geral = Geral + Agenda("ValorPago")
                        If CCur(Agenda("ValorPago")) < CCur(Total) Then
                            TotalP = TotalP + (CCur(Total) - CCur(Agenda("ValorPago")))
                        End If
                    End If
                End If
            End If
            Agenda.MoveNext
        Wend
        lblTotal = Format(TotalP, "##,##0.00")
        txtGeralCliente.Visible = True
        txtGeralCliente.Text = "Valor já pago pelo cliente: " & FormatCurrency(Geral)
    End If
End Sub
