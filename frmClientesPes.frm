VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmClientesPes 
   Caption         =   "Consulta de Clientes"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fechar"
      Height          =   315
      Left            =   5640
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmClientesPes.frx":0000
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "frmClientesPes.frx":0014
      TabIndex        =   2
      Top             =   720
      Width           =   6735
   End
   Begin VB.TextBox txtPesquisa 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Digite o nome do cliente"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1710
   End
End
Attribute VB_Name = "frmClientesPes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Data1.DatabaseName = App.Path & "\Dados.mdb"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtPesquisa_Change()
    Dim Pos As Integer
    On Error Resume Next
    If txtPesquisa.SelStart = 0 Then Exit Sub
    Data1.RecordSource = "SELECT Código, Nome FROM Clientes WHERE Nome Like '*" & Mid(txtPesquisa.Text, 1, txtPesquisa.SelStart) & "*' ORDER BY Nome Asc"
    Data1.Refresh
    Pos = txtPesquisa.SelStart
    txtPesquisa.SelStart = Pos
    txtPesquisa.SelLength = Len(txtPesquisa)
End Sub

Private Sub txtPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 8 Then 'Backspace
        txtPesquisa.SelStart = txtPesquisa.SelStart - 1
        txtPesquisa.SelLength = Len(txtPesquisa)
    ElseIf KeyCode = 40 Then
        DBGrid1.SetFocus
    End If
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        DBGrid1.Col = 0
        C_Cliente = DBGrid1.Text
        Unload Me
    End If
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.MarqueeStyle = 3
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.MarqueeStyle = 6
End Sub
