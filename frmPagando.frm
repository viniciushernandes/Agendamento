VERSION 5.00
Begin VB.Form frmPagando 
   Caption         =   "Pagamentos"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtPago 
         Alignment       =   1  'Right Justify
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
         Left            =   2880
         TabIndex        =   0
         Text            =   "0,00"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblTroco 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   2880
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Troco........................"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   9
         Top             =   2280
         Width           =   1950
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor pago................."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   8
         Top             =   1800
         Width           =   1950
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   2880
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total a pagar............."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
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
         TabIndex        =   5
         Top             =   600
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Informe o valor pago e clique em 'OK' para confirmar."
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   3375
   End
End
Attribute VB_Name = "frmPagando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()
    CancelaConfirmação = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    CancelaConfirmação = False
    If CCur(txtPago.Text) > CCur(lblTotal) Then
        ValorPago = CCur(lblTotal)
    Else
        ValorPago = CCur(txtPago.Text)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    lblTotal = Format(ValorServiço, "##,##0.00")
    lblCliente = C_ClientePagto
End Sub

Private Sub txtPago_GotFocus()
    txtPago.SelStart = 0
    txtPago.SelLength = Len(txtPago.Text)
End Sub

Private Sub txtPago_LostFocus()
    If Trim(txtPago.Text) = "" Then
        txtPago.Text = 0
    End If
    txtPago.Text = Format(txtPago.Text, "##,##0.00")
    lblTroco = Format(CCur(txtPago.Text) - CCur(lblTotal), "##,##0.00")
    If CCur(lblTroco) < 0 Then
        lblTroco = "0,00"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

