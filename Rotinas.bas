Attribute VB_Name = "Rotinas"
Option Explicit

Public Base As Database
Public Clientes As Recordset
Public Serviços As Recordset
Public Agenda As Recordset
Public SAgenda As Recordset
Public Livre As Recordset
Public C_Cliente As String
Public C_ClientePagto As String
Public C_Serviço As String
Public N_Agenda As String
Public PauBase As Boolean
Public AlteraAgenda As Boolean
Public CancelaConfirmação As Boolean
Public ValorServiço As Currency
Public ValorPago As Currency

Public Sub AbreBase()
    On Error GoTo Erro:
    Set Base = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\Dados.mdb")
    Set Clientes = Base.OpenRecordset("Clientes", dbOpenTable)
    Set Serviços = Base.OpenRecordset("Serviços", dbOpenTable)
    Set Agenda = Base.OpenRecordset("Agenda", dbOpenTable)
    Set SAgenda = Base.OpenRecordset("ServiçosAgenda", dbOpenTable)
    Set Livre = Base.OpenRecordset("Livre", dbOpenTable)
    Exit Sub
Erro:
    MsgBox "Erro na Base de Dados!!!" & vbCrLf & Err.Description, vbCritical
    PauBase = True
End Sub
