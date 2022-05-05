Attribute VB_Name = "Rotinas"
Option Explicit

Public Base As Database
Public Clientes As Recordset
Public Servi�os As Recordset
Public Agenda As Recordset
Public SAgenda As Recordset
Public Livre As Recordset
Public C_Cliente As String
Public C_ClientePagto As String
Public C_Servi�o As String
Public N_Agenda As String
Public PauBase As Boolean
Public AlteraAgenda As Boolean
Public CancelaConfirma��o As Boolean
Public ValorServi�o As Currency
Public ValorPago As Currency

Public Sub AbreBase()
    On Error GoTo Erro:
    Set Base = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\Dados.mdb")
    Set Clientes = Base.OpenRecordset("Clientes", dbOpenTable)
    Set Servi�os = Base.OpenRecordset("Servi�os", dbOpenTable)
    Set Agenda = Base.OpenRecordset("Agenda", dbOpenTable)
    Set SAgenda = Base.OpenRecordset("Servi�osAgenda", dbOpenTable)
    Set Livre = Base.OpenRecordset("Livre", dbOpenTable)
    Exit Sub
Erro:
    MsgBox "Erro na Base de Dados!!!" & vbCrLf & Err.Description, vbCritical
    PauBase = True
End Sub
