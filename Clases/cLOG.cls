VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ACCIONES DISPONIBLES
'
'       1.- Eliminar facturas
'       2.- Facturacion
'       3.- Contabilizar facturas
'       4.- Correccion precios
'       5.- Facturacion desde hasta COMPRAS

Public Function Insertar(Accion As Byte, ByRef ElUsuario As CSesion, Descripcion As String) As Boolean
Dim SQL As String

    On Error GoTo EI
    Insertar = False
    
    SQL = "insert into `slog` (`fecha`,`accion`,`usuario`,`pc`,`descripcion`) values ( "
    SQL = SQL & " now()," & Accion & " ,'" & DevNombreSQL(ElUsuario.Login) & "','"
    SQL = SQL & DevNombreSQL(ElUsuario.PC) & "',"
    If Descripcion = "" Then
        SQL = SQL & "NULL)"
    Else
        SQL = SQL & "'" & DevNombreSQL(Descripcion) & "')"
    End If
    Conn.Execute SQL
    
    
    Insertar = True
    Exit Function
EI:
    MuestraError Err.Number, Err.Description

End Function




Public Function DevuelveAcciones(ByRef ListaAcciones As Collection) As Boolean
 On Error GoTo EDevuelveAcciones
    ListaAcciones.Add "1|Eliminar facturas|"
    ListaAcciones.Add "2|Facturación|"
    ListaAcciones.Add "3|Contabilizar facturas|"
    ListaAcciones.Add "4|Corrección precios|"
    ListaAcciones.Add "5|Facturacion COMPRAS|"
    DevuelveAcciones = True
    Exit Function
EDevuelveAcciones:
    MuestraError Err.Number
End Function

