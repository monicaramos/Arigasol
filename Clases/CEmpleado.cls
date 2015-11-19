VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Variables locales que contienen valores de propiedad

Private mCodEmple As Integer 'codigo empleado
Private mNomEmple As String 'nombre empleado
Private mApeEmple As String 'apellido empleado


'------------------------------------------------
'Propiedades del modulo CEmpleado
'------------------------------------------------

'codigo de cliente
Public Property Let CodEmpleado(ByVal vdata As Integer)
     mCodEmple = vdata
End Property

Public Property Get CodEmpleado() As Integer
     CodEmpleado = mCodEmple
End Property


'Nombre
Public Property Let NombreEmple(ByVal vdata As String)
     mNomEmple = vdata
End Property

Public Property Get NombreEmple() As String
     NombreEmple = mNomEmple
End Property


'Apellido
Public Property Let ApellidoEmple(ByVal vdata As String)
     mApeEmple = vdata
End Property

Public Property Get ApellidoEmple() As String
     ApellidoEmple = mApeEmple
End Property





'****   LEER  ***
Public Function LeerDatos(codigo As String) As Boolean
'Leer los datos del empleado
'Lee de la BD: PlannerTours, Tabla: empleado

    Dim RS As ADODB.Recordset
    Dim SQL As String, mens As String

    On Error GoTo ELeer
    
    SQL = "SELECT * FROM empleado WHERE codemple=" & codigo

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    mens = ""
    
    If RS.EOF Then
        LeerDatos = False
    Else
        mCodEmple = RS!CodEmple

        mNomEmple = DBLet(RS!NomEmple, "T")
        mApeEmple = DBLet(RS!Apeemple, "T")
        
        mens = "Datos del empleado"
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    mens = "Se ha producido un error. " & mens & vbCrLf
    mens = mens & "Número: " & Err.Number & vbCrLf
    mens = mens & "Descripción: " & Err.Description
    MsgBox mens, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function


'Public Function NombreCompleto() As String
'    NombreCompleto = NombreEmple & " " & ApellidoEmple
'End Function
