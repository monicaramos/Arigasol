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
Public Property Let CodEmpleado(ByVal vData As Integer)
     mCodEmple = vData
End Property

Public Property Get CodEmpleado() As Integer
     CodEmpleado = mCodEmple
End Property


'Nombre
Public Property Let NombreEmple(ByVal vData As String)
     mNomEmple = vData
End Property

Public Property Get NombreEmple() As String
     NombreEmple = mNomEmple
End Property


'Apellido
Public Property Let ApellidoEmple(ByVal vData As String)
     mApeEmple = vData
End Property

Public Property Get ApellidoEmple() As String
     ApellidoEmple = mApeEmple
End Property





'****   LEER  ***
Public Function LeerDatos(Codigo As String) As Boolean
'Leer los datos del empleado
'Lee de la BD: PlannerTours, Tabla: empleado

    Dim RS As ADODB.Recordset
    Dim SQL As String, Mens As String

    On Error GoTo ELeer
    
    SQL = "SELECT * FROM empleado WHERE codemple=" & Codigo

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If RS.EOF Then
        LeerDatos = False
    Else
        mCodEmple = RS!CodEmple

        mNomEmple = DBLet(RS!NomEmple, "T")
        mApeEmple = DBLet(RS!Apeemple, "T")
        
        Mens = "Datos del empleado"
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function


