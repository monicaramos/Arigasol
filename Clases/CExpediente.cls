VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExpediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Explicit
' +-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: LAURA     +-+
' +-+- Fecha: 09/05/06  +-+
' +-+-+-+-+-+-+-+-+-+-+-+-+

'TABLA: expgrupo

'Variables locales que contienen valores de propiedad
Private mNumVenta As String 'nº de venta
Private mNumExped As String 'nº de expediente
Private mFechVent As String 'fecha de la venta
Private mClasVent As Byte 'clase de la venta (1=Individuales, 2=Grupos)
Private mClaseAge As Byte 'clase de la agencia (1=Minorista, 2=Mayorista)
Private mCodEmpre As Integer 'codigo empresa
Private mCodAgenc As Integer 'codigo agencia
Private mEstVenta As Byte 'estado de la venta (1=Pendiente, 6=En Gestión)
Private mCodClien As Long 'codigo cliente
                                         
Private mCodEmple As Integer 'codigo empleado
Private mCodEmges As Integer 'codigo empleado gestion

Private mSolicita As String 'quien solicita la venta



'parametros para ventas de Minorista(Indv. y grupos) y Mayorista (Grupos)
'-------------------------------------------------------------------------
Private mNumPlaza As Integer 'nº de plazas


'parametros para ventas de Mayorista (Indiv.)
'--------------------------------------------
Private mNumAdult As Integer 'nº de adultos
Private mNumNinyo As Integer 'nº de niños
Private mLocaliza As String 'localizador
Private mFechaSal As String 'fecha de salida






'------------------------------------------------
'Propiedades del modulo CVenta
'------------------------------------------------

'Nº de venta
Public Property Let NumVenta(ByVal vdata As String)
     mNumVenta = vdata
End Property

Public Property Get NumVenta() As String
     NumVenta = mNumVenta
End Property



'Nº de Expediente
Public Property Let NumExped(ByVal vdata As String)
     mNumExped = vdata
End Property

Public Property Get NumExped() As String
     NumExped = mNumExped
End Property


'Fecha de venta
Public Property Let FechaVenta(ByVal vdata As String)
     mFechVent = vdata
End Property

Public Property Get FechaVenta() As String
     FechaVenta = mFechVent
End Property


'clase de venta (indiv o grupos)
Public Property Let ClaseVenta(ByVal vdata As Byte)
     mClasVent = vdata
End Property

Public Property Get ClaseVenta() As Byte
     ClaseVenta = mClasVent
End Property


'clase de agencia (minorista o mayorista)
Public Property Let ClaseAgencia(ByVal vdata As Byte)
     mClasVent = vdata
End Property

Public Property Get ClaseAgencia() As Byte
     ClaseAgencia = mClasVent
End Property


'codigo empresa
Public Property Let CodEmpresa(ByVal vdata As Integer)
     mCodEmpre = vdata
End Property

Public Property Get CodEmpresa() As Integer
     CodEmpresa = mCodEmpre
End Property


'codigo agencia
Public Property Let CodAgencia(ByVal vdata As Integer)
     mCodAgenc = vdata
End Property

Public Property Get CodAgencia() As Integer
     CodAgencia = mCodAgenc
End Property


'estado de la venta: 1=Pendiente, 6=En gestion,
Public Property Let EstadoVenta(ByVal vdata As Byte)
     mEstVenta = vdata
End Property

Public Property Get EstadoVenta() As Byte
     EstadoVenta = mEstVenta
End Property


'codigo cliente
Public Property Let CodCliente(ByVal vdata As Long)
     mCodClien = vdata
End Property

Public Property Get CodCliente() As Long
     CodCliente = mCodClien
End Property


'codigo empleado
Public Property Let CodEmple(ByVal vdata As Integer)
     mCodEmple = vdata
End Property

Public Property Get CodEmple() As Integer
     CodEmple = mCodEmple
End Property


'codigo empleado gestión
Public Property Let CodEmpleGes(ByVal vdata As Integer)
     mCodEmges = vdata
End Property

Public Property Get CodEmpleGes() As Integer
     CodEmpleGes = mCodEmges
End Property



'quien solicita la venta
Public Property Let Solicita(ByVal vdata As String)
     mSolicita = vdata
End Property

Public Property Get Solicita() As String
     Solicita = mSolicita
End Property




'nº de plazas
Public Property Let NumPlazas(ByVal vdata As Integer)
     mNumPlaza = vdata
End Property

Public Property Get NumPlazas() As Integer
     NumPlazas = mNumPlaza
End Property


'nº de adultos
Public Property Let NumAdultos(ByVal vdata As Integer)
     mNumAdult = vdata
End Property

Public Property Get NumAdultos() As Integer
     NumAdultos = mNumAdult
End Property


'nº de niños
Public Property Let NumNinyos(ByVal vdata As Integer)
     mNumNinyo = vdata
End Property

Public Property Get NumNinyos() As Integer
     NumNinyos = mNumNinyo
End Property


'Localiza
Public Property Let Localizador(ByVal vdata As String)
     mLocaliza = vdata
End Property

Public Property Get Localizador() As String
     Localizador = mLocaliza
End Property

'Fecha de salida
Public Property Let FechaSal(ByVal vdata As String)
     mFechaSal = vdata
End Property

Public Property Get FechaSal() As String
     FechaSal = mFechaSal
End Property




'------------------------------------------------
'Procedimientos del modulo CVenta
'------------------------------------------------


'****   LEER  ***
Public Function LeerDatos(codigo1 As String, codigo2 As String) As Boolean
'Leer los datos del proveedor
'Lee de la BD: PlannerTours, Tabla: Preventa

    Dim RS As ADODB.Recordset
    Dim SQL As String, mens As String

    On Error GoTo ELeer
    
    SQL = "SELECT * FROM preventa WHERE numventa=" & codigo1 & " AND codempre=" & codigo2 'vUsu.CadenaConexion

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    mens = ""
    
    If RS.EOF Then
        LeerDatos = False
    Else
        mNumVenta = RS!NumVenta
        mNumExped = DBLet(RS!NumExped, "T")
        mFechVent = DBLet(RS!fechvent, "F") 'fecha venta
        mClasVent = RS!ClasVent 'clase de la venta (1=Individuales, 2=Grupos)
        mClaseAge = RS!ClaseAge 'clase de la agencia (1=Minorista, 2=Mayorista)
        mCodEmpre = RS!codEmpre 'codigo empresa
        mCodAgenc = RS!codAgenc 'codigo agencia
        mEstVenta = RS!EstVenta 'estado de la venta (1=Pendiente, 6=En Gestión)
        mCodClien = RS!CodClien 'cliente
        mCodEmple = RS!CodEmple 'empleado
        mCodEmges = DBLet(RS!CodEmges, "N") 'empleado gestion
        mSolicita = DBLet(RS!Solicita, "T") 'quien lo solicita
                
        mNumPlaza = DBLet(RS!NumPlaza, "N") 'nº de plazas
        mNumAdult = DBLet(RS!NumAdult, "N") 'nº de adultos
        mNumNinyo = DBLet(RS!NumNinyo, "N") 'nº de ninyos
        mLocaliza = DBLet(RS!localiza, "T") 'localizador mayorista
        mFechaSal = DBLet(RS!FechaSal, "F") 'fecha de salida
                
        mens = "Datos de la venta"
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



'Public Function EstadoPresuModificable() As Boolean
''Comprueba si se va a poder modificar el estado de los presupuestos de una venta
'Dim b As Boolean
'Dim SQL As String
'
'    'si ya se ha creado un expediente para la venta, ya no se podra modificar el estado
'    'del presupuesto
'    If NumExped <> "" Then
'        b = False
'    Else
'        'si la venta ya tiene una confirmacion de servicions tampoco se podra modificar
'        'el estado de los presupuestos
'        SQL = "SELECT count(*) FROM confserv WHERE numventa=" & NumVenta & " AND codempre=" & CodEmpresa
'        If RegistrosAListar(SQL) > 0 Then
'            b = False
'        Else
'            'si la venta ya tiene un contrato de viaje combinado ya no se puede modificar estado presupuesto
'            SQL = "SELECT count(*) FROM contrviaj WHERE numventa=" & NumVenta & " AND codempre=" & CodEmpresa
'            If RegistrosAListar(SQL) > 0 Then
'                b = False
'            Else
'                b = True
'            End If
'        End If
'    End If
'    EstadoPresuModificable = b
'End Function


