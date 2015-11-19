VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExpedGru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' +-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: LAURA     +-+
' +-+- Fecha: 09/05/06  +-+
' +-+-+-+-+-+-+-+-+-+-+-+-+

'Expedientes de Grupos
'TABLA: expgrupo

''Variables locales que contienen valores de propiedad
Private mNumExped As String 'nº de expediente
Private mCodEmpre As Integer 'codigo empresa
Private mNumVenta As String 'nº de preventa
Private mCodAgenc As Integer 'codigo agencia
Private mCodEmple As Integer 'codigo empleado
Private mEmpleGes As Integer 'codigo empleado gestion
Private mCodClien As Long 'codigo cliente

Private mlocaliza As String 'localizador
Private mcodDesti As String 'destino
Private mcodProdu As Integer 'producto
Private mcodFovia As Long 'folleto
Private mcodItine As Long 'itinerario

Private mNumerPax As Integer 'nº de plazas
Private mVentaPax As Byte 'Tipo de venta (1=Grupo, 2=Sueltas)
Private mDesVentaPax As String '(Grupo / Sueltas)

Private mFechaSal As String 'fecha de salida
Private mFechaReg As String 'fecha de regreso




'------------------------------------------------
'Propiedades del modulo CExpedGru
'------------------------------------------------

'Nº de Expediente
Public Property Let numExped(ByVal vData As String)
     mNumExped = vData
End Property

Public Property Get numExped() As String
     numExped = mNumExped
End Property



'codigo empresa
Public Property Let CodEmpresa(ByVal vData As Integer)
     mCodEmpre = vData
End Property

Public Property Get CodEmpresa() As Integer
     CodEmpresa = mCodEmpre
End Property



'Nº de Venta
Public Property Let NumVenta(ByVal vData As String)
     mNumVenta = vData
End Property

Public Property Get NumVenta() As String
     NumVenta = mNumVenta
End Property


'codigo agencia
Public Property Let CodAgencia(ByVal vData As Integer)
     mCodAgenc = vData
End Property

Public Property Get CodAgencia() As Integer
     CodAgencia = mCodAgenc
End Property



'codigo empleado
Public Property Let CodEmple(ByVal vData As Integer)
     mCodEmple = vData
End Property

Public Property Get CodEmple() As Integer
     CodEmple = mCodEmple
End Property


'codigo empleado gestión
Public Property Let CodEmpleGes(ByVal vData As Integer)
     mEmpleGes = vData
End Property

Public Property Get CodEmpleGes() As Integer
     CodEmpleGes = mEmpleGes
End Property



'codigo cliente
Public Property Let CodCliente(ByVal vData As Long)
     mCodClien = vData
End Property

Public Property Get CodCliente() As Long
     CodCliente = mCodClien
End Property



'localizador
Public Property Let Localizador(ByVal vData As String)
     mlocaliza = vData
End Property

Public Property Get Localizador() As String
      Localizador = mlocaliza
End Property


'codigo destino
Public Property Let CodDestino(ByVal vData As String)
     mcodDesti = vData
End Property

Public Property Get CodDestino() As String
     CodDestino = mcodDesti
End Property



'codigo producto
Public Property Let CodProducto(ByVal vData As Integer)
     mcodProdu = vData
End Property

Public Property Get CodProducto() As Integer
     CodProducto = mcodProdu
End Property


'codigo folleto
Public Property Let CodFolleto(ByVal vData As Long)
     mcodFovia = vData
End Property

Public Property Get CodFolleto() As Long
     CodFolleto = mcodFovia
End Property



'codigo itinerario
Public Property Let CodItine(ByVal vData As Long)
     mcodItine = vData
End Property

Public Property Get CodItine() As Long
     CodItine = mcodItine
End Property





'nº de plazas
Public Property Let NumerPax(ByVal vData As Integer)
     mNumerPax = vData
End Property

Public Property Get NumerPax() As Integer
     NumerPax = mNumerPax
End Property



'tipo venta pax (1=Grupo, 2=Sueltas)
Public Property Let VentaPax(ByVal vData As Integer)
     mVentaPax = vData
End Property

Public Property Get VentaPax() As Integer
     VentaPax = mVentaPax
End Property


'Descripcion venta pax
Public Property Let DesVentaPax(ByVal vData As String)
     mDesVentaPax = vData
End Property

Public Property Get DesVentaPax() As String
     DesVentaPax = mDesVentaPax
End Property



'Fecha de salida
Public Property Let FechaSal(ByVal vData As String)
     mFechaSal = vData
End Property

Public Property Get FechaSal() As String
     FechaSal = mFechaSal
End Property



'Fecha de regreso
Public Property Let FechaReg(ByVal vData As String)
     mFechaReg = vData
End Property

Public Property Get FechaReg() As String
     FechaReg = mFechaReg
End Property





'------------------------------------------------
'Procedimientos del modulo CExpedGru
'------------------------------------------------


'****   LEER  ***
Public Function LeerDatos(numExped As String, codEmpre As String) As Boolean
'Leer los datos del expediente de grupos
'Lee de la BD: PlannerTours, Tabla: expgrupo

    Dim RS As ADODB.Recordset
    Dim SQL As String, mens As String

    On Error GoTo ELeer

    SQL = "SELECT * FROM expgrupo WHERE numexped=" & numExped & " AND codempre=" & codEmpre

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    mens = ""

    If RS.EOF Then
        LeerDatos = False
    Else
        mNumExped = DBLet(RS!numExped, "T")
        mCodEmpre = RS!codEmpre 'codigo empresa
        mNumVenta = DBLet(RS!NumVenta, "T")
        mCodAgenc = RS!codAgenc 'codigo agencia
        mCodEmple = RS!CodEmple 'empleado
        mEmpleGes = DBLet(RS!Empleges, "N") 'empleado gestion
        mCodClien = RS!CodClien 'cliente
        mlocaliza = DBLet(RS!localiza, "T") 'localizador mayorista
        mcodDesti = DBLet(RS!CodDesti, "T") 'destino
        mcodProdu = DBLet(RS!CodProdu, "N") 'producto
        mcodFovia = DBLet(RS!codFovia, "N") 'folleto
        mcodItine = DBLet(RS!CodItine, "N") 'itinerario

        mNumerPax = DBLet(RS!NumerPax, "N") 'nº de plazas
        mVentaPax = DBLet(RS!VentaPax, "N") 'Tipo de venta (1=Grupo, 2=Sueltas)
        
        If mVentaPax = 1 Then
            mDesVentaPax = "GRUPO"
        Else
            mDesVentaPax = "SUELTAS"
        End If
        
        mFechaSal = DBLet(RS!FechaSal, "T") 'fecha de salida
        mFechaReg = DBLet(RS!FechaReg, "T") 'fecha de regreso
        

        mens = "Datos del Expediente de grupos"
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




Public Function DesDestino() As String

    If Me.CodDestino <> "" Then
        DesDestino = DevuelveDesdeBDNew(cPTours, "destinos", "nomdesti", "coddesti", Me.CodDestino, "N")
        If DesDestino = "" Then MsgBox "No existe el destino: " & Format(Me.CodDestino, "0000")
    End If
End Function





