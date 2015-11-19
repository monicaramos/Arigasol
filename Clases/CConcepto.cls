VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Variables locales que contienen valores de propiedad
Private mCodConce As String 'codigo
Private mNomConce As String 'nombre concepto

Private mSignoCon As String * 1 'Signo operacion: + o -
Private mComision As Boolean 'Si es o no comisionable
Private mCodigIva As Byte 'Codigo de IVA

'Datos del tipo de IVA
'---------------------
Private mPorceIva As Currency 'Codigo de IVA




'------------------------------------------------
'Propiedades del modulo CConcepto
'------------------------------------------------

'Nombre concepto
Public Property Let NomConce(ByVal vData As String)
     mNomConce = vData
End Property

Public Property Get NomConce() As String
     NomConce = mNomConce
End Property

'Signo concepto
Public Property Let SignoConce(ByVal vData As String)
     mSignoCon = vData
End Property

Public Property Get SignoConce() As String
     SignoConce = mSignoCon
End Property


'Es comisionable
Public Property Let Comisionable(ByVal vData As Boolean)
     mComision = vData
End Property

Public Property Get Comisionable() As Boolean
     Comisionable = mComision
End Property


'Codigo IVA
Public Property Let CodIVA(ByVal vData As Byte)
     mCodigIva = vData
End Property

Public Property Get CodIVA() As Byte
     CodIVA = mCodigIva
End Property


'% del tipo de IVA
Public Property Let PorceIVA(ByVal vData As Currency)
     mPorceIva = vData
End Property

Public Property Get PorceIVA() As Currency
     PorceIVA = mPorceIva
End Property

'****   LEER  ***
Public Function LeerDatos(codigo As String) As Boolean
'Leer los datos de conceptos
'Lee de la BD: PlannerTours, Tabla: concegrp

    Dim RS As ADODB.Recordset
    Dim SQL As String, mens As String

    On Error GoTo ELeer
    
    SQL = "SELECT * FROM concegrp WHERE codconce=" & codigo

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    mens = ""
    
    If RS.EOF Then
        LeerDatos = False
    Else
        mCodConce = RS!codConce
        mNomConce = RS!NomConce
        mSignoCon = DBLet(RS!signocon, "T")
        
        mComision = DBLet(RS!comision, "N")
        mCodigIva = DBLet(RS!CodigIva, "N")
        
        mens = "Datos del concepto"
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



Public Function LeerIVA() As Boolean
'Leer los datos del codigo de IVA del concepto
'Lee de la BD: Conta, Tabla: tiposiva

    Dim RS As ADODB.Recordset
    Dim SQL As String, mens As String

    On Error GoTo ELeer
    
    SQL = "SELECT * FROM tiposiva WHERE codigiva=" & mCodigIva

    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    mens = ""
    
    If RS.EOF Then
        LeerIVA = False
    Else
'        mCodConce = RS!codConce
        mPorceIva = RS!PorceIVA
        
        mens = "Datos del tipo de IVA"
        LeerIVA = True
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
    LeerIVA = False
End Function

