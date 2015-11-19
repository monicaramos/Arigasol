VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Variables locales que contienen valores de propiedad del cliente
Private mCodClien As String 'copia local
Private mTipClien As Byte 'tipo de cliente (1=Persona, 2=Empresa)
Private mApeRaso As String 'apellido si es Persona / razon social si es empresa
Private mNomCome As String 'Nombre si es Persona / nombre comercial si es empesa
Private mNIFCIF As String 'NIF si es Persona / CIF si es Empresa
Private mDomClien As String 'domicilio del cliente


'Datos de facturacion del CLIENTE
'almacenados en la tabla: cltefact
Private mCodForpa As Integer 'Forma de pago
Private mDesCodForpa As String 'Descripcion de la forma de pago



'Private mWebCom As String 'web comercial
'Private mMailCom As String 'email comercial




'------------------------------------------------
'Propiedades del modulo CCliente
'------------------------------------------------

'Tipo de cliente (1=Persona, 2=Empresa)
Public Property Let TipoClien(ByVal vData As Byte)
     mTipClien = vData
End Property

Public Property Get TipoClien() As Byte
     TipoClien = mTipClien
End Property


'Apellido cliente / Razon social
Public Property Let Ape_RazSoc(ByVal vData As String)
     mApeRaso = vData
End Property

Public Property Get Ape_RazSoc() As String
     Ape_RazSoc = mApeRaso
End Property


'Nombre  / nombre comercial
Public Property Let Nom_Come(ByVal vData As String)
     mNomCome = vData
End Property

Public Property Get Nom_Come() As String
     Nom_Come = mNomCome
End Property


'NIF  / CIF
Public Property Let NIF_CIF(ByVal vData As String)
     mNIFCIF = vData
End Property

Public Property Get NIF_CIF() As String
     NIF_CIF = mNIFCIF
End Property



'Domicilio
Public Property Let Domicilio(ByVal vData As String)
     mDomClien = vData
End Property

Public Property Get Domicilio() As String
     Domicilio = mDomClien
End Property




'Forma de pago del cliente con una empresa
Public Property Let ForPago(ByVal vData As Integer)
     mCodForpa = vData
End Property

Public Property Get ForPago() As Integer
     ForPago = mCodForpa
End Property


'Descripcion de la forma de pago
Public Property Let DescForPago(ByVal vData As String)
     mDesCodForpa = vData
End Property

Public Property Get DescForPago() As String
     DescForPago = mDesCodForpa
End Property



'****   LEER  ***
Public Function LeerDatos(Codigo As String) As Boolean
'Leer los datos del cliente
'Lee de la BD: PlannerTours, Tabla: Cliente

    Dim RS As ADODB.Recordset
    Dim SQL As String, Mens As String

    On Error GoTo ELeer
    
    SQL = "SELECT * FROM clientes WHERE codclien=" & Codigo 'vUsu.CadenaConexion

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If RS.EOF Then
        LeerDatos = False
    Else
        mCodClien = RS!codClien
        mTipClien = RS!tipclien

        mApeRaso = DBLet(RS!Ape_Raso, "T")
        mNomCome = DBLet(RS!Nom_Come, "T")
        mNIFCIF = DBLet(RS!NIF_CIF1, "T")
        mDomClien = DBLet(RS!domclien, "T")
        
        Mens = "Datos del Cliente"
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




Public Function LeerDatosFactu(codClien As String, codEmpre As String) As Boolean
'Leer los datos de facturacion del cliente para una empresa
'Lee de la BD: PlannerTours, Tabla: cltefact

    Dim RS As ADODB.Recordset
    Dim SQL As String, Mens As String

    On Error GoTo ELeerF
    
    SQL = "SELECT cltefact.*,desforpa FROM cltefact "
    SQL = SQL & " LEFT OUTER JOIN forpagos ON cltefact.codforpa=forpagos.codforpa "
    SQL = SQL & " WHERE codclien=" & codClien & " AND codempre=" & codEmpre

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If RS.EOF Then
        LeerDatosFactu = False
    Else
        mCodForpa = RS!codforpa
        mDesCodForpa = DBLet(RS!desforpa, "T")
        
        Mens = "Datos de Facturación del Cliente"
        LeerDatosFactu = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeerF:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set RS = Nothing
    LeerDatosFactu = False
End Function


