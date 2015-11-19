VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TABLA: sprove
'PROVEEDORES
'Variables locales que contienen valores de propiedad
Private mCodProve As String 'Cod. proveedor
Private mNomProve As String 'Nombre proveedor
Private mNomComer As String 'Nombre comercial proveedor
Private mDomProve As String 'Domicilio proveedor
Private mCodpobla As String 'cod. postal proveedor
Private mPobprove As String 'poblacion proveedor
Private mProprove As String 'Provincia proveedor
Private mNifProve As String 'NIF proveedor
Private mTipProve As Byte   'Tipo de proveedor: 0=Nacional, 1=Intracomunitario, 2=Extranjero
Private mTfn1Prov As String 'Telefono
Private mMaiProv1 As String 'EMail de administracion
Private mWwwProve As String 'Web proveedor

Private mCodForpa As String 'cod. forma de pago
Private mDtoPpago As String 'descuento pronto pago
Private mDtoGnral As String 'descuento general
Private mTipoDtos As Byte   'Tipo descuentos: 0=aditivo, 1=Sobre resto

'[Monica]22/11/2013: tema iban
Private mIban As String ' iban

Private mBancoPr As String  'cod. banco propio
Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria

Private mCodmacta As String 'cta contable del proveedor

Private mObservac As String 'Observaciones
Private mProVario As Boolean 'si es proveedor de varios


'------------------------------------------------
'------------------------------------------------
'Propiedades del modulo CProveedor
'------------------------------------------------
'------------------------------------------------


'**** Codigo del proveedor

Public Property Let Codigo(ByVal vData As String)
     mCodProve = vData
End Property

Public Property Get Codigo() As String
     Codigo = mCodProve
End Property


'**** Nombre del proveedor

Public Property Let Nombre(ByVal vData As String)
     mNomProve = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomProve
End Property

'**** Domicilio del proveedor

Public Property Let Domicilio(ByVal vData As String)
     mDomProve = vData
End Property

Public Property Get Domicilio() As String
     Domicilio = mDomProve
End Property

'**** CPostal del proveedor

Public Property Let CPostal(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodpobla
End Property

'**** Poblacion del proveedor

Public Property Let Poblacion(ByVal vData As String)
     mPobprove = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobprove
End Property

'**** Provincia del proveedor

Public Property Let Provincia(ByVal vData As String)
     mProprove = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProprove
End Property

'**** NIF del proveedor

Public Property Let NIF(ByVal vData As String)
     mNifProve = vData
End Property

Public Property Get NIF() As String
     NIF = mNifProve
End Property

'**** Tipo de proveedor

Public Property Let TipoProv(ByVal vData As Byte)
     mTipProve = vData
End Property

Public Property Get TipoProv() As Byte
     TipoProv = mTipProve
End Property

'**** Telefono administracion del proveedor

Public Property Let TfnoAdmon(ByVal vData As String)
     mTfn1Prov = vData
End Property

Public Property Get TfnoAdmon() As String
     TfnoAdmon = mTfn1Prov
End Property



'**** EMail administracion del proveedor

Public Property Let EMailAdmon(ByVal vData As String)
     mMaiProv1 = vData
End Property

Public Property Get EMailAdmon() As String
     EMailAdmon = mMaiProv1
End Property



'**** Web del proveedor

Public Property Let WebProve(ByVal vData As String)
     mWwwProve = vData
End Property

Public Property Get WebProve() As String
     WebProve = mWwwProve
End Property




'**** Forma de pago del proveedor

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property

'**** Descuento pronto pago del proveedor

Public Property Let DtoPPago(ByVal vData As String)
     mDtoPpago = vData
End Property

Public Property Get DtoPPago() As String
     DtoPPago = mDtoPpago
End Property

'**** Descuento general del proveedor

Public Property Let DtoGnral(ByVal vData As String)
     mDtoGnral = vData
End Property

Public Property Get DtoGnral() As String
     DtoGnral = mDtoGnral
End Property

'**** Tipo Descuento del proveedor

Public Property Let TipoDto(ByVal vData As Byte)
     mTipoDtos = vData
End Property

Public Property Get TipoDto() As Byte
     TipoDto = mTipoDtos
End Property

'[Monica]22/11/2013: tema Iban
'**** IBAN
Public Property Let Iban(ByVal vData As String)
     mIban = vData
End Property

Public Property Get Iban() As String
     Iban = mIban
End Property


'**** Banco propio del proveedor

Public Property Let BancoPropio(ByVal vData As String)
     mBancoPr = vData
End Property

Public Property Get BancoPropio() As String
     BancoPropio = mBancoPr
End Property

'**** Banco del proveedor

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal del proveedor

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property

'**** Digito Control de la cuenta del proveedor
Public Property Let DigControl(ByVal vData As String)
     mDigContr = vData
End Property

Public Property Get DigControl() As String
     DigControl = mDigContr
End Property

'**** Cuenta Bancaria del proveedor

Public Property Let CuentaBan(ByVal vData As String)
     mCuentaba = vData
End Property

Public Property Get CuentaBan() As String
     CuentaBan = mCuentaba
End Property


'**** Cuenta contable del proveedor

Public Property Let CuentaCble(ByVal vData As String)
     mCodmacta = vData
End Property

Public Property Get CuentaCble() As String
     CuentaCble = mCodmacta
End Property

'**** Observaciones del proveedor

Public Property Let Observaciones(ByVal vData As String)
     mObservac = vData
End Property

Public Property Get Observaciones() As String
     Observaciones = mObservac
End Property

'**** proveedor de varios (SI/NO)

Public Property Let DeVarios(ByVal vData As Boolean)
     mProVario = vData
End Property

Public Property Get DeVarios() As Boolean
     DeVarios = mProVario
End Property



'------------------------------------------------
'------------------------------------------------
'Procedimientos del modulo CProveedor
'------------------------------------------------
'------------------------------------------------

Public Function Existe(vCodProve As String) As Boolean
'Comprueba si existe el proveedor en la BD
Dim devuelve As String

    On Error GoTo EExiste
    
    devuelve = DevuelveDesdeBDNew(cPTours, "proveedor", "codprove", "codprove", vCodProve, "N")
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe el Proveedor " & vCodProve & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function


Public Function LeerDatos(vCodProve As String) As Boolean
'Leer los datos de un proveedor dado
'Lee de la BD: Ariges, Tabla: sprove
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim Mens As String
Dim SQL As String

    On Error GoTo ELeer
    LeerDatos = False
    
    SQL = "SELECT codprove,nomprove,nomcomer,codpobla,domprove,pobprove,proprove,tipprove,nifprove,telprov1,maiprov1,wwwprove,provario,codforpa,tipodtos,dtoppago,dtognral,codbanpr,codbanco,codsucur,digcontr,cuentaba,codmacta,observac"
    
    '[Monica]22/11/2013: tema iban
    SQL = SQL & ", iban  "
   
    SQL = SQL & " FROM proveedor "
    SQL = SQL & " WHERE codprove=" & vCodProve
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If RS.EOF Then
        LeerDatos = False
    Else
        Mens = "Datos Proveedor"
        mCodProve = CStr(RS!CodProve)
        
        mNomProve = RS!NomProve
        mNomComer = RS!NomComer
        mDomProve = DBLet(RS!domprove, "T")
        mCodpobla = RS!codPobla
        mPobprove = RS!pobprove
        mProprove = RS!proprove
        mNifProve = RS!NIFProve
        mTfn1Prov = DBLet(RS!telprov1, "T")
        mMaiProv1 = DBLet(RS!maiprov1, "T")
        mWwwProve = DBLet(RS!wwwprove, "T")
        
        mTipProve = RS!tipprove
       
        mCodForpa = CStr(RS!Codforpa)
        mDtoPpago = CStr(DBLet(RS!DtoPPago, "N"))
        mDtoGnral = CStr(DBLet(RS!DtoGnral, "N"))
        mTipoDtos = RS!TipoDtos
        
        '[Monica]22/11/2013: tema iban
        mIban = CStr(DBLet(RS!Iban, "T"))
        
        mBancoPr = CStr(RS!codBanpr)
        mCodbanco = CStr(DBLet(RS!codbanco, "N"))
        mCodsucur = CStr(DBLet(RS!codsucur, "N"))
        mDigContr = DBLet(RS!digcontr, "T")
        mCuentaba = DBLet(RS!cuentaba, "T")
        
        mCodmacta = DBLet(RS!Codmacta, "T")
        
        mProVario = RS!provario
        mObservac = DBLet(RS!observac)
        
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


Public Function LeerDatosProveVario(vNIF As String) As Boolean
Dim RS As ADODB.Recordset
Dim Mens As String
Dim SQL As String

    'si es proveedor de varios recuperamos los datos nombre,etc.
    'de la tabla de varios sprvar
    
    On Error GoTo ELeer
    LeerDatosProveVario = False
    
    SQL = "SELECT nifprove,nomprove,domprove,codpobla,pobprove,proprove,telprov2 "
    SQL = SQL & " FROM sprvar "
    SQL = SQL & " WHERE nifprove='" & vNIF & "'"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If RS.EOF Then
        LeerDatosProveVario = False
        MsgBox "No existe el Proveedor de varios: " & vNIF, vbInformation
    Else
        Mens = "Datos Proveedor de Varios"
        
        mNifProve = RS!NIFProve
        mNomProve = RS!NomProve
        mNomComer = ""
        mDomProve = RS!domprove
        mCodpobla = DBLet(RS!codPobla, "T")
        mPobprove = DBLet(RS!pobprove, "T")
        mProprove = DBLet(RS!proprove, "T")
        mTfn1Prov = DBLet(RS!telprov2, "T")
        
        LeerDatosProveVario = True
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
    LeerDatosProveVario = False
End Function



Public Function ActualizarProveV(vNIF As String) As Boolean
'Actualizar los datos de un Proveedor de varios
On Error GoTo EActualizarProve
Dim SQL As String
Dim devuelve As String

    ActualizarProveV = False
    devuelve = DevuelveDesdeBDNew(cPTours, "sprvar", "nifprove", "nifprove", vNIF, "T")
    
    If devuelve <> "" Then 'Existe el proveedor de varios y lo actualizamos
        SQL = "UPDATE sprvar SET nomprove=" & DBSet(Nombre, "T") & ", domprove=" & DBSet(Domicilio, "T") & ", "
        SQL = SQL & " codpobla=" & DBSet(CPostal, "T") & ", pobprove=" & DBSet(Poblacion, "T") & ", "
        SQL = SQL & " proprove=" & DBSet(Provincia, "T") & ", telprov2=" & DBSet(TfnoAdmon, "T")
        SQL = SQL & " WHERE nifprove=" & DBSet(vNIF, "T")
    Else 'No existe el proveedor y lo añadimos
        SQL = "INSERT INTO sprvar (nifprove, nomprove, domprove, codpobla, pobprove, proprove, telprov2) "
        SQL = SQL & " VALUES (" & DBSet(vNIF, "T") & ", " & DBSet(Nombre, "T") & ", " & DBSet(Domicilio, "T") & ", "
        SQL = SQL & DBSet(CPostal, "T") & ", " & DBSet(Poblacion, "T") & ", " & DBSet(Provincia, "T") & ", " & DBSet(TfnoAdmon, "T") & ")"
    End If
    Conn.Execute SQL, , adCmdText
    
    ActualizarProveV = True
    
EActualizarProve:
    If Err.Number <> 0 Then
        ActualizarProveV = False
        MuestraError Err.Number, "Actualizando Proveedor de Varios", Err.Description
    End If
End Function




Public Function ActualizaFechaUltCompra(codprov As String, fechacom As String) As Boolean
Dim cad As String
Dim b As Boolean

    On Error GoTo EActFecha
    
    'Actualizar el campo fecha ult.compra(fechamov) en la tabla proveedores (sprove)
    cad = DevuelveDesdeBDNew(cPTours, "proveedor", "fechamov", "codprove", codprov, "N")
    If (cad = "") Then cad = "01/01/1900"
    
    If CDate(fechacom) > CDate(cad) Then
        cad = "UPDATE proveedor SET fechamov=" & DBSet(fechacom, "F")
        cad = cad & " WHERE codprove=" & codprov
        Conn.Execute cad, , adCmdText
    End If
    b = True
    
EActFecha:
    If Err.Number <> 0 Then
        b = False
'        MuestraError Err.Number, "Actualizando fecha ult. compra en proveedor.", Err.Description
    End If
    ActualizaFechaUltCompra = b
End Function
    
