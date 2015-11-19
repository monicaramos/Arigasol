VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ### [Monica] 04/10/2006
' clase modificada para la tabla de ssocios de arigasol


Option Explicit

'Socios
'TABLA: ssocio


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodsocio As String 'Cod. socio
Private mCodcoope As String 'Cod. colectivo
Private mNomsocio As String 'Nombre socio
Private mNomComer As String 'Nombre comercial socio
Private mDomsocio As String 'Domicilio socio
Private mCodpobla As String 'cod. postal socio
Private mPobsocio As String 'poblacion socio
Private mProsocio As String 'Provincia socio
Private mNIFsocio As String 'NIF socio
Private mTfnsocio As String 'Telefono administracion
Private mMaisocio As String 'Mail administracion

Private mCodEnvio As String 'cod. envio
Private mCodAgent As String 'cod. agente
Private mCodForpa As String 'cod. forma de pago
Private mDtoLitro As String 'descuento por litro

'[Monica]22/11/2013: tema iban
Private mIban As String ' iban

Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria
Private mCuentaConta As String 'cuenta contable

Private mCodSitua As Byte 'codigo de situacion

Private mWwwsocio As String 'web del socio

Private mDiaPago As String 'Dia de pago



'------------------------------------------------
'Propiedades del modulo Csocio
'------------------------------------------------

'**** Codigo del socio

Public Property Let Codigo(ByVal vData As String)
     mCodsocio = vData
End Property

Public Property Get Codigo() As String
     Codigo = mCodsocio
End Property

'**** Colectivo del socio

Public Property Let Colectivo(ByVal vData As String)
     mCodcoope = vData
End Property

Public Property Get Colectivo() As String
     Colectivo = mCodcoope
End Property


'**** Nombre del socio

Public Property Let Nombre(ByVal vData As String)
     mNomsocio = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomsocio
End Property


'**** Domicilio del socio

Public Property Let Domicilio(ByVal vData As String)
     mDomsocio = vData
End Property

Public Property Get Domicilio() As String
     Domicilio = mDomsocio
End Property

'**** CPostal del socio

Public Property Let CPostal(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodpobla
End Property

'**** Poblacion del socio

Public Property Let Poblacion(ByVal vData As String)
     mPobsocio = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobsocio
End Property

'**** Provincia del socio

Public Property Let Provincia(ByVal vData As String)
     mProsocio = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProsocio
End Property

'**** NIF del socio

Public Property Let NIF(ByVal vData As String)
     mNIFsocio = vData
End Property

Public Property Get NIF() As String
     NIF = mNIFsocio
End Property

'**** Telefono administracion del socio

Public Property Let Tfnosocio(ByVal vData As String)
     mTfnsocio = vData
End Property

Public Property Get Tfnosocio() As String
     Tfnosocio = mTfnsocio
End Property

'**** e-MAil administracion del socio

Public Property Let EMailAdm(ByVal vData As String)
     mMaisocio = vData
End Property

Public Property Get EMailAdm() As String
     EMailAdm = mMaisocio
End Property


'**** Web del socio

Public Property Let Websocio(ByVal vData As String)
     mWwwsocio = vData
End Property

Public Property Get Websocio() As String
     Websocio = mWwwsocio
End Property



'**** Forma de pago del socio

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property


'**** Descuento por litro

Public Property Let DtoLitro(ByVal vData As String)
     mDtoLitro = vData
End Property

Public Property Get DtoLitro() As String
     DtoLitro = mDtoLitro
End Property

'[Monica]22/11/2013: tema Iban
'**** IBAN
Public Property Let Iban(ByVal vData As String)
     mIban = vData
End Property

Public Property Get Iban() As String
     Iban = mIban
End Property


'**** Banco del socio

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal del socio

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property


'**** Digito Control de la cuenta del socio

Public Property Let Digcontrol(ByVal vData As String)
     mDigContr = vData
End Property

Public Property Get Digcontrol() As String
     Digcontrol = mDigContr
End Property


'**** Cuenta Bancaria de la Factura

Public Property Let CuentaBan(ByVal vData As String)
     mCuentaba = vData
End Property

Public Property Get CuentaBan() As String
     CuentaBan = mCuentaba
End Property


'**** Cuenta Contable

Public Property Let CuentaConta(ByVal vData As String)
     mCuentaConta = vData
End Property

Public Property Get CuentaConta() As String
     CuentaConta = mCuentaConta
End Property


'**** situacion socio

Public Property Let Situacion(ByVal vData As Byte)
     mCodSitua = vData
End Property

Public Property Get Situacion() As Byte
     Situacion = mCodSitua
End Property


'**** Dia de pago

Public Property Let DiaPago(ByVal vData As String)
     mDiaPago = vData
End Property

Public Property Get DiaPago() As String
     DiaPago = mDiaPago
End Property


'------------------------------------------------
'Procedimientos del modulo Csocio
'------------------------------------------------

Public Function Existe(vcodsocio As String) As Boolean
'Comprueba si existe el socio en la BD
Dim devuelve As String

    On Error GoTo EExiste
    
    devuelve = DevuelveDesdeBD("codsocio", "ssocio", "codsocio", vcodsocio, "N")
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe el socio " & vcodsocio & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function

Public Function EsDeVarios(vcodsocio As String) As Boolean
'Comprueba si existe el socio en la BD
Dim devuelve As String

    On Error Resume Next
    
    devuelve = DevuelveDesdeBD("esdevarios", "ssocio", "codsocio", vcodsocio, "N")
    EsDeVarios = (devuelve = "1")
    

End Function

Public Function LeerDatos(vcodsocio As String) As Boolean
'Leer los datos de un socio dado
'Lee de la BD: Ariges, Tabla: ssocio
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    
    LeerDatos = False
    
    SQL = "SELECT codsocio,codcoope,nomsocio,domsocio,codposta,pobsocio,prosocio,nifsocio,wwwsocio,"
    SQL = SQL & "codforpa,codbanco,codsucur,digcontr,cuentaba,telsocio,maisocio,dtolitro,codsitua,"
    SQL = SQL & "fechaalt,fechabaj, codmacta, diapago "
    
    '[Monica]22/11/2013: tema iban
    SQL = SQL & ", iban  "
    
    SQL = SQL & " FROM ssocio "
    SQL = SQL & " WHERE codsocio=" & vcodsocio
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatos = False
    Else
        mCodsocio = CStr(RS!codsocio)
        
        mNomsocio = RS!NomSocio
        mCodcoope = RS!codcoope
        mDomsocio = RS!domsocio
        mCodpobla = RS!codPosta
        mPobsocio = RS!pobsocio
        mProsocio = DBLet(RS!prosocio, "T")
        mNIFsocio = DBLet(RS!nifsocio, "T")
        mWwwsocio = DBLet(RS!wwwsocio, "T")
        mTfnsocio = DBLet(RS!telsocio, "T")
        mMaisocio = DBLet(RS!maisocio, "T")
       
        mCodForpa = CStr(RS!Codforpa)
        mDtoLitro = CStr(RS!DtoLitro)
        
        '[Monica]22/11/2013: tema iban
        mIban = CStr(DBLet(RS!Iban, "T"))
        
        mCodbanco = CStr(DBLet(RS!codbanco, "N"))
        mCodsucur = CStr(DBLet(RS!codsucur, "N"))
        mDigContr = DBLet(RS!digcontr, "T")
        mCuentaba = DBLet(RS!cuentaba, "T")
        mCuentaConta = DBLet(RS!Codmacta, "T")
        
        mCodSitua = RS!codsitua
        
        mDiaPago = CStr(DBLet(RS!DiaPago, "T"))
        
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos socio" & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function



Public Function ActualizarsocioV(vNIF As String) As Boolean
'Actualiza los datos del socio de varios en la tabla: sclvar
Dim SQL As String
Dim devuelve As String

    On Error GoTo EActualizarsocio
    ActualizarsocioV = False
    
    devuelve = DevuelveDesdeBDNew("ssocior", "nifsocio", "nifsocio", vNIF, "T")
    
    If devuelve <> "" Then 'Existe el socio de varios y lo actualizamos
        SQL = "UPDATE ssocio SET nomsocio=" & DBSet(Nombre, "T") & ", domsocio=" & DBSet(Domicilio, "T") & ", "
        SQL = SQL & " codpobla=" & DBSet(CPostal, "T") & ", pobsocio=" & DBSet(Poblacion, "T") & ", "
        SQL = SQL & " prosocio=" & DBSet(Provincia, "T") & ", telsocio=" & DBSet(Tfnosocio, "T")
        SQL = SQL & " WHERE nifsocio=" & DBSet(vNIF, "T")
    Else 'No existe el socio y lo añadimos
        SQL = "INSERT INTO ssocio (nifsocio, nomsocio, domsocio, codpobla, pobsocio, prosocio, telsocio) "
        SQL = SQL & " VALUES (" & DBSet(vNIF, "T") & ", " & DBSet(Nombre, "T") & ", " & DBSet(Domicilio, "T") & ", "
        SQL = SQL & DBSet(CPostal, "T") & ", " & DBSet(Poblacion, "T") & ", " & DBSet(Provincia, "T") & ", " & DBSet(Tfnosocio, "T") & ")"
    End If
    Conn.Execute SQL, , adCmdText
    
    ActualizarsocioV = True
    
EActualizarsocio:
    If Err.Number <> 0 Then
        ActualizarsocioV = False
        MuestraError Err.Number, "Actualizando socio de Varios.", Err.Description
    End If
End Function


Public Function SocioBloqueado() As Boolean
'devuelve true si el socio esta bloqueado
'si la situación del socio es distinta de NORMAL(codsitua=0) entonces
'mostrar un mensaje con la situación especial del socio
Dim tipo As String
Dim devuelve As String

    On Error GoTo EBloqueado
    SocioBloqueado = False
    
    If Situacion <> 0 Then '0: situacion normal
        tipo = "tipositu"
        devuelve = DevuelveDesdeBD("ssitua", "nomsitua", "codsitua", Situacion, "N", tipo)
        
        If tipo = "1" Then 'socio Bloqueado por Situación Especial.
            MsgBox UCase("Socio Bloqueado por: ") & vbCrLf & devuelve, vbInformation, "Situación Especial del Socio."
            SocioBloqueado = True
        Else
            MsgBox devuelve, vbInformation, "Situación Especial del Socio."
        End If
    End If
    
EBloqueado:
    If Err.Number <> 0 Then Err.Clear
End Function

'**** tipo de contabilizacion del colectivo
' 0 = cta contable socio
' 1 = cta contable cliente
Public Function TipoConta() As Byte
' nos devuelve el tipo de contabilizacion que utiliza el colectivo al que pertenece el socio
Dim devuelve As String

    '[Monica]25/09/2014: el tipo de contabilizacion pasa a estar en el socio en lugar de en el colectivo
    devuelve = DevuelveDesdeBD("tipconta", "ssocio", "codsocio", Codigo, "N")
    TipoConta = DBLet(devuelve, "N")

End Function

'**** tipo de facturacion del colectivo
' 0 = tarjeta
' 1 = cliente
Public Function TipoFactu() As Byte
' nos devuelve el tipo de contabilizacion que utiliza el colectivo al que pertenece el socio
Dim devuelve As String
    
    devuelve = DevuelveDesdeBD("tipfactu", "scoope", "codcoope", Colectivo, "N")
    TipoFactu = DBLet(devuelve, "N")

End Function


