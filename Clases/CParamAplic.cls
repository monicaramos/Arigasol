VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros SPARAM

'Datos de la Contabilidad
Private mvarSerConta As String 'Servidor BD de Contabilidad
Private mvarUsuConta As String 'Usuario BD de Contabilidad
Private mvarPasConta As String 'Password BD de Contabilidad
Private mvarNumConta As Byte 'Numero BD de Contabilidad
Private mvarCtaConta As String 'Cuenta Contable
Private mvarCtaNegtat As String 'Cuenta Diferencias negativas
Private mvarCtaPosit As String 'Cuenta Diferencias positivas
Private mvarCtaImpue As String 'Cuenta Contable Impuesto
Private mvarTeximpue As String  'Texto Impuesto
Private mvarBonifact As Byte 'Bonificaciones
Private mvarArticDto As Long 'Articulo de descuento
Private mvarCooperativa As Integer 'Cooperativa

Private mvarRaizCtaSoc As String 'Raiz de Cuenta Contable de socio
Private mvarRaizCtaCli As String 'Raiz de Cuenta Contable de cliente
Private mvarCtaFamDefecto As String 'Cuenta Contable por defecto para las familia

Private mvarWebSoporte As String  'Web Soporte
Private mvarDireMail As String 'direccion mail
Private mvarSmtphost As String 'Host
Private mvarSmtpUser As String 'User
Private mvarSmtpPass As String 'Password

' datos de declaracion de gasoleo profesional
Private mvarCim As String 'Cim
Private mvarCee As String 'Cee

'diario y conceptos para contabilizacion
Private mvarConceDebe As Byte 'concepto al debe
Private mvarConceHaber As Byte 'concepto al haber
Private mvarConceHaberResto As Byte 'concepto al haber resto
Private mvarNumDiario  As Byte 'numero de diario

Private mvarColecDefecto As Integer 'Colectivo defecto <> 0 solo para regaixo
Private mvarFamiliaDescuento As Integer 'Familia descuento solo para alzira

Private mvarImpresoraTarjetas As String  'impresora de tarjetas
Private mvarImpresoraTickets As String  'impresora de tickets

Private mvarDiaPago1 As Byte 'Dia de pago 1 para compras
Private mvarDiaPago2 As Byte 'Dia de pago 2 para compras
Private mvarDiaPago3 As Byte 'Dia de pago 3 para compras
Private mvarMesnogir As Byte 'Mes a no girar para compras
Private mvarCtrStock As Boolean   'Si hay Control de Stock (S/N)
Private mvarTipoPrecio As Byte 'Tipo precio

' facturacion interna
'diario y conceptos para contabilizacion
Private mvarTipoIvaExento As Byte 'tipo de iva exento
Private mvarLetraInt As String 'Letra de facturas Internas

Private mvarEnvioDesdeOutlook As Boolean  'Si envia por outlook
Private mvarExeEnvioMail As String

Private mvarPathFacturaE As String  'FacturaE


Private mvarHayCRM As Byte 'hay CRM

Private mvarFechaCam As String 'Fecha de cambio para el calculo de impuesto


Private mvarLimiteFra As Currency ' Limite factura de efectivo


Private mvarCodIvaGnral As Integer 'Codigo de iva general
Private mvarCodIvaRedu As Integer 'Codigo de iva reducido
Private mvarCodIvaSRedu As Integer 'Codigo de iva superreducido
Private mvarCodIvaGnralAnt As Integer 'Codigo de iva general antiguo
Private mvarCodIvaReduAnt As Integer 'Codigo de iva reducido antiguo

Private mvarFechaCamIva As String ' fecha de cambio de iva 01/09/2012

Private mvarContabilidadNueva As Boolean
Private mvarAgenteAriconta As Integer 'Agente que se grabará en la contabilidad


Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes


'------------------------------------------------
'Propiedades del modulo CParamAplic
'------------------------------------------------


'servidor conta
Public Property Let ServidorConta(ByVal vData As String)
     mvarSerConta = vData
End Property

Public Property Get ServidorConta() As String
     ServidorConta = mvarSerConta
End Property

Public Property Let UsuarioConta(ByVal vData As String)
     mvarUsuConta = vData
End Property

Public Property Get UsuarioConta() As String
     UsuarioConta = mvarUsuConta
End Property

Public Property Let PasswordConta(ByVal vData As String)
     mvarPasConta = vData
End Property

Public Property Get PasswordConta() As String
     PasswordConta = mvarPasConta
End Property

Public Property Let NumeroConta(ByVal vData As Byte)
     mvarNumConta = vData
End Property

Public Property Get NumeroConta() As Byte
     NumeroConta = mvarNumConta
End Property

'******
Public Property Let CtaContable(ByVal vData As String)
     mvarCtaConta = vData
End Property

Public Property Get CtaContable() As String
     CtaContable = mvarCtaConta
End Property

Public Property Let CtaNegativa(ByVal vData As String)
     mvarCtaNegtat = vData
End Property

Public Property Get CtaNegativa() As String
     CtaNegativa = mvarCtaNegtat
End Property

Public Property Let CtaPositiva(ByVal vData As String)
     mvarCtaPosit = vData
End Property

Public Property Get CtaPositiva() As String
     CtaPositiva = mvarCtaPosit
End Property

Public Property Let CtaImpuesto(ByVal vData As String)
     mvarCtaImpue = vData
End Property

Public Property Get CtaImpuesto() As String
     CtaImpuesto = mvarCtaImpue
End Property

Public Property Let TextoImpuesto(ByVal vData As String)
     mvarTeximpue = vData
End Property

Public Property Get TextoImpuesto() As String
     TextoImpuesto = mvarTeximpue
End Property

Public Property Let Bonifact(ByVal vData As Byte)
     mvarBonifact = vData
End Property

Public Property Get Bonifact() As Byte
     Bonifact = mvarBonifact
End Property

Public Property Let ArticDto(ByVal vData As Long)
     mvarArticDto = vData
End Property

Public Property Get ArticDto() As Long
     ArticDto = mvarArticDto
End Property

Public Property Let RaizCtaSoc(ByVal vData As String)
     mvarRaizCtaSoc = vData
End Property

Public Property Get RaizCtaSoc() As String
     RaizCtaSoc = mvarRaizCtaSoc
End Property

Public Property Let RaizCtaCli(ByVal vData As String)
     mvarRaizCtaCli = vData
End Property

Public Property Get RaizCtaCli() As String
     RaizCtaCli = mvarRaizCtaCli
End Property

Public Property Let CtaFamDefecto(ByVal vData As String)
     mvarCtaFamDefecto = vData
End Property

Public Property Get CtaFamDefecto() As String
     CtaFamDefecto = mvarCtaFamDefecto
End Property

Public Property Let WebSoporte(ByVal vData As String)
     mvarWebSoporte = vData
End Property

Public Property Get WebSoporte() As String
     WebSoporte = mvarWebSoporte
End Property

Public Property Let Cooperativa(ByVal vData As Integer)
     mvarCooperativa = vData
End Property

Public Property Get Cooperativa() As Integer
     Cooperativa = mvarCooperativa
End Property


Public Property Let DireMail(ByVal vData As String)
     mvarDireMail = vData
End Property

Public Property Get DireMail() As String
     DireMail = mvarDireMail
End Property

Public Property Let Smtphost(ByVal vData As String)
     mvarSmtphost = vData
End Property

Public Property Get Smtphost() As String
     Smtphost = mvarSmtphost
End Property

Public Property Let SmtpUser(ByVal vData As String)
     mvarSmtpUser = vData
End Property

Public Property Get SmtpUser() As String
     SmtpUser = mvarSmtpUser
End Property

Public Property Let Smtppass(ByVal vData As String)
     mvarSmtpPass = vData
End Property

Public Property Get Smtppass() As String
     Smtppass = mvarSmtpPass
End Property


Public Property Let Cim(ByVal vData As String)
     mvarCim = vData
End Property

Public Property Get Cim() As String
     Cim = mvarCim
End Property

Public Property Let Cee(ByVal vData As String)
     mvarCee = vData
End Property

Public Property Get Cee() As String
     Cee = mvarCee
End Property


Public Property Let ConceptoDebe(ByVal vData As Byte)
     mvarConceDebe = vData
End Property

Public Property Get ConceptoDebe() As Byte
     ConceptoDebe = mvarConceDebe
End Property

Public Property Let ConceptoHaber(ByVal vData As Byte)
     mvarConceHaber = vData
End Property

Public Property Get ConceptoHaber() As Byte
     ConceptoHaber = mvarConceHaber
End Property

Public Property Let ConceptoHaberResto(ByVal vData As Byte)
     mvarConceHaberResto = vData
End Property

Public Property Get ConceptoHaberResto() As Byte
     ConceptoHaberResto = mvarConceHaberResto
End Property

Public Property Let NumDiario(ByVal vData As Byte)
     mvarNumDiario = vData
End Property

Public Property Get NumDiario() As Byte
     NumDiario = mvarNumDiario
End Property

Public Property Let ColecDefecto(ByVal vData As Integer)
     mvarColecDefecto = vData
End Property

Public Property Get ColecDefecto() As Integer
     ColecDefecto = mvarColecDefecto
End Property

Public Property Let ImpresoraTarjetas(ByVal vData As String)
     mvarImpresoraTarjetas = vData
End Property

Public Property Get ImpresoraTarjetas() As String
     ImpresoraTarjetas = mvarImpresoraTarjetas
End Property

Public Property Let ImpresoraTickets(ByVal vData As String)
     mvarImpresoraTickets = vData
End Property

Public Property Get ImpresoraTickets() As String
     ImpresoraTickets = mvarImpresoraTickets
End Property

Public Property Let FamDto(ByVal vData As Integer)
     mvarFamiliaDescuento = vData
End Property

Public Property Get FamDto() As Integer
     FamDto = mvarFamiliaDescuento
End Property

Public Property Let DiaPago1(ByVal vData As Byte)
     mvarDiaPago1 = vData
End Property

Public Property Get DiaPago1() As Byte
     DiaPago1 = mvarDiaPago1
End Property

Public Property Let DiaPago2(ByVal vData As Byte)
     mvarDiaPago2 = vData
End Property

Public Property Get DiaPago2() As Byte
     DiaPago2 = mvarDiaPago2
End Property

Public Property Let DiaPago3(ByVal vData As Byte)
     mvarDiaPago3 = vData
End Property

Public Property Get DiaPago3() As Byte
     DiaPago3 = mvarDiaPago3
End Property


Public Property Let MesNoGirar(ByVal vData As Byte)
     mvarMesnogir = vData
End Property

Public Property Get MesNoGirar() As Byte
     MesNoGirar = mvarMesnogir
End Property

Public Property Let ControlStock(ByVal vData As Boolean)
     mvarCtrStock = vData
End Property

Public Property Get ControlStock() As Boolean
     ControlStock = mvarCtrStock
End Property


Public Property Let TipoPrecio(ByVal vData As Byte)
     mvarTipoPrecio = vData
End Property

Public Property Get TipoPrecio() As Byte
     TipoPrecio = mvarTipoPrecio
End Property



' FACTURACION INTERNA
' codigo de iva exento

Public Property Let TipoIvaExento(ByVal vData As String)
     mvarTipoIvaExento = vData
End Property

Public Property Get TipoIvaExento() As String
     TipoIvaExento = mvarTipoIvaExento
End Property


Public Property Let LetraInt(ByVal vData As String)
     mvarLetraInt = vData
End Property

Public Property Get LetraInt() As String
     LetraInt = mvarLetraInt
End Property

Public Property Get EnvioDesdeOutlook() As Boolean
     EnvioDesdeOutlook = mvarEnvioDesdeOutlook
End Property

Public Property Let EnvioDesdeOutlook(ByVal vData As Boolean)
     mvarEnvioDesdeOutlook = vData
End Property

Public Property Get ExeEnvioMail() As String
     ExeEnvioMail = mvarExeEnvioMail
End Property

Public Property Let ExeEnvioMail(ByVal vData As String)
     mvarExeEnvioMail = vData
End Property

'Facturacion Electrónica FacturaE
Public Property Get PathFacturaE() As String
     PathFacturaE = mvarPathFacturaE
End Property

Public Property Let PathFacturaE(ByVal vData As String)
     mvarPathFacturaE = vData
End Property


Public Property Let HayCRM(ByVal vData As Byte)
     mvarHayCRM = vData
End Property

Public Property Get HayCRM() As Byte
     HayCRM = mvarHayCRM
End Property

'Fecha de cambio para el calculo de impuestos
Public Property Get FechaCam() As String
     FechaCam = mvarFechaCam
End Property

Public Property Let FechaCam(ByVal vData As String)
     mvarFechaCam = vData
End Property

' Importe Limite de factura de efectivo
Public Property Let LimiteFra(ByVal vData As Currency)
     mvarLimiteFra = vData
End Property

Public Property Get LimiteFra() As Currency
     LimiteFra = mvarLimiteFra
End Property

'**************************************
'****** Tipos de iva nuevos y antiguos

Public Property Let CodIvaGnral(ByVal vData As Integer)
     mvarCodIvaGnral = vData
End Property

Public Property Get CodIvaGnral() As Integer
     CodIvaGnral = mvarCodIvaGnral
End Property

Public Property Let CodIvaRedu(ByVal vData As Integer)
     mvarCodIvaRedu = vData
End Property

Public Property Get CodIvaRedu() As Integer
     CodIvaRedu = mvarCodIvaRedu
End Property

Public Property Let CodIvaSRedu(ByVal vData As Integer)
     mvarCodIvaSRedu = vData
End Property

Public Property Get CodIvaSRedu() As Integer
     CodIvaSRedu = mvarCodIvaSRedu
End Property

Public Property Let CodIvaGnralAnt(ByVal vData As Integer)
     mvarCodIvaGnralAnt = vData
End Property

Public Property Get CodIvaGnralAnt() As Integer
     CodIvaGnralAnt = mvarCodIvaGnralAnt
End Property

Public Property Let CodIvaReduAnt(ByVal vData As Integer)
     mvarCodIvaReduAnt = vData
End Property

Public Property Get CodIvaReduAnt() As Integer
     CodIvaReduAnt = mvarCodIvaReduAnt
End Property

'Fecha de cambio de iva para el calculo de iva en facturas
Public Property Get FechaCamIva() As String
     FechaCamIva = mvarFechaCamIva
End Property

Public Property Let FechaCamIva(ByVal vData As String)
     mvarFechaCamIva = vData
End Property

'[Monica]13/09/2016: si hay contabilidad nueva
Public Property Get ContabilidadNueva() As Boolean
     ContabilidadNueva = mvarContabilidadNueva
End Property

Public Property Let AgenteAriconta(ByVal vData As Integer)
     mvarAgenteAriconta = vData
End Property

Public Property Get AgenteAriconta() As Integer
     AgenteAriconta = mvarAgenteAriconta
End Property


'******
'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------

'****   LEER  ***
Public Function Leer() As Byte
Dim Rs As ADODB.Recordset
On Error GoTo Err3
    
    SQL = "SELECT * FROM sparam "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic

    If Rs.EOF Then
        Leer = 1
    Else
        
        'Datos Contabilidad
        mvarSerConta = DBLet(Rs!serConta, "T") 'servidor conta
        mvarUsuConta = DBLet(Rs!UsuConta, "T") 'usuario conta
        mvarPasConta = DBLet(Rs!pasconta, "T") 'password
        mvarNumConta = DBLet(Rs!NumConta, "N") 'numero conta
        mvarCtaConta = DBLet(Rs!CtaConta, "T") 'cuenta contable
        mvarCtaNegtat = DBLet(Rs!ctanegtat, "T") 'cuenta dif negativas
        mvarCtaPosit = DBLet(Rs!ctaposit, "T") 'cta dif positivas
        mvarCtaImpue = DBLet(Rs!ctaimpue, "T") 'cta impuestos
        mvarTeximpue = DBLet(Rs!teximpue, "T") 'texto impuesto
        mvarBonifact = DBLet(Rs!Bonifact, "N") 'bonificacion
        mvarArticDto = DBLet(Rs!ArticDto, "N") 'articulo de descuento
        mvarRaizCtaSoc = DBLet(Rs!RaizCtaSoc, "T") ' raiz cta contable socio
        mvarRaizCtaCli = DBLet(Rs!RaizCtaCli, "T") ' raiz cta contable cliente
        mvarCtaFamDefecto = DBLet(Rs!CtaFamDefecto, "T") ' cta contable por defecto para familia
        mvarWebSoporte = DBLet(Rs!WebSoporte, "T") 'web soporte
        mvarCooperativa = DBLet(Rs!Cooperativa, "N") ' cooperativa
        mvarDireMail = DBLet(Rs!DireMail, "T") 'mail
        mvarSmtphost = DBLet(Rs!Smtphost, "T") 'host
        mvarSmtpUser = DBLet(Rs!SmtpUser, "T") 'user
        mvarSmtpPass = DBLet(Rs!Smtppass, "T") 'pass
        mvarCim = DBLet(Rs!Cim, "T") 'cim
        mvarCee = DBLet(Rs!Cee, "T") 'cee
        mvarConceDebe = DBLet(Rs!concedebe, "N")
        mvarConceHaber = DBLet(Rs!concehaber, "N")
        mvarConceHaberResto = DBLet(Rs!concehaberresto, "N")
        mvarNumDiario = DBLet(Rs!NumDiari, "N")
        mvarColecDefecto = DBLet(Rs!coopdefecto, "N")
        mvarFamiliaDescuento = DBLet(Rs!FamDto, "N")
        mvarImpresoraTarjetas = DBLet(Rs!impresoratarjeta, "T")
        mvarImpresoraTickets = DBLet(Rs!impresoraticket, "T")
        
        mvarDiaPago1 = DBLet(Rs!DiaPago1, "N") 'dia pago 1 para compras
        mvarDiaPago2 = DBLet(Rs!DiaPago2, "N") 'dia pago 2 para compras
        mvarDiaPago3 = DBLet(Rs!DiaPago3, "N") 'dia pago 3 para compras
        mvarMesnogir = DBLet(Rs!mesnogir, "N") 'mes a no girar para compras
        mvarCtrStock = Rs!CtrStock  'Si hay Control de Stock (S/N)
        mvarTipoPrecio = DBLet(Rs!TipoPrecio, "N") ' tipo de precio
        
        'facturas interna
        mvarTipoIvaExento = DBLet(Rs!TipoIvaExento, "N")
        mvarLetraInt = DBLet(Rs!LetraInt, "T") 'letra de serie de las facturas internas
        
        Me.EnvioDesdeOutlook = DBLet(Rs!EnvioDesdeOutlook, "N") = 1
        mvarExeEnvioMail = DBLet(Rs!arigesmail, "T")
        
        'FacturaE
        mvarPathFacturaE = DBLet(Rs!PathFacturaE, "T")
        
        '[Monica]15/11/2012: hay crm
        mvarHayCRM = DBLet(Rs!HayCRM, "N")
        
        'Fecha de cambio de para el calculo de impuesto
        mvarFechaCam = DBLet(Rs!FechaCam, "T")
                
        'Limite de factura de efectivo
        mvarLimiteFra = DBLet(Rs!LimiteFra, "N") ' limite fra.de efectivo
        
        'Ivas nuevos e ivas antiguos
        mvarCodIvaGnral = DBLet(Rs!codiva1, "N") 'codigo de iva gnral
        mvarCodIvaRedu = DBLet(Rs!codiva2, "N") 'codigo de iva reducido
        mvarCodIvaSRedu = DBLet(Rs!codiva3, "N") 'codigo de iva superreducido
        mvarCodIvaGnralAnt = DBLet(Rs!codiva1old, "N") 'codigo de iva gnral antiguo
        mvarCodIvaReduAnt = DBLet(Rs!codiva2old, "N") 'codigo de iva reducido antiguo
        
        mvarFechaCamIva = DBLet(Rs!feccambioiva, "F") ' fecha de cambio de iva
        
        mvarContabilidadNueva = DBLet(Rs!NuevaContabilidad, "N") = 1
        mvarAgenteAriconta = DBLet(Rs!AgenteAriconta, "N")
        Leer = 0
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Set Rs = Nothing
    Leer = 1
End Function


Public Function Modificar(Codigo As Byte) As Byte
'Modifica los parametros de la Aplicación
'Modifica la Tabla: sparam, BD: Arigasol
On Error GoTo eModificar

    SQL = "UPDATE sparam SET "
    SQL = SQL & " serconta= '" & mvarSerConta & "'"
    SQL = SQL & ", usuconta= " & DBSet(mvarUsuConta, "T")
    SQL = SQL & ", pasconta= " & DBSet(mvarPasConta, "T")
    SQL = SQL & ", numconta= " & DBSet(mvarNumConta, "N", "S")
    SQL = SQL & ", ctaconta= " & DBSet(mvarCtaConta, "T")
    SQL = SQL & ", ctanegtat= " & DBSet(mvarCtaNegtat, "T")
    SQL = SQL & ", ctaposit= " & DBSet(mvarCtaPosit, "T")
    SQL = SQL & ", ctaimpue= " & DBSet(mvarCtaImpue, "T")
    SQL = SQL & ", teximpue= " & DBSet(mvarTeximpue, "T")
    SQL = SQL & ", bonifact= " & DBSet(mvarBonifact, "N", "N")
    SQL = SQL & ", articdto= " & DBSet(mvarArticDto, "N", "S")
    SQL = SQL & ", raizctasoc= " & DBSet(mvarRaizCtaSoc, "T")
    SQL = SQL & ", raizctacli= " & DBSet(mvarRaizCtaCli, "T")
    SQL = SQL & ", ctafamdefecto= " & DBSet(mvarCtaFamDefecto, "T")
    SQL = SQL & ", websoporte= " & DBSet(mvarWebSoporte, "T")
    SQL = SQL & ", cooperativa= " & DBSet(mvarCooperativa, "T")
    SQL = SQL & ", diremail= " & DBSet(mvarDireMail, "T")
    SQL = SQL & ", smtphost= " & DBSet(mvarSmtphost, "T")
    SQL = SQL & ", smtpuser= " & DBSet(mvarSmtpUser, "T")
    SQL = SQL & ", smtppass= " & DBSet(mvarSmtpPass, "T")
    SQL = SQL & ", cim= " & DBSet(mvarCim, "T")
    SQL = SQL & ", cee= " & DBSet(mvarCee, "T")
    SQL = SQL & ", concedebe= " & DBSet(mvarConceDebe, "N")
    SQL = SQL & ", concehaber= " & DBSet(mvarConceHaber, "N")
    SQL = SQL & ", concehaberresto= " & DBSet(mvarConceHaberResto, "N")
    SQL = SQL & ", numdiari= " & DBSet(mvarNumDiario, "N")
    SQL = SQL & ", coopdefecto= " & DBSet(mvarColecDefecto, "N")
    SQL = SQL & ", famdto= " & DBSet(mvarFamiliaDescuento, "N", "S")
    SQL = SQL & ", impresoratarjeta= " & DBSet(mvarImpresoraTarjetas, "T")
    SQL = SQL & ", impresoraticket= " & DBSet(mvarImpresoraTickets, "T")
    SQL = SQL & ", diapago1= " & DBSet(mvarDiaPago1, "N")
    SQL = SQL & ", diapago2= " & DBSet(mvarDiaPago2, "N")
    SQL = SQL & ", diapago3= " & DBSet(mvarDiaPago3, "N")
    SQL = SQL & ", mesnogir= " & DBSet(mvarMesnogir, "N")
    SQL = SQL & ", ctrstock = " & DBSet(mvarCtrStock, "B")
    SQL = SQL & ", tipoprecio= " & DBSet(mvarTipoPrecio, "N")
    
    ' facturacion interna
    SQL = SQL & ", tipoivaexento= " & DBSet(mvarTipoIvaExento, "N")
    SQL = SQL & ", letraint= " & DBSet(mvarLetraInt, "T")
    
    '[Monica]01/08/2011: arigesmail
    SQL = SQL & ", EnvioDesdeOutlook = " & DBSet(Me.EnvioDesdeOutlook, "B")
    SQL = SQL & ", arigesmail = " & DBSet(ExeEnvioMail, "T", "S")
    
    'FacturaE
    SQL = SQL & ", pathfacturae = " & DBSet(mvarPathFacturaE, "T")

    ' Limite Fra de efectivo
    SQL = SQL & ", limitefra= " & DBSet(mvarLimiteFra, "N")
    
    ' codigos de iva nuevos y antiguos
    SQL = SQL & ", codiva1= " & DBSet(mvarCodIvaGnral, "N")
    SQL = SQL & ", codiva2= " & DBSet(mvarCodIvaRedu, "N")
    SQL = SQL & ", codiva3= " & DBSet(mvarCodIvaSRedu, "N")
    SQL = SQL & ", codiva1old= " & DBSet(mvarCodIvaGnralAnt, "N")
    SQL = SQL & ", codiva2old= " & DBSet(mvarCodIvaReduAnt, "N")
    
    SQL = SQL & ", agenteariconta= " & DBSet(mvarAgenteAriconta, "N")
    
    SQL = SQL & " WHERE codparam =" & Codigo & ";"

    Conn.Execute SQL

    Modificar = 1
    Exit Function

eModificar:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 0
End Function

