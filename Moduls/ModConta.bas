Attribute VB_Name = "ModConta"
Option Explicit

'=============================================================================
'   MODULO PARA ACCEDER A LOS DATOS DE LA CONTABILIDAD
'=============================================================================


'=============================================================================
'==========     CUENTAS
'=============================================================================
'LAURA
Public Function PonerNombreCuenta(ByRef Txt As TextBox, Modo As Byte, Optional clien As String) As String
'Obtener el nombre de una cuenta
Dim DevfrmCCtas As String
Dim cad As String

' ### [Monica] 07/09/2006 añadida la linea siguiente condicion vParamAplic.NumeroConta = 0
' para que no saque nada si no hay contabilidad
    If Not vParamAplic Is Nothing Then
        If vParamAplic.NumeroConta = 0 Then
            PonerNombreCuenta = ""
            Exit Function
        End If
    End If
    If Txt.Text = "" Then
         PonerNombreCuenta = ""
         Exit Function
    End If
    DevfrmCCtas = Txt.Text
    If CuentaCorrectaUltimoNivel(DevfrmCCtas, cad) Then
        ' ### [Monica] 07/09/2006
        If InStr(cad, "No existe la cuenta") > 0 Then
            Txt.Text = DevfrmCCtas
'            If (Modo = 4) And clien <> "" Then 'si insertar antes estaba lo de abajo
            If (Modo = 3 Or Modo = 4) And clien <> "" Then 'si insertar o modificar
                cad = cad & "  ¿Desea crearla?"
                If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                    If InStr(1, Txt.Tag, "ssocio") > 0 Then
                        InsertarCuentaCble DevfrmCCtas, clien
                    ElseIf InStr(1, Txt.Tag, "proveedor") > 0 Then
                        InsertarCuentaCble DevfrmCCtas, "", clien
                    ElseIf InStr(1, Txt.Tag, "sbanco") > 0 Then
                        InsertarCuentaCble DevfrmCCtas, "", "", clien
                    End If
                    PonerNombreCuenta = clien
                End If
            Else
                MsgBox cad, vbExclamation
            End If
        Else
            Txt.Text = DevfrmCCtas
            PonerNombreCuenta = cad
        End If
    Else
        MsgBox cad, vbExclamation
'        Txt.Text = ""
        PonerNombreCuenta = ""
'        PonerFoco Txt
    End If
    DevfrmCCtas = ""

End Function




'DAVID: Cuentas del la Contabilidad
Public Function CuentaCorrectaUltimoNivel(ByRef cuenta As String, ByRef devuelve As String) As Boolean
    'Comprueba si es numerica
    Dim sql As String
    Dim otroCampo As String
    
    CuentaCorrectaUltimoNivel = False
    If cuenta = "" Then
        devuelve = "Cuenta vacia"
        Exit Function
    End If

    If Not IsNumeric(cuenta) Then
        devuelve = "La cuenta debe de ser numérica: " & cuenta
        Exit Function
    End If

    'Rellenamos si procede
    cuenta = RellenaCodigoCuenta(cuenta)

    '==========
    If Not EsCuentaUltimoNivel(cuenta) Then
        devuelve = "No es cuenta de último nivel: " & cuenta
        Exit Function
    End If
    '==================

    otroCampo = "apudirec"
    'BD 2: conexion a BD Conta
    sql = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", cuenta, "T", otroCampo)
    If sql = "" Then
        devuelve = "No existe la cuenta : " & cuenta
        CuentaCorrectaUltimoNivel = True
        Exit Function
    End If

    'Llegados aqui, si que existe la cuenta
    If otroCampo = "S" Then 'Si es apunte directo
        CuentaCorrectaUltimoNivel = True
        devuelve = sql
    Else
        devuelve = "No es apunte directo: " & cuenta
    End If
End Function


'DAVID
Public Function RellenaCodigoCuenta(vCodigo As String) As String
'Rellena con ceros hasta poner una cuenta.
'Ejemplo: 43.1 --> 430000001
Dim i As Integer
Dim J As Integer
Dim Cont As Integer
Dim cad As String

    RellenaCodigoCuenta = vCodigo
    If Len(vCodigo) > vEmpresa.DigitosUltimoNivel Then Exit Function
    
    i = 0: Cont = 0
    Do
        i = i + 1
        i = InStr(i, vCodigo, ".")
        If i > 0 Then
            If Cont > 0 Then Cont = 1000
            Cont = Cont + i
        End If
    Loop Until i = 0

    'Habia mas de un punto
    If Cont > 1000 Or Cont = 0 Then Exit Function

    'Cambiamos el punto por 0's  .-Utilizo la variable maximocaracteres, para no tener k definir mas
    i = Len(vCodigo) - 1 'el punto lo quito
    J = vEmpresa.DigitosUltimoNivel - i
    cad = ""
    For i = 1 To J
        cad = cad & "0"
    Next i

    cad = Mid(vCodigo, 1, Cont - 1) & cad
    cad = cad & Mid(vCodigo, Cont + 1)
    RellenaCodigoCuenta = cad
End Function

'DAVID
Public Function EsCuentaUltimoNivel(cuenta As String) As Boolean
    EsCuentaUltimoNivel = (Len(cuenta) = vEmpresa.DigitosUltimoNivel)
End Function

' ### [Monica] 07/09/2006
' copia de la gestion
Private Function InsertarCuentaCble(cuenta As String, cadSocio As String, Optional cadProve As String, Optional cadBanco As String) As Boolean
Dim sql As String
Dim SqlBan As String
Dim Rs As ADODB.Recordset
Dim vsocio As CSocio
Dim vProve As CProveedor
Dim b As Boolean

    On Error GoTo EInsCta
    
    sql = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,maidatos,webdatos,obsdatos,pais, entidad, oficina, cc, cuentaba "
    '[Monica]22/11/2013: tema iban
    If vEmpresa.HayNorma19_34Nueva = 1 Then
        sql = sql & ", iban) "
    Else
        sql = sql & ") "
    End If
    
    sql = sql & " VALUES (" & DBSet(cuenta, "T") & ","
    
    If cadSocio <> "" Then
        Set vsocio = New CSocio
        If vsocio.LeerDatos(cadSocio) Then
            sql = sql & DBSet(vsocio.Nombre, "T") & ",'S',1," & DBSet(cuenta, "T") & "," & DBSet(vsocio.Domicilio, "T") & ","
            sql = sql & DBSet(vsocio.CPostal, "T") & "," & DBSet(vsocio.Poblacion, "T") & "," & DBSet(vsocio.Provincia, "T") & "," & DBSet(vsocio.NIF, "T") & "," & DBSet(vsocio.EMailAdm, "T") & "," & DBSet(vsocio.Websocio, "T") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(vsocio.Banco, "N") & "," & DBSet(vsocio.Sucursal, "N") & "," & DBSet(vsocio.DigControl, "T") & "," & DBSet(vsocio.CuentaBan, "T") ' & ")"
            
            '[Monica]22/11/2013: tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                sql = sql & "," & DBSet(vsocio.Iban, "T") & ")"
            Else
                sql = sql & ")"
            End If
            
            ConnConta.Execute sql
            cadSocio = vsocio.Nombre
            b = True
        Else
            b = False
        End If
        Set vsocio = Nothing
    End If
    
    If cadProve <> "" Then
        Set vProve = New CProveedor
        If vProve.LeerDatos(cadProve) Then
            sql = sql & DBSet(vProve.Nombre, "T") & ",'S',1," & DBSet(vProve.Nombre, "T") & "," & DBSet(vProve.Domicilio, "T") & ","
            sql = sql & DBSet(vProve.CPostal, "T") & "," & DBSet(vProve.Poblacion, "T") & "," & DBSet(vProve.Provincia, "T") & "," & DBSet(vProve.NIF, "T") & "," & DBSet(vProve.EMailAdmon, "T") & "," & DBSet(vProve.WebProve, "T") & "," & ValorNulo & ",'ESPAÑA'," & DBSet(vProve.Banco, "N") & "," & DBSet(vProve.Sucursal, "N") & "," & DBSet(vProve.DigControl, "T") & "," & DBSet(vProve.CuentaBan, "T")  '& ")"
            
            '[Monica]22/11/2013: tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                sql = sql & "," & DBSet(vProve.Iban, "T") & ")"
            Else
                sql = sql & ")"
            End If
            
            
            ConnConta.Execute sql
            cadProve = vProve.Nombre
            b = True
        Else
            b = False
        End If
        Set vProve = Nothing
    End If
    
    If cadBanco <> "" Then
        SqlBan = "select * from sbanco where codbanpr = " & DBSet(cadBanco, "N")
        
        Set Rs = New ADODB.Recordset
        Rs.Open SqlBan, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            sql = sql & DBSet(Rs!nombanco, "T") & ",'S',1," & DBSet(Rs!nombanco, "T") & "," & DBSet(Rs!dombanco, "T") & ","
            sql = sql & DBSet(Rs!codposta, "T") & "," & DBSet(Rs!pobbanco, "T") & "," & DBSet(Rs!probanco, "T") & "," & ValorNulo & "," & DBSet(Rs!maibanco, "T") & "," & DBSet(Rs!wwwbanco, "T") & "," & ValorNulo & ",'ESPAÑA'," & DBSet(Rs!codBanco, "N") & "," & DBSet(Rs!codSucur, "N") & "," & DBSet(Rs!DigContr, "T") & "," & DBSet(Rs!CuentaBa, "T")  '& ")"
            
            '[Monica]22/11/2013: tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                sql = sql & "," & DBSet(Rs!Iban, "T") & ")"
            Else
                sql = sql & ")"
            End If
            
            
            ConnConta.Execute sql
            cadBanco = Rs!nombanco
            b = True
        Else
            b = False
        End If
        Set Rs = Nothing
    End If
    
EInsCta:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Description, "Insertando cuenta contable", Err.Description
    End If
    InsertarCuentaCble = b
End Function


'=============================================================================
'==========     CENTROS DE COSTE
'=============================================================================
'LAURA
Public Function PonerNombreCCoste(Empresa As String, ByRef Txt As TextBox) As String
'Obtener el nombre de un centro de coste
Dim codCCoste As String
Dim cad As String

     If Txt.Text = "" Then
         PonerNombreCCoste = ""
         Exit Function
    End If
    codCCoste = Txt.Text
    If CCosteCorrecto(Empresa, codCCoste, cad) Then
        Txt.Text = codCCoste
        PonerNombreCCoste = cad
    Else
        MsgBox cad, vbExclamation
'        Txt.Text = ""
        PonerNombreCCoste = ""
        PonerFoco Txt
    End If
'    codCCoste = ""
End Function

'LAURA
Public Function CCosteCorrecto(Empresa As String, ByRef Centro As String, ByRef devuelve As String) As Boolean
    Dim sql As String
    
    CCosteCorrecto = False
 
    'BD 2: conexion a BD Conta
    If Val(Empresa) <> Val(vEmpresa.codEmpre) Then
        sql = DevuelveDesdeBDNew(3, "cabccost", "nomccost", "codccost", Centro, "T")
    Else
        sql = DevuelveDesdeBDNew(cConta, "cabccost", "nomccost", "codccost", Centro, "T")
    End If
    If sql = "" Then
        devuelve = "No existe el Centro de coste : " & Centro
        Exit Function
    Else
        devuelve = sql
        CCosteCorrecto = True
    End If
End Function




'=============================================================================
'==========     CONCEPTOS
'=============================================================================
'LAURA
Public Function PonerNombreConcepto(ByRef Txt As TextBox) As String
'Obtener el nombre de un concepto
Dim codConce As String
Dim cad As String

     If Txt.Text = "" Then
         PonerNombreConcepto = ""
         Exit Function
    End If
    codConce = Txt.Text
    If ConceptoCorrecto(codConce, cad) Then
        Txt.Text = Format(codConce, "000")
        PonerNombreConcepto = cad
    Else
        MsgBox cad, vbExclamation
        Txt.Text = ""
        PonerNombreConcepto = ""
        PonerFoco Txt
    End If
End Function


'LAURA
Public Function ConceptoCorrecto(ByRef Concep As String, ByRef devuelve As String) As Boolean
    Dim sql As String
    
    ConceptoCorrecto = False
 
    'BD 2: conexion a BD Conta
    sql = DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", Concep, "N")
    If sql = "" Then
        devuelve = "No existe el concepto : " & Concep
        Exit Function
    Else
        devuelve = sql
        ConceptoCorrecto = True
    End If
End Function

' ### [Monica] 27/09/2006
Public Function FacturaContabilizada(numserie As String, numfactu As String, Anofactu As String) As Boolean
Dim sql As String
Dim NumAsi As Currency

    FacturaContabilizada = False
    sql = ""
    sql = DevuelveDesdeBDNew(cConta, "cabfact", "numasien", "numserie", numserie, "T", , "codfaccl", numfactu, "N", "anofaccl", Anofactu, "N")
    
    If sql = "" Then Exit Function
    
    NumAsi = DBLet(sql, "N")
    
    If NumAsi <> 0 Then FacturaContabilizada = True

End Function

' ### [Monica] 27/09/2006
Public Function FacturaRemesada(numserie As String, numfactu As String, fecfactu As String) As Boolean
Dim sql As String
Dim NumRem As Currency
Dim Rs As ADODB.Recordset

    FacturaRemesada = False
    
    '[Monica]22/05/2014: hay que ver que ningun vencimiento este remesado si hay mas de uno
'    sql = ""
'    sql = DevuelveDesdeBDNew(cConta, "scobro", "codrem", "numserie", numserie, "T", , "codfaccl", numfactu, "N", "fecfaccl", fecfactu, "F")
    sql = "select codrem from scobro where numserie = " & DBSet(numserie, "T") & " and codfaccl = " & DBSet(numfactu, "N") & " and fecfaccl = " & DBSet(fecfactu, "F")
    sql = sql & " order by codrem desc "
    Set Rs = New ADODB.Recordset
    Rs.Open sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
    
    Else
        Exit Function
    End If
    
'    If sql = "" Then Exit Function
    
    NumRem = DBLet(Rs!codrem, "N")
    
    If NumRem <> 0 Then FacturaRemesada = True
    
    Set Rs = Nothing
    
End Function

' ### [Monica] 27/09/2006
Public Function FacturaCobrada(numserie As String, numfactu As String, fecfactu As String) As Boolean
Dim sql As String
Dim ImpCob As Currency

    FacturaCobrada = False
    sql = ""
    sql = DevuelveDesdeBDNew(cConta, "scobro", "impcobro", "numserie", numserie, "T", , "codfaccl", numfactu, "N", "fecfaccl", fecfactu, "F")
    If sql = "" Then Exit Function
    ImpCob = DBLet(sql, "N")
    
    If ImpCob <> 0 Then FacturaCobrada = True
    
End Function

' ### [Monica] 27/09/2006
Public Function ModificaClienteFacturaContabilidad(letraser As String, numfactu As String, fecfactu As String, CtaConta As String, Tipo As Byte) As Boolean
Dim sql As String
Dim Anyo As Currency

    On Error GoTo eModificaClienteFacturaContabilidad

    ModificaClienteFacturaContabilidad = False

    Anyo = Year(CDate(fecfactu))
    
    '[Monica]24/07/2013: añadido el tipo 2 (hco1)
    If Tipo = 0 Or Tipo = 2 Then
        sql = "update cabfact set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(letraser, "T") & " and " & _
                  "codfaccl = " & DBSet(numfactu, "N") & " and anofaccl = " & DBSet(Anyo, "N")
        ConnConta.Execute sql
    End If
    
    sql = "update scobro set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(letraser, "T") & " and " & _
              "codfaccl = " & DBSet(numfactu, "N") & " and fecfaccl = " & DBSet(fecfactu, "F")
              
    ConnConta.Execute sql
              
    ModificaClienteFacturaContabilidad = True
    
eModificaClienteFacturaContabilidad:
    If Err.Number <> 0 Then
        MsgBox "Error en ModificaClienteFacturaContabilidad: " & Err.Description, vbExclamation
    End If

End Function

' ### [Monica] 27/09/2006
Public Sub ModificaFormaPagoTesoreria(letraser As String, numfactu As String, fecfactu As String, forpa As String, forpaant As String)
Dim sql As String
Dim SQL1 As String
Dim TipForpa As String
Dim TipForpaAnt As String
Dim cadWHERE As String

    cadWHERE = " numserie = " & DBSet(letraser, "T") & " and " & _
              "codfaccl = " & numfactu & " and fecfaccl = " & DBSet(fecfactu, "F")
    
    sql = "update scobro set codforpa = " & forpa & " where " & cadWHERE

    ConnConta.Execute sql

End Sub

' ### [Monica] 29/09/2006
Public Function ModificaImportesFacturaContabilidad(letraser As String, numfactu As String, fecfactu As String, Importe As String, forpa As String, vTabla As String) As Boolean
Dim sql As String
Dim vWhere As String
Dim b As Boolean
Dim CadValues As String
Dim vsocio As CSocio
Dim Rs As ADODB.Recordset
Dim TipForpa As String

    On Error GoTo eModificaImportesFacturaContabilidad
    
    b = False
    
    vWhere = "numserie = " & DBSet(letraser, "T") & " and codfaccl = " & _
              numfactu & " and anofaccl = " & Format(Year(fecfactu), "0000")
     
    
    sql = "select codsocio from " & vTabla & " where letraser = " & DBSet(letraser, "T") & " and numfactu = " & _
           numfactu & " and fecfactu = " & DBSet(fecfactu, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then Rs.MoveFirst
    
    Set vsocio = New CSocio
    If vsocio.LeerDatos(Rs.Fields(0).Value) Then
        '[Monica]24/07/2013
        If vTabla = "schfac" Or vTabla = "schfac1" Then
            sql = "delete from linfact where " & vWhere
            ConnConta.Execute sql
        
            sql = "delete from cabfact where " & vWhere
            ConnConta.Execute sql
            
            '[Monica]24/07/2013
            If vTabla = "schfac" Then
                sql = "schfac.letraser = " & DBSet(letraser, "T") & " and numfactu = " & numfactu
                sql = sql & " and fecfactu = " & DBSet(fecfactu, "F")
            Else
                sql = "schfac1.letraser = " & DBSet(letraser, "T") & " and numfactu = " & numfactu
                sql = sql & " and fecfactu = " & DBSet(fecfactu, "F")
            End If
            
            b = CrearTMPErrFact("schfac")
            If b Then b = PasarFactura2(sql, vsocio, vTabla)
        End If
        
        If vTabla = "schfacr" Then
            b = CrearTMPErrFact("schfacr")
        End If
        
        ' 09/02/2007
        TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", forpa, "N")
        '[Monica]04/01/2013: efectivos
                                                     
        If TipForpa <> "0" And TipForpa <> "6" And b Then
            b = ModificaCobroTesoreria(letraser, numfactu, fecfactu, vsocio, vTabla)
        End If
    End If
    
    ModificaImportesFacturaContabilidad = b
    
eModificaImportesFacturaContabilidad:
    If Err.Number <> 0 Then
        MsgBox "Error en ModificaImportesFacturaContabilidad: " & Err.Description, vbExclamation
    End If
End Function

Public Function ModificaCobroTesoreria(letraser As String, numfactu As String, fecfactu As String, vsocio As CSocio, vTabla As String) As Boolean
Dim sql As String
Dim Rs As ADODB.Recordset
Dim cadWHERE As String
Dim BanPr As String
Dim Mens As String
Dim b As Boolean

    On Error GoTo eModificaCobroTesoreria

    ModificaCobroTesoreria = False
    b = True
    
    ' antes de borrar he de obtener la fecha de vencimiento y el codmacta para sacar el banco propio que le pasaré
    ' a la rutina de InsertarEnTesoreria
            
    sql = "select fecvenci, ctabanc1 from scobro where numserie = " & DBSet(letraser, "T") & " and codfaccl = " & DBSet(numfactu, "N")
    sql = sql & " and fecfaccl = " & DBSet(fecfactu, "F") & " and numorden = 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        Rs.MoveFirst
        
        cadWHERE = vTabla & ".letraser =" & DBSet(letraser, "T") & " and numfactu=" & DBLet(numfactu, "N")
        cadWHERE = cadWHERE & " and fecfactu=" & DBSet(fecfactu, "F")

        BanPr = ""
        BanPr = DevuelveDesdeBDNew(cPTours, "sbanco", "codbanpr", "codmacta", Rs.Fields(1).Value, "T")

        sql = "delete from scobro where "
        sql = sql & " numserie = " & DBSet(letraser, "T") & " and codfaccl = " & numfactu
        sql = sql & " and fecfaccl = " & DBSet(fecfactu, "F")
        
        ConnConta.Execute sql
            
        ' hemos de crear el cobro nuevamente
        Mens = "Insertando en Tesoreria "
        b = InsertarEnTesoreria(cadWHERE, CStr(Rs.Fields(0).Value), BanPr, Mens, vsocio, vTabla)
    End If
    
    ModificaCobroTesoreria = b
    
eModificaCobroTesoreria:
    If Err.Number <> 0 Then
        MsgBox "Error en ModificaCobroTesoreria " & Err.Description, vbExclamation
    End If
End Function


Public Function CalcularIva(Importe As String, articulo As String) As Currency
'devuelve el iva del Importe
'Ej el 16% de 120 = 19.2
Dim vImp As Currency
Dim vIva As Currency
Dim vArt As Currency
Dim CodIVA As String

Dim IvaArt As Integer
Dim iva As String
Dim impiva As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    articulo = ComprobarCero(articulo)
    
    CodIVA = DevuelveDesdeBD("codigiva", "sartic", "codartic", articulo, "N")
    iva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIVA, "N")
    
    vImp = CCur(Importe)
    vIva = CCur(iva)
    
    impiva = ((vImp * vIva) / 100)
    impiva = Round(impiva, 2)
    
    CalcularIva = CStr(impiva)
    If Err.Number <> 0 Then Err.Clear

End Function


Public Function CalcularBase(Importe As String, articulo As String) As Currency
'devuelve la base del Importe
'Ej el 16% de 120 = 120-19.2 = 100.8
Dim vImp As Currency
Dim vIva As Currency
Dim vArt As Currency
Dim CodIVA As String

Dim IvaArt As Integer
Dim iva As String
Dim impiva As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    articulo = ComprobarCero(articulo)
    
    CodIVA = DevuelveDesdeBD("codigiva", "sartic", "codartic", articulo, "N")
    iva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIVA, "N")
    
    vImp = CCur(Importe)
    vIva = CCur(iva)
    
    impiva = Round2(Importe / (1 + (vIva / 100)), 2)
    
    CalcularBase = CStr(impiva)
    If Err.Number <> 0 Then Err.Clear

End Function


'MONICA: Cuentas del la Contabilidad
Public Function NombreCuentaCorrecta(ByRef cuenta As String) As String
    'Comprueba si es numerica
    Dim sql As String
    Dim otroCampo As String
    
' ### [Monica] 27/10/2006 añadida la linea siguiente condicion vParamAplic.NumeroConta = 0
' para que no saque nada si no hay contabilidad
    If cuenta = "" Or vParamAplic.NumeroConta = 0 Then
         NombreCuentaCorrecta = ""
         Exit Function
    End If
    
    NombreCuentaCorrecta = ""
    If cuenta = "" Then
        MsgBox "Cuenta vacia", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(cuenta) Then
        MsgBox "La cuenta debe de ser numérica: " & cuenta, vbExclamation
        Exit Function
    End If

    'BD 2: conexion a BD Conta
    sql = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", cuenta, "T")
    If sql = "" Then
        MsgBox "No existe la cuenta : " & cuenta, vbExclamation
    Else
        NombreCuentaCorrecta = sql
    End If

End Function

