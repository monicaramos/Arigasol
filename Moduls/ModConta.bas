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
Public Function CuentaCorrectaUltimoNivel(ByRef Cuenta As String, ByRef devuelve As String) As Boolean
    'Comprueba si es numerica
    Dim SQL As String
    Dim otroCampo As String
    
    CuentaCorrectaUltimoNivel = False
    If Cuenta = "" Then
        devuelve = "Cuenta vacia"
        Exit Function
    End If

    If Not IsNumeric(Cuenta) Then
        devuelve = "La cuenta debe de ser numérica: " & Cuenta
        Exit Function
    End If

    'Rellenamos si procede
    Cuenta = RellenaCodigoCuenta(Cuenta)

    '==========
    If Not EsCuentaUltimoNivel(Cuenta) Then
        devuelve = "No es cuenta de último nivel: " & Cuenta
        Exit Function
    End If
    '==================

    otroCampo = "apudirec"
    'BD 2: conexion a BD Conta
    SQL = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Cuenta, "T", otroCampo)
    If SQL = "" Then
        devuelve = "No existe la cuenta : " & Cuenta
        CuentaCorrectaUltimoNivel = True
        Exit Function
    End If

    'Llegados aqui, si que existe la cuenta
    If otroCampo = "S" Then 'Si es apunte directo
        CuentaCorrectaUltimoNivel = True
        devuelve = SQL
    Else
        devuelve = "No es apunte directo: " & Cuenta
    End If
End Function


'DAVID
Public Function RellenaCodigoCuenta(vCodigo As String) As String
'Rellena con ceros hasta poner una cuenta.
'Ejemplo: 43.1 --> 430000001
Dim i As Integer
Dim J As Integer
Dim cont As Integer
Dim cad As String

    RellenaCodigoCuenta = vCodigo
    If Len(vCodigo) > vEmpresa.DigitosUltimoNivel Then Exit Function
    
    i = 0: cont = 0
    Do
        i = i + 1
        i = InStr(i, vCodigo, ".")
        If i > 0 Then
            If cont > 0 Then cont = 1000
            cont = cont + i
        End If
    Loop Until i = 0

    'Habia mas de un punto
    If cont > 1000 Or cont = 0 Then Exit Function

    'Cambiamos el punto por 0's  .-Utilizo la variable maximocaracteres, para no tener k definir mas
    i = Len(vCodigo) - 1 'el punto lo quito
    J = vEmpresa.DigitosUltimoNivel - i
    cad = ""
    For i = 1 To J
        cad = cad & "0"
    Next i

    cad = Mid(vCodigo, 1, cont - 1) & cad
    cad = cad & Mid(vCodigo, cont + 1)
    RellenaCodigoCuenta = cad
End Function

'DAVID
Public Function EsCuentaUltimoNivel(Cuenta As String) As Boolean
    EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresa.DigitosUltimoNivel)
End Function

' ### [Monica] 07/09/2006
' copia de la gestion
Private Function InsertarCuentaCble(Cuenta As String, cadSocio As String, Optional cadProve As String, Optional cadBanco As String) As Boolean
Dim SQL As String
Dim SqlBan As String
Dim Rs As ADODB.Recordset
Dim vsocio As CSocio
Dim vProve As CProveedor
Dim b As Boolean
Dim vIban As String

    On Error GoTo EInsCta
    If vParamAplic.ContabilidadNueva Then
        SQL = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,maidatos,webdatos,obsdatos,codpais,forpa "
    Else
        SQL = "INSERT INTO cuentas (codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,maidatos,webdatos,obsdatos,pais,forpa,entidad, oficina, cc, cuentaba "
    End If
    '[Monica]22/11/2013: tema iban
    If vEmpresa.HayNorma19_34Nueva = 1 Then
        SQL = SQL & ", iban) "
    Else
        SQL = SQL & ") "
    End If
    
    SQL = SQL & " VALUES (" & DBSet(Cuenta, "T") & ","
    
    If cadSocio <> "" Then
        Set vsocio = New CSocio
        If vsocio.LeerDatos(cadSocio) Then
            SQL = SQL & DBSet(vsocio.Nombre, "T") & ",'S',1," & DBSet(Cuenta, "T") & "," & DBSet(vsocio.Domicilio, "T") & ","
            SQL = SQL & DBSet(vsocio.CPostal, "T") & "," & DBSet(vsocio.POBLACION, "T") & "," & DBSet(vsocio.Provincia, "T") & "," & DBSet(vsocio.NIF, "T") & "," & DBSet(vsocio.EMailAdm, "T") & "," & DBSet(vsocio.Websocio, "T") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(vsocio.ForPago, "N")
            
            If Not vParamAplic.ContabilidadNueva Then
                SQL = SQL & "," & DBSet(vsocio.Banco, "N") & "," & DBSet(vsocio.Sucursal, "N") & "," & DBSet(vsocio.DigControl, "T") & "," & DBSet(vsocio.CuentaBan, "T") ' & ")"
            
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    SQL = SQL & "," & DBSet(vsocio.IBAN, "T") & ")"
                Else
                    SQL = SQL & ")"
                End If
            Else
                vIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                
                SQL = SQL & "," & DBSet(vIban, "T") & ")"
            End If
            
            ConnConta.Execute SQL
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
            SQL = SQL & DBSet(vProve.Nombre, "T") & ",'S',1," & DBSet(vProve.Nombre, "T") & "," & DBSet(vProve.Domicilio, "T") & ","
            SQL = SQL & DBSet(vProve.CPostal, "T") & "," & DBSet(vProve.POBLACION, "T") & "," & DBSet(vProve.Provincia, "T") & "," & DBSet(vProve.NIF, "T") & "," & DBSet(vProve.EMailAdmon, "T") & "," & DBSet(vProve.WebProve, "T") & "," & ValorNulo
            If Not vParamAplic.ContabilidadNueva Then
                SQL = SQL & ",'ESPAÑA'," & DBSet(vProve.ForPago, "N") & "," & DBSet(vProve.Banco, "N") & "," & DBSet(vProve.Sucursal, "N") & "," & DBSet(vProve.DigControl, "T") & "," & DBSet(vProve.CuentaBan, "T")
            
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    SQL = SQL & "," & DBSet(vProve.IBAN, "T") & ")"
                Else
                    SQL = SQL & ")"
                End If
            Else
                SQL = SQL & ",'ES'," & DBSet(vProve.ForPago, "N")
                
                vIban = MiFormat(vProve.IBAN, "") & MiFormat(vProve.Banco, "0000") & MiFormat(vProve.Sucursal, "0000") & MiFormat(vProve.DigControl, "00") & MiFormat(vProve.CuentaBan, "0000000000")
            
                SQL = SQL & "," & DBSet(vIban, "T") & ")"
            End If
            
            ConnConta.Execute SQL
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
            SQL = SQL & DBSet(Rs!NomBanco, "T") & ",'S',1," & DBSet(Rs!NomBanco, "T") & "," & DBSet(Rs!dombanco, "T") & ","
            SQL = SQL & DBSet(Rs!codPosta, "T") & "," & DBSet(Rs!pobbanco, "T") & "," & DBSet(Rs!probanco, "T") & "," & ValorNulo & "," & DBSet(Rs!maibanco, "T") & "," & DBSet(Rs!wwwbanco, "T") & "," & ValorNulo
            
            If Not vParamAplic.ContabilidadNueva Then
            
                SQL = SQL & ",'ESPAÑA'," & ValorNulo & "," & DBSet(Rs!codbanco, "N") & "," & DBSet(Rs!codsucur, "N") & "," & DBSet(Rs!digcontr, "T") & "," & DBSet(Rs!cuentaba, "T")
                
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    SQL = SQL & "," & DBSet(Rs!IBAN, "T") & ")"
                Else
                    SQL = SQL & ")"
                End If
            Else
                SQL = SQL & ",'ES'," & ValorNulo
                
                vIban = MiFormat(Rs!IBAN, "") & MiFormat(Rs!codbanco, "0000") & "," & MiFormat(Rs!codsucur, "0000") & "," & MiFormat(Rs!digcontr, "00") & "," & MiFormat(Rs!cuentaba, "0000000000")
                
                SQL = SQL & "," & DBSet(vIban, "T") & ")"
            
            End If
            
            ConnConta.Execute SQL
            cadBanco = Rs!NomBanco
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
    Dim SQL As String
    
    CCosteCorrecto = False
 
    'BD 2: conexion a BD Conta
    If Val(Empresa) <> Val(vEmpresa.codEmpre) Then
        SQL = DevuelveDesdeBDNew(3, "cabccost", "nomccost", "codccost", Centro, "T")
    Else
        SQL = DevuelveDesdeBDNew(cConta, "cabccost", "nomccost", "codccost", Centro, "T")
    End If
    If SQL = "" Then
        devuelve = "No existe el Centro de coste : " & Centro
        Exit Function
    Else
        devuelve = SQL
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
    Dim SQL As String
    
    ConceptoCorrecto = False
 
    'BD 2: conexion a BD Conta
    SQL = DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", Concep, "N")
    If SQL = "" Then
        devuelve = "No existe el concepto : " & Concep
        Exit Function
    Else
        devuelve = SQL
        ConceptoCorrecto = True
    End If
End Function

' ### [Monica] 27/09/2006
Public Function FacturaContabilizada(numserie As String, numfactu As String, Anofactu As String) As Boolean
Dim SQL As String
Dim NumAsi As Currency

    FacturaContabilizada = False
    SQL = ""
    If vParamAplic.ContabilidadNueva Then
        SQL = DevuelveDesdeBDNew(cConta, "factcli", "numasien", "numserie", numserie, "T", , "numfactu", numfactu, "N", "anofactu", Anofactu, "N")
    Else
        SQL = DevuelveDesdeBDNew(cConta, "cabfact", "numasien", "numserie", numserie, "T", , "codfaccl", numfactu, "N", "anofaccl", Anofactu, "N")
    End If
    
    If SQL = "" Then Exit Function
    
    NumAsi = DBLet(SQL, "N")
    
    If NumAsi <> 0 Then FacturaContabilizada = True

End Function

' ### [Monica] 27/09/2006
Public Function FacturaRemesada(numserie As String, numfactu As String, fecfactu As String) As Boolean
Dim SQL As String
Dim NumRem As Currency
Dim Rs As ADODB.Recordset

    FacturaRemesada = False
    
    '[Monica]22/05/2014: hay que ver que ningun vencimiento este remesado si hay mas de uno
'    sql = ""
'    sql = DevuelveDesdeBDNew(cConta, "scobro", "codrem", "numserie", numserie, "T", , "codfaccl", numfactu, "N", "fecfaccl", fecfactu, "F")
    SQL = "select codrem from scobro where numserie = " & DBSet(numserie, "T") & " and codfaccl = " & DBSet(numfactu, "N") & " and fecfaccl = " & DBSet(fecfactu, "F")
    SQL = SQL & " order by codrem desc "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
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
Dim SQL As String
Dim ImpCob As Currency

    FacturaCobrada = False
    SQL = ""
    SQL = DevuelveDesdeBDNew(cConta, "scobro", "impcobro", "numserie", numserie, "T", , "codfaccl", numfactu, "N", "fecfaccl", fecfactu, "F")
    If SQL = "" Then Exit Function
    ImpCob = DBLet(SQL, "N")
    
    If ImpCob <> 0 Then FacturaCobrada = True
    
End Function

' ### [Monica] 27/09/2006
Public Function ModificaClienteFacturaContabilidad(letraser As String, numfactu As String, fecfactu As String, CtaConta As String, Tipo As Byte) As Boolean
Dim SQL As String
Dim Anyo As Currency

    On Error GoTo eModificaClienteFacturaContabilidad

    ModificaClienteFacturaContabilidad = False

    Anyo = Year(CDate(fecfactu))
    
    '[Monica]24/07/2013: añadido el tipo 2 (hco1)
    If Tipo = 0 Or Tipo = 2 Then
        If vParamAplic.ContabilidadNueva Then
            SQL = "update factcli set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(letraser, "T") & " and " & _
                      "numfactu = " & DBSet(numfactu, "N") & " and anofactu = " & DBSet(Anyo, "N")
        Else
            SQL = "update cabfact set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(letraser, "T") & " and " & _
                      "codfaccl = " & DBSet(numfactu, "N") & " and anofaccl = " & DBSet(Anyo, "N")
        End If
        ConnConta.Execute SQL
    End If
    
    If vParamAplic.ContabilidadNueva Then
        SQL = "update cobros set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(letraser, "T") & " and " & _
                  "numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(fecfactu, "F")
    
    Else
        SQL = "update scobro set codmacta = " & DBSet(CtaConta, "T") & " where numserie = " & DBSet(letraser, "T") & " and " & _
                  "codfaccl = " & DBSet(numfactu, "N") & " and fecfaccl = " & DBSet(fecfactu, "F")
    End If
    ConnConta.Execute SQL
              
    ModificaClienteFacturaContabilidad = True
    
eModificaClienteFacturaContabilidad:
    If Err.Number <> 0 Then
        MsgBox "Error en ModificaClienteFacturaContabilidad: " & Err.Description, vbExclamation
    End If

End Function

' ### [Monica] 27/09/2006
Public Sub ModificaFormaPagoTesoreria(letraser As String, numfactu As String, fecfactu As String, Forpa As String, ForpaAnt As String)
Dim SQL As String
Dim Sql1 As String
Dim TipForpa As String
Dim TipForpaAnt As String
Dim cadWhere As String

    If vParamAplic.ContabilidadNueva Then
        cadWhere = " numserie = " & DBSet(letraser, "T") & " and " & _
                  "numfactu = " & numfactu & " and fecfactu = " & DBSet(fecfactu, "F")
        
        SQL = "update cobros set codforpa = " & Forpa & " where " & cadWhere
    
    Else

        cadWhere = " numserie = " & DBSet(letraser, "T") & " and " & _
                  "codfaccl = " & numfactu & " and fecfaccl = " & DBSet(fecfactu, "F")
        
        SQL = "update scobro set codforpa = " & Forpa & " where " & cadWhere
    End If
    ConnConta.Execute SQL

End Sub

' ### [Monica] 29/09/2006
Public Function ModificaImportesFacturaContabilidad(letraser As String, numfactu As String, fecfactu As String, Importe As String, Forpa As String, vTabla As String) As Boolean
Dim SQL As String
Dim vWhere As String
Dim b As Boolean
Dim CadValues As String
Dim vsocio As CSocio
Dim Rs As ADODB.Recordset
Dim TipForpa As String

    On Error GoTo eModificaImportesFacturaContabilidad
    
    b = False
    
    If vParamAplic.ContabilidadNueva Then
        vWhere = "numserie = " & DBSet(letraser, "T") & " and numfactu = " & _
                  numfactu & " and anofactu = " & Format(Year(fecfactu), "0000")
    
    Else
        vWhere = "numserie = " & DBSet(letraser, "T") & " and codfaccl = " & _
                  numfactu & " and anofaccl = " & Format(Year(fecfactu), "0000")
    End If
    
    SQL = "select codsocio from " & vTabla & " where letraser = " & DBSet(letraser, "T") & " and numfactu = " & _
           numfactu & " and fecfactu = " & DBSet(fecfactu, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then Rs.MoveFirst
    
    Set vsocio = New CSocio
    If vsocio.LeerDatos(Rs.Fields(0).Value) Then
        '[Monica]24/07/2013
        If vTabla = "schfac" Or vTabla = "schfac1" Then
            If vParamAplic.ContabilidadNueva Then
                SQL = "delete from factcli_lineas where " & vWhere
                ConnConta.Execute SQL
                
                SQL = "delete from factcli_totales where " & vWhere
                ConnConta.Execute SQL
            
                SQL = "delete from factcli where " & vWhere
                ConnConta.Execute SQL
            
            Else
                SQL = "delete from linfact where " & vWhere
                ConnConta.Execute SQL
            
                SQL = "delete from cabfact where " & vWhere
                ConnConta.Execute SQL
            End If
            '[Monica]24/07/2013
            If vTabla = "schfac" Then
                SQL = "schfac.letraser = " & DBSet(letraser, "T") & " and numfactu = " & numfactu
                SQL = SQL & " and fecfactu = " & DBSet(fecfactu, "F")
            Else
                SQL = "schfac1.letraser = " & DBSet(letraser, "T") & " and numfactu = " & numfactu
                SQL = SQL & " and fecfactu = " & DBSet(fecfactu, "F")
            End If
            
            b = CrearTMPErrFact("schfac")
            If b Then b = PasarFactura2(SQL, vsocio, vTabla)
        End If
        
        If vTabla = "schfacr" Then
            b = CrearTMPErrFact("schfacr")
        End If
        
        ' 09/02/2007
        TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", Forpa, "N")
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
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim cadWhere As String
Dim Banpr As String
Dim Mens As String
Dim b As Boolean

    On Error GoTo eModificaCobroTesoreria

    ModificaCobroTesoreria = False
    b = True
    
    ' antes de borrar he de obtener la fecha de vencimiento y el codmacta para sacar el banco propio que le pasaré
    ' a la rutina de InsertarEnTesoreria
            
    SQL = "select fecvenci, ctabanc1 from scobro where numserie = " & DBSet(letraser, "T") & " and codfaccl = " & DBSet(numfactu, "N")
    SQL = SQL & " and fecfaccl = " & DBSet(fecfactu, "F") & " and numorden = 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        Rs.MoveFirst
        
        cadWhere = vTabla & ".letraser =" & DBSet(letraser, "T") & " and numfactu=" & DBLet(numfactu, "N")
        cadWhere = cadWhere & " and fecfactu=" & DBSet(fecfactu, "F")

        Banpr = ""
        Banpr = DevuelveDesdeBDNew(cPTours, "sbanco", "codbanpr", "codmacta", Rs.Fields(1).Value, "T")

        SQL = "delete from scobro where "
        SQL = SQL & " numserie = " & DBSet(letraser, "T") & " and codfaccl = " & numfactu
        SQL = SQL & " and fecfaccl = " & DBSet(fecfactu, "F")
        
        ConnConta.Execute SQL
            
        ' hemos de crear el cobro nuevamente
        Mens = "Insertando en Tesoreria "
        b = InsertarEnTesoreria(cadWhere, CStr(Rs.Fields(0).Value), Banpr, Mens, vsocio, vTabla)
    End If
    
    ModificaCobroTesoreria = b
    
eModificaCobroTesoreria:
    If Err.Number <> 0 Then
        MsgBox "Error en ModificaCobroTesoreria " & Err.Description, vbExclamation
    End If
End Function


Public Function CalcularIva(Importe As String, Articulo As String) As Currency
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
    Articulo = ComprobarCero(Articulo)
    
    CodIVA = DevuelveDesdeBD("codigiva", "sartic", "codartic", Articulo, "N")
    iva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIVA, "N")
    
    vImp = CCur(Importe)
    vIva = CCur(iva)
    
    impiva = ((vImp * vIva) / 100)
    impiva = Round(impiva, 2)
    
    CalcularIva = CStr(impiva)
    If Err.Number <> 0 Then Err.Clear

End Function


Public Function CalcularBase(Importe As String, Articulo As String) As Currency
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
    Articulo = ComprobarCero(Articulo)
    
    CodIVA = DevuelveDesdeBD("codigiva", "sartic", "codartic", Articulo, "N")
    iva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIVA, "N")
    
    vImp = CCur(Importe)
    vIva = CCur(iva)
    
    impiva = Round2(Importe / (1 + (vIva / 100)), 2)
    
    CalcularBase = CStr(impiva)
    If Err.Number <> 0 Then Err.Clear

End Function


'MONICA: Cuentas del la Contabilidad
Public Function NombreCuentaCorrecta(ByRef Cuenta As String) As String
    'Comprueba si es numerica
    Dim SQL As String
    Dim otroCampo As String
    
' ### [Monica] 27/10/2006 añadida la linea siguiente condicion vParamAplic.NumeroConta = 0
' para que no saque nada si no hay contabilidad
    If Cuenta = "" Or vParamAplic.NumeroConta = 0 Then
         NombreCuentaCorrecta = ""
         Exit Function
    End If
    
    NombreCuentaCorrecta = ""
    If Cuenta = "" Then
        MsgBox "Cuenta vacia", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(Cuenta) Then
        MsgBox "La cuenta debe de ser numérica: " & Cuenta, vbExclamation
        Exit Function
    End If

    'BD 2: conexion a BD Conta
    SQL = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Cuenta, "T")
    If SQL = "" Then
        MsgBox "No existe la cuenta : " & Cuenta, vbExclamation
    Else
        NombreCuentaCorrecta = SQL
    End If

End Function

