Attribute VB_Name = "ModContabilizar"
' copia del ariges

Option Explicit



'===================================================================================
'CONTABILIZAR FACTURAS:
'Modulo para el traspaso de registros de cabecera y lineas de tablas de FACTURACION
'A las tablas de FACTURACION de Contabilidad
'====================================================================================
Private DtoGnral As Currency
Private DtoPPago As Currency
Private BaseImp As Currency
Private TotalFac As Currency
Private CCoste As String
Private conCtaAlt As Boolean 'el cliente utiliza cuentas alternativas

Private AnyoFacPr As Integer 'año factura proveedor, es el ano de fecha_recepcion


Public Function CrearTMPFacturas(cadTABLA As String, cadWHERE As String) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    If cadTABLA = "scafpc" Then
        Sql = "CREATE TEMPORARY TABLE tmpfactu ( "
        Sql = Sql & "codprove int(6) NOT NULL default '0',"
        Sql = Sql & "numfactu varchar(10) NOT NULL default '',"
        Sql = Sql & "fecfactu date NOT NULL default '0000-00-00') "
        Conn.Execute Sql
         
         
        Sql = "SELECT codprove, numfactu, fecfactu"
        Sql = Sql & " FROM " & cadTABLA
        Sql = Sql & " WHERE " & cadWHERE
        Sql = " INSERT INTO tmpfactu " & Sql
        Conn.Execute Sql
    
        CrearTMPFacturas = True
    
    
    Else
    
        Sql = "CREATE TEMPORARY TABLE tmpfactu ( "
        Sql = Sql & "letraser char(3) NOT NULL default '',"
        Sql = Sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
        Sql = Sql & "fecfactu date NOT NULL default '0000-00-00') "
        Conn.Execute Sql
         
         
        Sql = "SELECT letraser, numfactu, fecfactu"
        Sql = Sql & " FROM " & cadTABLA
        Sql = Sql & " WHERE " & cadWHERE
        Sql = " INSERT INTO tmpfactu " & Sql
        Conn.Execute Sql
    
        CrearTMPFacturas = True
        
    End If
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpfactu;"
        Conn.Execute Sql
    End If
End Function


Public Sub BorrarTMPFacturas()
On Error Resume Next

    Conn.Execute " DROP TABLE IF EXISTS tmpfactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function CrearTMPErrFact(cadTABLA As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    Sql = "CREATE TEMPORARY TABLE tmperrfac ( "
    If cadTABLA = "schfac" Or cadTABLA = "schfacr" Then
        Sql = Sql & "codtipom char(1) NOT NULL default '',"
        Sql = Sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    End If
    Sql = Sql & "fecfactu date NOT NULL default '0000-00-00', "
    Sql = Sql & "error varchar(100) NULL )"
    Conn.Execute Sql
     
    CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmperrfac;"
        Conn.Execute Sql
    End If
End Function


Public Function CrearTMPErrComprob() As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrComprob = False
    
    Sql = "CREATE TEMPORARY TABLE tmperrcomprob ( "
    Sql = Sql & "error varchar(100) NULL )"
    Conn.Execute Sql
     
    CrearTMPErrComprob = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrComprob = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmperrcomprob;"
        Conn.Execute Sql
    End If
End Function



Public Sub BorrarTMPErrFact()
On Error Resume Next
    Conn.Execute " DROP TABLE IF EXISTS tmperrfac;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BorrarTMPErrComprob()
On Error Resume Next
    Conn.Execute " DROP TABLE IF EXISTS tmperrcomprob;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BorrarTMPAsiento()
On Error Resume Next
    Conn.Execute " DROP TABLE IF EXISTS tmpasien;"
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function ComprobarLetraSerie(cadTABLA As String) As Boolean
'Para Facturas VENTA a clientes
'Comprueba que la letra del serie del tipo de movimiento es  correcta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cad As String, devuelve As String

On Error GoTo EComprobarLetra

    ComprobarLetraSerie = False
    
    'Comprobar que existe la letra de serie en contabilidad
    If cadTABLA = "schfac" Then
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
        Sql = "Select distinct tiporegi from contadores"
        Set RSconta = New ADODB.Recordset
        RSconta.Open Sql, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
        If RSconta.EOF Then
            RSconta.Close
            Set RSconta = Nothing
            Exit Function
        End If
            
    
        'obtenemos los distintos tipos de movimiento que vamos a contabilizar
        'de las facturas seleccionadas
        Sql = "select distinct letraser from tmpfactu "

        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        b = True
        While Not Rs.EOF 'And b
            'comprobar que todas las letras serie existen en Arigasol
            Sql = "letraser"
            devuelve = DevuelveDesdeBD("letraser", "stipom", "letraser", DBLet(Rs!letraser), "T", Sql)
            If devuelve = "" Then
                b = False
                cad = Rs!letraser & " en BD de Arigasol."
                InsertarError "No existe la letra de serie " & cad
            Else
                'comprobar que todas las letras serie existen en la contabilidad
                devuelve = "tiporegi= '" & devuelve & "'"
                RSconta.MoveFirst
                RSconta.Find (devuelve), , adSearchForward
                If RSconta.EOF Then
                    'no encontrado
                    b = False
                    cad = Sql & " en BD de Contabilidad."
                    InsertarError "No existe la letra de serie " & cad
                End If
            End If
            If b Then cad = cad & DBSet(Rs!letraser, "T") & ","
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        RSconta.Close
        Set RSconta = Nothing
        
        If Not b Then 'Hay algun movimiento que no existe
            devuelve = "No existe el tipo de movimiento: " & cad & vbCrLf
            devuelve = devuelve & "Consulte con el administrador."
'            MsgBox devuelve, vbExclamation
            Exit Function
        End If
        
        'Todos los Tipo de movimiento existen
        If cad <> "" Then
            cad = Mid(cad, 1, Len(cad) - 1) 'quitamos ult. coma
        
            'miramos si hay algun movimiento de factura que la letra serie sea nulo
            Sql = "select count(*) from stipom "
            Sql = Sql & "where letraser IN (" & cad & ") and (isnull(letraser) or letraser='')"
            If RegistrosAListar(Sql) > 0 Then
                Sql = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                Sql = Sql & "Comprobar en la tabla de tipos de movimiento: " & cad
                InsertarError Sql
'                MsgBox sql, vbExclamation
                Exit Function
            End If
        End If
        ComprobarLetraSerie = True
    Else
        ComprobarLetraSerie = True
    End If

EComprobarLetra:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Letra Serie", Err.Description
    End If
End Function


Public Function ComprobarNumFacturas(cadTABLA As String, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas = False
    
    Sql = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
    Sql = Sql & " WHERE " & cadWConta
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        Sql = "SELECT DISTINCT tmpfactu.letraser,tmpfactu.numfactu,tmpfactu.fecfactu "
        Sql = Sql & " FROM tmpfactu "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
' quitado el 12022007
'            SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND codfaccl=" & DBSet(RS!numfactu, "N") & " AND anofaccl=" & Year(RS!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            Sql = ""
            Sql = DevuelveDesdeBDNew(cConta, "cabfact", "codfaccl", "codfaccl", Rs!numfactu, "N", , "numserie", Rs!letraser, "T", "anofaccl", Year(Rs!fecfactu), "N")
            If Sql <> "" Then
                b = False
                Sql = "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                Sql = Sql & "          Fecha: " & Rs!fecfactu
                
                Sql = "Ya existe la factura: " & vbCrLf & Sql
                InsertarError Sql
            
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            Sql = "Ya existe la factura: " & vbCrLf & Sql
            Sql = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & Sql
            
            'MsgBox sql, vbExclamation
            ComprobarNumFacturas = False
        Else
            ComprobarNumFacturas = True
        End If
    Else
        ComprobarNumFacturas = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompFactu:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Nº Facturas", Err.Description
    End If
End Function



Public Function ComprobarCtaContable(cadTABLA As String, Opcion As Byte, Optional cadWHERE As String) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim enc As String
    
    On Error GoTo ECompCta

    ComprobarCtaContable = False
    
    Sql = "SELECT codmacta FROM cuentas "
    Sql = Sql & " WHERE apudirec='S'"
    If cadG <> "" Then Sql = Sql & cadG
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        If Opcion = 1 Then
            Select Case cadTABLA
                Case "schfac"
                    'Seleccionamos los distintos clientes,cuentas que vamos a facturar
                    Sql = "SELECT DISTINCT schfac.codsocio, ssocio.codmacta "
                    Sql = Sql & " FROM (schfac INNER JOIN ssocio ON schfac.codsocio=ssocio.codsocio) "
                    Sql = Sql & " INNER JOIN tmpfactu ON schfac.letraser=tmpfactu.letraser AND schfac.numfactu=tmpfactu.numfactu AND schfac.fecfactu=tmpfactu.fecfactu "
                Case "ssocio"
                    Sql = "SELECT DISTINCT scaalb.codsocio, ssocio.codmacta "
                    Sql = Sql & " FROM scaalb, ssocio, sforpa  "
                    Sql = Sql & " where " & cadWHERE & " and scaalb.codsocio=ssocio.codsocio and scaalb.codforpa = sforpa.codforpa "
                Case "schfacr"
                    Sql = "SELECT DISTINCT schfacr.codsocio, ssocio.codmacta "
                    Sql = Sql & " FROM (schfacr INNER JOIN ssocio ON schfacr.codsocio=ssocio.codsocio) "
                    Sql = Sql & " INNER JOIN tmpfactu ON schfacr.letraser=tmpfactu.letraser AND schfacr.numfactu=tmpfactu.numfactu AND schfacr.fecfactu=tmpfactu.fecfactu "
            End Select
        ElseIf Opcion = 2 Then
                Sql = "SELECT distinct sartic.codartic "
                Sql = Sql & ", sartic.codmacta, sartic.codmaccl"
                Sql = Sql & " from ((slhfac "
                Sql = Sql & " INNER JOIN tmpfactu ON slhfac.letraser=tmpfactu.letraser AND slhfac.numfactu=tmpfactu.numfactu AND slhfac.fecfactu=tmpfactu.fecfactu) "
                Sql = Sql & "INNER JOIN sartic ON slhfac.codartic=sartic.codartic) "
                Sql = Sql & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
        ElseIf Opcion = 3 Then
                'si hay analitica comprobar que todas las cuentas
                'empiezan por el digito que hay en conta.parametros.grupovta
                cadG = DevuelveDesdeBDNew(cConta, "parametros", "grupovta", "", "", "")
        
                Sql = "SELECT distinct sartic.codartic "
                Sql = Sql & ", sartic.codmacta, sartic.codmaccl"
                Sql = Sql & " from ((slhfac "
                Sql = Sql & " INNER JOIN tmpfactu ON slhfac.letraser=tmpfactu.letraser AND slhfac.numfactu=tmpfactu.numfactu AND slhfac.fecfactu=tmpfactu.fecfactu) "
                Sql = Sql & "INNER JOIN sartic ON slhfac.codartic=sartic.codartic) "
                Sql = Sql & " where sartic.codmacta "
                If cadG <> "" Then
                     Sql = Sql & " AND not ((sartic.codmacta like '" & cadG & "%') and (sartic.codmaccl like '" & cadG & "%'))"
                End If
        ElseIf Opcion = 4 Then
            Sql = "select codmacta from sbanco where codbanpr = " & cadTABLA
        ElseIf Opcion = 5 Then
            Sql = "select codmacta from sforpa where cuadresn = 1 and not codmacta is null and mid(codmacta,1,1) <> ' '"
        ElseIf Opcion = 6 Then
            Sql = "select ctaposit from sparam"
        ElseIf Opcion = 7 Then
            Sql = "select ctanegtat from sparam"
        End If
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF 'And b
            If Opcion = 3 Then
                b = False
                Sql = DBLet(Rs!Codmacta, "T") & " o " & DBLet(Rs!Codmaccl, "T")
                Sql = "La cuenta " & Sql & " del articulo " & Rs!codArtic & " no es del grupo correcto."
                InsertarError Sql
            Else
                If Opcion = 6 Or Opcion = 7 Then
                    Sql = "codmacta= " & DBSet(Rs.Fields(0).Value, "T") '& " and apudirec='S' "
                Else
                    Sql = "codmacta= " & DBSet(Rs!Codmacta, "T") '& " and apudirec='S' "
                End If
' comentado 12022007
'                RSconta.MoveFirst
'                RSconta.Find (SQL), , adSearchForward
'                If RSconta.EOF Then
                 enc = ""
                 If Opcion = 6 Or Opcion = 7 Then
                    enc = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", DBLet(Rs.Fields(0).Value, "T"), "T")
                 Else
                    enc = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", DBLet(Rs!Codmacta, "T"), "T")
                 End If
                 
                 If enc = "" Then
                    b = False 'no encontrado
                    If Opcion = 1 Then
                        If cadTABLA = "schfac" Or cadTABLA = "ssocio" Or cadTABLA = "schfacr" Then
                            Sql = DBLet(Rs!Codmacta, "T") & " del Cliente " & Format(Rs!codsocio, "000000")
                            Sql = "No existe la cta contable " & Sql
                            InsertarError Sql
                        End If
                    End If
                    If Opcion = 2 Then
                        Sql = DBLet(Rs!Codmacta, "T") & " del Artículo " & Format(Rs!codArtic, "000000")
                        Sql = "No existe la cta contable " & Sql
                        InsertarError Sql
                    End If
                    If Opcion = 4 Then
                        Sql = DBLet(Rs!Codmacta, "T") & " del Banco " & Format(CCur(cadTABLA), "000")
                        Sql = "No existe la cta contable " & Sql
                        InsertarError Sql
                    End If
                    If Opcion = 6 Or Opcion = 7 Then
                        Sql = "No existe la cta contable " & Sql
                        InsertarError Sql
                    End If
                End If
                
                ' en caso de que estemos comprobando las cuentas contables del articulo
                ' comprobamos tb la cuenta contable socio del articulo
                '---------------------------------------------------------------------
                If Opcion = 2 Then
                    If Not IsNull(Rs!Codmaccl) Then
                        Sql = "codmacta= " & DBSet(Rs!Codmaccl, "T") '& " and apudirec='S' "
                        enc = ""
                        enc = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", DBLet(Rs!Codmaccl, "T"), "T")
                        If enc = "" Then
' comentado el 12022007
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
                            b = False 'no encontrado
                            Sql = DBLet(Rs!Codmaccl, "T") & " del artículo " & Format(Rs!codArtic, "000000")
                            Sql = "No existe la cta contable cliente " & Sql
                            InsertarError Sql
                        End If
                    Else
                        b = False 'no encontrado
                        Sql = DBLet(Rs!Codmaccl, "T") & " del artículo " & Format(Rs!codArtic, "000000")
                        Sql = "No existe la cta contable cliente " & Sql
                        InsertarError Sql
                    End If
                End If
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            ComprobarCtaContable = False
        Else
            ComprobarCtaContable = True
        End If
    Else
        ComprobarCtaContable = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function





Public Function ComprobarTiposIVA(cadTABLA As String) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    Sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For i = 1 To 3
            If cadTABLA = "schfac" Then
                Sql = "SELECT DISTINCT schfac.tipoiva" & i
                Sql = Sql & " FROM schfac "
                Sql = Sql & " INNER JOIN tmpfactu ON schfac.letraser=tmpfactu.letraser AND schfac.numfactu=tmpfactu.numfactu AND schfac.fecfactu=tmpfactu.fecfactu "
                Sql = Sql & " WHERE not isnull(tipoiva" & i & ")"
            Else
                If cadTABLA = "scafpc" Then
                    Sql = "SELECT DISTINCT scafpc.tipoiva" & i
                    Sql = Sql & " FROM scafpc "
                    Sql = Sql & " INNER JOIN tmpfactu ON scafpc.codprove=tmpfactu.codprove AND scafpc.numfactu=tmpfactu.numfactu AND scafpc.fecfactu=tmpfactu.fecfactu "
                    Sql = Sql & " WHERE not isnull(tipoiva" & i & ")"
                End If
            End If

            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            b = True
            While Not Rs.EOF 'And b
                If Rs.Fields(0) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                    Sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                    RSconta.MoveFirst
                    RSconta.Find (Sql), , adSearchForward
                    If RSconta.EOF Then
                        b = False 'no encontrado
                        Sql = "No existe el " & Sql
                        Sql = "Tipo de IVA: " & Rs.Fields(0)
                        InsertarError Sql
                    End If
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not b Then
                ComprobarTiposIVA = False
                Exit For
            Else
                ComprobarTiposIVA = True
            End If
        Next i
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function


Public Function PasarFactura(cadWHERE As String, fecvenci As String, BanPr As String, CodCCost As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim Sql As String
Dim vsocio As CSocio
Dim codsoc As Long

Dim LetraInt As String  ' letra de serie de las facturas internas

Dim Rs As ADODB.Recordset

Dim RSx As ADODB.Recordset
Dim Sql2 As String
Dim codfor As Integer
Dim TipForpa As String
Dim Mc As CContadorContab
Dim Obs As String





    On Error GoTo EContab

    ConnConta.BeginTrans
    Conn.BeginTrans
    
    'seleccionamos el socio de la factura
    '[Monica]04/03/2011: Facturas internas añado en el select la letra de serie
    Sql = "select codsocio, letraser, fecfactu from schfac where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenStatic, adLockPessimistic, adCmdText
    
    codsoc = 0
    
    If Not Rs.EOF Then
        codsoc = Rs.Fields(0).Value
        LetraInt = Rs.Fields(1).Value
    End If
    
    Set vsocio = New CSocio
    If vsocio.LeerDatos(CStr(codsoc)) Then
'[Monica]25/07/2013: serie internas
'        '[Monica]04/03/2011: Facturas internas añado en el select la letra de serie
'        If LetraInt = vParamAplic.LetraInt Then
        If EsInterna(LetraInt) Then
            Set Mc = New CContadorContab
            
            If Mc.ConseguirContador("0", (Rs!fecfactu <= CDate(FFin)), True) = 0 Then
            
                Obs = "Contabilización Factura Interna de Fecha " & Format(Rs!fecfactu, "dd/mm/yyyy")
            
                'Insertar en la conta Cabecera Asiento
                b = InsertarCabAsientoDia(vEmpresa.NumDiarioInt, Mc.Contador, Rs!fecfactu, Obs, cadMen)
                cadMen = "Insertando Cab. Asiento: " & cadMen
            Else
                b = False
            End If
        Else
            'Insertar en la conta Cabecera Factura
            b = InsertarCabFact(cadWHERE, cadMen)
            cadMen = "Insertando Cab. Factura: " & cadMen
        End If
            
        ' insertar en tesoreria
        If b Then
            Sql2 = "select codforpa from schfac where " & cadWHERE
            Set RSx = New ADODB.Recordset
            RSx.Open Sql2, Conn, adOpenStatic, adLockPessimistic, adCmdText
            
            If Not RSx.EOF Then codfor = RSx.Fields(0).Value
            TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", DBSet(RSx.Fields(0).Value, "N"), "N")
            
'[Monica]16/12/2010: solo se inserta en tesoreria si no hacen la contabilizacion de cierre de turno
            '[Monica]04/01/2013: Efectivos
            '[Monica]11/01/2013: En Ribarroja se inserta en Tesoreria
            If (TipForpa <> "0" And TipForpa <> "6") Or vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 5 Then
            
                b = InsertarEnTesoreria(cadWHERE, fecvenci, BanPr, cadMen, vsocio, "schfac")
                cadMen = "Insertando en Tesoreria: " & cadMen
            End If
            
            Set RSx = Nothing
            
        End If
    
        If b Then
'[Monica]25/07/2013: serie internas
'            If LetraInt = vParamAplic.LetraInt Then
            If EsInterna(LetraInt) Then
                b = InsertarLinAsientoFactInt("schfac", cadWHERE, cadMen, vsocio, Mc.Contador)
                cadMen = "Insertando Lin. Factura Interna: " & cadMen
            
                Set Mc = Nothing
            Else
        '        CCoste = CodCCost
                'Insertar lineas de Factura en la Conta
                '21032007
                '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
                If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 5 Then ' si Alzira o Pobla del Duc o Ribarroja
                    b = InsertarLinFact("schfac", cadWHERE, cadMen, vsocio)
                Else
                    b = InsertarLinFactReg("schfac", cadWHERE, cadMen, vsocio)
                End If
                cadMen = "Insertando Lin. Factura: " & cadMen
                
            End If
            If b Then
                'Poner intconta=1 en arigasol.scafac
                b = ActualizarCabFact("schfac", cadWHERE, cadMen)
                cadMen = "Actualizando Factura: " & cadMen
            End If
        End If
        
        If Not b Then
            Sql = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
            Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
            Sql = Sql & " WHERE " & Replace(cadWHERE, "schfac", "tmpfactu")
            Conn.Execute Sql
        End If
    End If
    
    Set vsocio = Nothing
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        Conn.CommitTrans
        PasarFactura = True
    Else
        ConnConta.RollbackTrans
        Conn.RollbackTrans
        PasarFactura = False
    End If
End Function

Public Function PasarFactura2(cadWHERE As String, ByRef vsocio As CSocio, vTabla As String) As Boolean   ' , CodCCost As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim Sql As String

    On Error GoTo EContab
    
    'Insertar en la conta Cabecera Factura
    b = InsertarCabFact(cadWHERE, cadMen, vTabla)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
'        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        b = InsertarLinFact("schfac", cadWHERE, cadMen, vsocio)
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            'Poner intconta=1 en arigasol.scafac
            b = ActualizarCabFact("schfac", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    If Not b Then
        Sql = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "scafac", "tmpfactu")
        Conn.Execute Sql
    End If
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        PasarFactura2 = True
    Else
        PasarFactura2 = False
    End If
End Function

Public Function PasarFactura3(cadWHERE As String, fecvenci As String, BanPr As String, CodCCost As String, ByRef cadMen As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
'Dim cadMen As String
Dim Sql As String
Dim vsocio As CSocio
Dim codsoc As Long
Dim Rs As ADODB.Recordset

Dim RSx As ADODB.Recordset
Dim Sql2 As String
Dim codfor As Integer
Dim TipForpa As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    Conn.BeginTrans
    
    'seleccionamos el socio de la factura
    Sql = "select codsocio from schfacr where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenStatic, adLockPessimistic, adCmdText
    
    codsoc = 0
    
    If Not Rs.EOF Then codsoc = Rs.Fields(0).Value
    
    
    Set vsocio = New CSocio
    If vsocio.LeerDatos(CStr(codsoc)) Then
    
        
        ' insertar en tesoreria
        Sql2 = "select codforpa from schfacr where " & cadWHERE
        Set RSx = New ADODB.Recordset
        RSx.Open Sql2, Conn, adOpenStatic, adLockPessimistic, adCmdText
        
        If Not RSx.EOF Then codfor = RSx.Fields(0).Value
        TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", DBSet(RSx.Fields(0).Value, "N"), "N")
        '[Monica]04/01/2013: efectivos
        If TipForpa <> "0" And TipForpa <> "6" Then
            b = InsertarEnTesoreria(cadWHERE, fecvenci, BanPr, cadMen, vsocio, "schfacr")
            cadMen = "Insertando en Tesoreria: " & cadMen
        End If
        
        Set RSx = Nothing
        
        If b Then
            'Poner intconta=1 en arigasol.scafac
            b = ActualizarCabFact("schfacr", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
'--monica:07-04-2008
'        If Not b Then
'            sql = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
'            sql = sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
'            sql = sql & " WHERE " & Replace(cadwhere, "schfacr", "tmpfactu")
'            Conn.Execute sql
'        End If
    End If
    
    Set vsocio = Nothing
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura Ajena en Tesorería", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        Conn.CommitTrans
        PasarFactura3 = True
    Else
        ConnConta.RollbackTrans
        Conn.RollbackTrans
        PasarFactura3 = False
    End If
End Function


Public Function PasarFactura4(letraser As String, numfactu As String, fecfactu As String, ByRef vsocio As CSocio, NueNumfact As Long, NueFecFactu As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
Dim b As Boolean
Dim cadMen As String
Dim Sql As String
Dim Sql2 As String
Dim RSx As ADODB.Recordset
Dim cadWHERE As String
Dim codfor As String
Dim TipForpa As String
Dim ctabancl As String
Dim fecvenci As String
Dim BanPr As String

    On Error GoTo EContab
    
    'Insertar en la conta Cabecera Factura
    cadWHERE = "letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(NueNumfact, "N") & " and fecfactu = " & DBSet(NueFecFactu, "F")
    
    b = InsertarCabFact(cadWHERE, cadMen)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
'        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        b = InsertarLinFact("schfac", cadWHERE, cadMen, vsocio)
        cadMen = "Insertando Lin. Factura: " & cadMen

        Sql2 = "select codforpa from schfac where " & cadWHERE
        Set RSx = New ADODB.Recordset
        RSx.Open Sql2, Conn, adOpenStatic, adLockPessimistic, adCmdText
        
        If Not RSx.EOF Then codfor = RSx.Fields(0).Value
        TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", DBSet(RSx.Fields(0).Value, "N"), "N")
        
        ctabancl = "ctabanc1"
        '[Monica]04/01/2013 : efectivos
        If TipForpa <> "0" And TipForpa <> "6" Then
            fecvenci = ""
            fecvenci = DevuelveDesdeBDNew(cConta, "scobro", "fecvenci", "numserie", letraser, "T", ctabancl, "codfaccl", numfactu, "N", "fecfaccl", fecfactu, "F")
            If fecvenci <> "" Then
                BanPr = DevuelveDesdeBDNew(cPTours, "sbanco", "codbanpr", "codmacta", ctabancl, "T")
            
                b = InsertarEnTesoreria(cadWHERE, fecvenci, BanPr, cadMen, vsocio, "schfac")
                cadMen = "Insertando en Tesoreria: " & cadMen
            End If
        End If
        
        Set RSx = Nothing
    End If
    
EContab:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Contabilizando Factura", Err.Description & cadMen
        PasarFactura4 = False
    Else
        PasarFactura4 = True
    End If
End Function



Private Function InsertarCabFact(cadWHERE As String, caderr As String, Optional vTabla As String) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String

    On Error GoTo EInsertar
    
    Sql = " SELECT letraser,numfactu,fecfactu, ssocio.codmacta, year(fecfactu) as anofaccl,"
    Sql = Sql & "baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "totalfac,tipoiva1,tipoiva2,tipoiva3 "
    '[Monica]24/07/2013:
    If vTabla <> "" Then
        Sql = Sql & " FROM " & vTabla
        Sql = Sql & "INNER JOIN " & "ssocio ON " & vTabla & ".codsocio=ssocio.codsocio"
    Else
        Sql = Sql & " FROM " & "schfac "
        Sql = Sql & "INNER JOIN " & "ssocio ON schfac.codsocio=ssocio.codsocio"
    End If
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        BaseImp = DBLet(Rs!baseimp1, "N") + DBLet(Rs!baseimp2, "N") + DBLet(Rs!baseimp3, "N")
        
        Sql = ""
        Sql = DBSet(Rs!letraser, "T") & "," & DBSet(Rs!numfactu, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Rs!Codmacta, "T") & "," & Year(Rs!fecfactu) & ",'FACTURACION',"
        Sql = Sql & DBSet(Rs!baseimp1, "N") & "," & DBSet(Rs!baseimp2, "N") & "," & DBSet(Rs!baseimp3, "N") & "," & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N") & "," & DBSet(Rs!porciva3, "N") & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!impoiva1, "N", "N") & "," & DBSet(Rs!impoiva2, "N") & "," & DBSet(Rs!impoiva3, "N") & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", "S") & "," & DBSet(Rs!TipoIVA3, "N", "S") & ",0,"
        Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        Sql = Sql & DBSet(Rs!fecfactu, "F")
        cad = cad & "(" & Sql & ")"
'        RS.MoveNext
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    'Insertar en la contabilidad
    Sql = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
    Sql = Sql & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
    Sql = Sql & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
    Sql = Sql & " VALUES " & cad
    ConnConta.Execute Sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFact = False
        caderr = Err.Description
    Else
        InsertarCabFact = True
    End If
End Function


Private Function InsertarLinAsientoFactInt(cadTABLA As String, cadWHERE As String, caderr As String, ByRef vsocio As CSocio, Optional Contador As Long) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim numdocum As String
Dim ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim i As Long
Dim b As Boolean
Dim cad As String
Dim cadMen As String
Dim FeFact As Date

    On Error GoTo eInsertarLinAsientoFactInt

    InsertarLinAsientoFactInt = False
    
    '[Monica]25/09/2014: cambiado tipoconta = 1 indica sobre cuenta contable del socio, 0 = cuenta contable del cliente
    If vsocio.TipoConta = 1 Then
        Sql = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmacta, " ' sartic.codmaccl, "
        Sql = Sql & " sum(implinea) as importe FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "schfac", "slhfac")
        Sql = Sql & " GROUP BY 1,2,3,5"
    Else
        Sql = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmaccl codmacta, "
        Sql = Sql & " sum(implinea) as importe FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "schfac", "slhfac")
        Sql = Sql & " GROUP BY 1,2,3,5"
    End If

    
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
            
    i = 0
    ImporteD = 0
    ImporteH = 0
    
    numdocum = Format(Rs!numfactu, "0000000")
    '[Monica]25/07/2013: letra de serie
'    ampliacion = vParamAplic.LetraInt & "-" & Format(Rs!numfactu, "0000000")
    ampliacion = Trim(Rs!letraser) & "-" & Format(Rs!numfactu, "0000000")
    ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & ampliacion
    ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & ampliacion
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    b = True
    
    While Not Rs.EOF And b
        i = i + 1
        
        FeFact = Rs!fecfactu
        
        cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(Contador, "N") & ","
        cad = cad & DBSet(i, "N") & "," & DBSet(Rs!Codmacta, "T") & "," & DBSet(numdocum, "T") & ","
        
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If Rs.Fields(5).Value < 0 Then
            ' importe al debe en positivo
            cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs.Fields(5).Value * (-1), "N") & ","
            cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(vsocio.CuentaConta, "T") & "," & ValorNulo & ",0"
        
            ImporteD = ImporteD + (CCur(Rs.Fields(5).Value) * (-1))
        Else
            ' importe al haber en positivo, cambiamos el signo
            cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            cad = cad & DBSet((Rs.Fields(5).Value), "N") & "," & ValorNulo & "," & DBSet(vsocio.CuentaConta, "T") & "," & ValorNulo & ",0"
        
            ImporteH = ImporteH + CCur(Rs.Fields(5).Value)
        End If
        
        cad = "(" & cad & ")"
        
        b = InsertarLinAsientoDia(cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & i

        Rs.MoveNext
    Wend
    
    If b And i > 0 Then
        i = i + 1
                
        ' el Total es sobre la cuenta del cliente
        cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FeFact, "F") & "," & DBSet(Contador, "N") & ","
        cad = cad & DBSet(i, "N") & ","
        cad = cad & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(numdocum, "T") & ","
            
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If ImporteD - ImporteH > 0 Then
            ' importe al debe en positivo
            cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
            cad = cad & DBSet(ImporteD - ImporteH, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        Else
            ' importe al haber en positivo, cambiamos el signo
            cad = cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet(((ImporteD - ImporteH) * -1), "N") & ","
            cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        End If
        
        cad = "(" & cad & ")"
        
        b = InsertarLinAsientoDia(cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & i
        
    End If
        
    Set Rs = Nothing
    InsertarLinAsientoFactInt = b
    Exit Function
    
eInsertarLinAsientoFactInt:
    caderr = "Insertar Linea Asiento Factura Interna: " & Err.Description
    caderr = caderr & cadMen
    InsertarLinAsientoFactInt = False
End Function


Private Function InsertarLinFact(cadTABLA As String, cadWHERE As String, caderr As String, ByRef vsocio As CSocio, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SqlAux As String
Dim Sql2 As String

Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim iva As String
Dim vIva As Currency


    On Error GoTo EInLinea

    If cadTABLA = "schfac" Then
        '[Monica]25/09/2014: cambiado tipoconta = 1 indica sobre cuenta contable del socio, 0 = cuenta contable del cliente
        If vsocio.TipoConta = 1 Then
            Sql = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmacta, " ' sartic.codmaccl, "
            Sql = Sql & " sum(implinea) as importe FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            Sql = Sql & " WHERE " & Replace(cadWHERE, "schfac", "slhfac")
            Sql = Sql & " GROUP BY 1,2,3,5"
        Else
            Sql = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmaccl, "
            Sql = Sql & " sum(implinea) as importe FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            Sql = Sql & " WHERE " & Replace(cadWHERE, "schfac", "slhfac")
            Sql = Sql & " GROUP BY 1,2,3,5"
        End If
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    i = 1
    totimp = 0
    SqlAux = ""
    While Not Rs.EOF
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        'de multibase
        'Let v_base = Round(basesfac / (1 + (porc_iva / 100)), 2)
'        Implinea = CCur(CalcularBase(CStr(RS!Importe), CStr(RS!codartic)))
        SqlAux = cad
        
        ImpLinea = CCur(CalcularBase(CStr(Rs.Fields(5).Value), CStr(Rs!codArtic)))
        
        ImpLinea = Round2(ImpLinea, 2)
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & i & ","
        
        '[Monica]25/09/2014: cambiado tipoconta = 1 indica sobre cuenta contable del socio, 0 = cuenta contable del cliente
        If vsocio.TipoConta = 1 Then
            Sql = Sql & DBSet(Rs!Codmacta, "T")
        Else
            Sql = Sql & DBSet(Rs!Codmaccl, "T")
        End If
        
        Sql2 = Sql & ","
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If CCoste = "" Then
            Sql = Sql & ValorNulo
        Else
            Sql = Sql & DBSet(CCoste, "T")
        End If
        
        cad = cad & "(" & Sql & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        
        If CCoste = "" Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SqlAux <> "" Then 'hay mas de una linea
            cad = SqlAux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
        
        
'        Aux = Replace(sql, DBSet(Implinea, "N"), DBSet(totimp, "N"))
'        cad = Replace(cad, sql, Aux)
    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        If cadTABLA = "schfac" Then
            Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        End If
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact = False
        caderr = Err.Description
    Else
        InsertarLinFact = True
    End If
End Function



Private Function InsertarLinFactReg(cadTABLA As String, cadWHERE As String, caderr As String, ByRef vsocio As CSocio, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SQL1 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim iva As String
Dim vIva As Currency
Dim impuesto As Currency
Dim Impue As Currency
Dim TotalImpuesto As Currency

Dim numfactu As Long
Dim letraser As String
Dim fecfactu As Date

    On Error GoTo EInLinea

    '[Monica]25/09/2014: cambiado tipoconta = 1 indica sobre cuenta contable del socio, 0 = cuenta contable del cliente
    If vsocio.TipoConta = 1 Then
        Sql = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmacta, " ' sartic.codmaccl, "
        Sql = Sql & " sum(implinea) as importe, sum(cantidad) as cantidad FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "schfac", "slhfac")
        Sql = Sql & " GROUP BY 1,2,3,5"
    Else
        Sql = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmaccl, "
        Sql = Sql & " sum(implinea) as importe, sum(cantidad) as cantidad FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "schfac", "slhfac")
        Sql = Sql & " GROUP BY 1,2,3,5"
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    i = 1
    
    totimp = 0
    TotalImpuesto = 0
    
    While Not Rs.EOF
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        'de multibase
        'Let v_base = Round(basesfac / (1 + (porc_iva / 100)), 2)
'        Implinea = CCur(CalcularBase(CStr(RS!Importe), CStr(RS!codartic)))
        
        numfactu = Rs!numfactu
        letraser = Rs!letraser
        fecfactu = Rs!fecfactu
        
        
        ' se quita el impuesto por linea
        SQL1 = ""
        SQL1 = DevuelveDesdeBD("impuesto", "sartic", "codartic", DBLet(Rs!codArtic), "N")
        If SQL1 = "" Then
            impuesto = 0
        Else
            impuesto = CCur(SQL1) ' Comprueba si es nulo y lo pone a 0 o ""
        End If
        
        If EsArticuloCombustible(Rs!codArtic) Then
            Impue = Round2((Rs.Fields(6).Value * impuesto), 2)
            TotalImpuesto = TotalImpuesto + Impue
        End If
        
        
        ImpLinea = CCur(CalcularBase(CStr(Rs.Fields(5).Value), CStr(Rs!codArtic))) - Impue
        ImpLinea = Round2(ImpLinea, 2)
        
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & i & ","
        
        '[Monica]25/09/2014: cambiado tipoconta = 1 indica sobre cuenta contable del socio, 0 = cuenta contable del cliente
        If vsocio.TipoConta = 1 Then
            Sql = Sql & DBSet(Rs!Codmacta, "T")
        Else
            Sql = Sql & DBSet(Rs!Codmaccl, "T")
        End If
        
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If CCoste = "" Then
            Sql = Sql & ValorNulo
        Else
            Sql = Sql & DBSet(CCoste, "T")
        End If
        
        cad = cad & "(" & Sql & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    totimp = totimp + TotalImpuesto
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Aux = Replace(Sql, DBSet(ImpLinea, "N"), DBSet(totimp, "N"))
        cad = Replace(cad, Sql, Aux)
    End If

    ' insertamos la linea de base de impuesto
    '20/12/2012: dependiendo de la fecha de cambio
    If fecfactu < CDate(vParamAplic.FechaCam) Then
        Sql = ""
        Sql = "'" & letraser & "'," & numfactu & "," & Year(fecfactu) & "," & i & ","
        Sql = Sql & DBSet(vParamAplic.CtaImpuesto, "T")
        Sql = Sql & "," & DBSet(TotalImpuesto, "N") & ","
        If CCoste = "" Then
            Sql = Sql & ValorNulo
        Else
            Sql = Sql & DBSet(CCoste, "T")
        End If
        cad = cad & "(" & Sql & "),"
    End If
    
    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        If cadTABLA = "schfac" Then
            Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        End If
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactReg = False
        caderr = Err.Description
    Else
        InsertarLinFactReg = True
    End If
End Function







Private Function ActualizarCabFact(cadTABLA As String, cadWHERE As String, caderr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    
    Sql = "UPDATE " & cadTABLA & " SET intconta=1 "
    Sql = Sql & " WHERE " & cadWHERE

    Conn.Execute Sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFact = False
        caderr = Err.Description
    Else
        ActualizarCabFact = True
    End If
End Function



' ### [Monica] 02/10/2006
' copiado de la clase de laura cfactura
Public Function InsertarEnTesoreria(cadWHERE As String, ByVal FechaVen As String, BanPr As String, MenError As String, ByRef vsocio As CSocio, vTabla As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim RSx As ADODB.Recordset
Dim Sql As String, textcsb33 As String, textcsb41 As String
Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Sql5 As String
Dim Rs3 As ADODB.Recordset
Dim Rs4 As ADODB.Recordset
Dim Rs5 As ADODB.Recordset

Dim textcsb42 As String, textcsb43 As String
Dim textcsb51 As String, textcsb52 As String, textcsb53 As String
Dim textcsb61 As String, textcsb62 As String, textcsb63 As String
Dim textcsb71 As String, textcsb72 As String, textcsb73 As String
Dim textcsb81 As String, textcsb82 As String, textcsb83 As String
Dim n_linea As Integer
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci1 As Date
Dim ImpVenci As Single
Dim i As Byte
Dim CodmacBPr As String
Dim cadWHERE2 As String

Dim FacturaFP As String

Dim ForPago As String
Dim Ndias As String
Dim fecvenci As Date
Dim rsVenci As ADODB.Recordset
Dim TotalFactura2 As Currency

Dim LetraS As String

    On Error GoTo EInsertarTesoreria

    b = False
    InsertarEnTesoreria = False
    CadValues = ""
    CadValues2 = ""

    Sql = "select * from " & vTabla & " where  " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
    
        textcsb33 = "FACT: " & DBLet(Rs!letraser, "T") & "-" & Format(DBLet(Rs!numfactu, "N"), "0000000") & " " & Format(DBLet(Rs!fecfactu, "F"), "dd/mm/yy")
        textcsb33 = textcsb33 & " de " & DBSet(Rs!TotalFac, "N")
        ' añadido 07022007
'        textcsb41 = "'B.IMP " & DBSet(RS!baseimp1, "N") & " IVA " & DBSet(RS!impoiva1, "N") & " TOTAL " & DBSet(RS!TOTALFAC, "N") & "',"
        ' end del añadido
        
        ' añadido 08022007
        textcsb41 = ""
        textcsb42 = ""
        textcsb43 = ""
        textcsb51 = ""
        textcsb52 = ""
        textcsb53 = ""
        textcsb61 = ""
        textcsb62 = ""
        textcsb63 = ""
        textcsb71 = ""
        textcsb72 = ""
        textcsb73 = ""
        textcsb81 = ""
        textcsb82 = ""
        textcsb83 = ""
        
'[Monica]22/11/2013: quitamos el resto de textos csbs

        Select Case vTabla
            Case "schfac"
                cadWHERE2 = Replace(cadWHERE, "schfac", "slhfac")
            Case "schfacr"
                cadWHERE2 = Replace(cadWHERE, "schfacr", "slhfacr")
            Case "schfac1"
                cadWHERE2 = Replace(cadWHERE, "schfac1", "slhfac1")
        End Select


'        Sql = "select count(distinct numalbar) from " & vTabla & " where " & cadWHERE
'        '[Monica]24/07/2013
'        Select Case vTabla
'            Case "schfac"
'                Sql = Replace(Sql, "schfac", "slhfac")
'                cadWHERE2 = Replace(cadWHERE, "schfac", "slhfac")
'            Case "schfacr"
'                Sql = Replace(Sql, "schfacr", "slhfacr")
'                cadWHERE2 = Replace(cadWHERE, "schfacr", "slhfacr")
'            Case "schfac1"
'                Sql = Replace(Sql, "schfac1", "slhfac1")
'                cadWHERE2 = Replace(cadWHERE, "schfac1", "slhfac1")
'        End Select
'        If TotalRegistros(Sql) <= 15 Then
''            cadwhere2 = Replace(cadwhere, "schfac", "slhfac")
'            Sql2 = "select numalbar, fecalbar, sum(implinea) "
'            Select Case vTabla
'                Case "schfac"
'                    Sql2 = Sql2 & " from slhfac where " & cadWHERE2
'                Case "schfacr"
'                    Sql2 = Sql2 & " from slhfacr where " & cadWHERE2
'                Case "schfac1"
'                    Sql2 = Sql2 & " from slhfac1 where " & cadWHERE2
'            End Select
'
'            Sql2 = Sql2 & " group by numalbar, fecalbar order by numalbar, fecalbar "
'            Set RSx = New ADODB.Recordset
'            RSx.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            n_linea = 0
'            While Not RSx.EOF
'                n_linea = n_linea + 1
'                Select Case n_linea
'                    Case 1
'                        textcsb41 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 2
'                        textcsb42 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 3
'                        textcsb43 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 4
'                        textcsb51 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 5
'                        textcsb52 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 6
'                        textcsb53 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 7
'                        textcsb61 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 8
'                        textcsb62 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 9
'                        textcsb63 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 10
'                        textcsb71 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 11
'                        textcsb72 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 12
'                        textcsb73 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 13
'                        textcsb81 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 14
'                        textcsb82 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 15
'                        textcsb83 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                End Select
'
'
'                RSx.MoveNext
'            Wend
'        End If
'        ' end del añadido 08022007
'[Monica]22/11/2013: hasta aqui el tema de no grabar los csbs
        
'[Monica]08/01/2014: lo cambiamos rellenando lo maximo que podemos
        If vParamAplic.Cooperativa = 5 Then
            Dim cad1 As String
            Dim cad2 As String
            Dim cad22 As String
            
            Sql = "select count(distinct numalbar) from " & vTabla & " where " & cadWHERE
            cad1 = ""
            Sql2 = "select numalbar, fecalbar, sum(implinea) "
            Select Case vTabla
                Case "schfac"
                    Sql2 = Sql2 & " from slhfac where " & cadWHERE2
                Case "schfacr"
                    Sql2 = Sql2 & " from slhfacr where " & cadWHERE2
                Case "schfac1"
                    Sql2 = Sql2 & " from slhfac1 where " & cadWHERE2
            End Select

            Sql2 = Sql2 & " group by numalbar, fecalbar order by numalbar, fecalbar "
            
            Set RSx = New ADODB.Recordset
            RSx.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            n_linea = 0
            cad2 = " "
            cad22 = ""
            While Not RSx.EOF
                n_linea = n_linea + 1
            
                cad1 = "T-" & Right("        " & DBLet(RSx.Fields(0).Value, "T"), 8) & " " & Format(DBLet(RSx.Fields(2).Value, "N"), "##0.00") & "€ "
                                
                If n_linea <= 2 Then
                    cad2 = cad2 & cad1
                Else
                    cad22 = cad22 & cad1
                End If
                RSx.MoveNext
            Wend
            If cad2 <> "" Then textcsb33 = textcsb33 & cad2
            textcsb41 = Mid(cad22, 1, InStrRev(Mid(cad22, 1, 40), "€"))
            If Len(cad22) > 40 Then textcsb41 = textcsb41 & "..."
        End If
        
        
        '--[Monica]05/08/2011: quito esto pq ahora ya no tiene sentido
'        'monica 01/06/2007
'        FacturaFP = ""
'        FacturaFP = DevuelveDesdeBDNew(cPTours, "ssocio", "facturafp", "codsocio", RS!codsocio, "N")
'        If CInt(FacturaFP) = 1 Then
'            Ndias = ""
'            Ndias = DevuelveDesdeBDNew(cPTours, "sforpa", "diasvto", "codforpa", RS!Codforpa, "N")
'            Ndias = ComprobarCero(Ndias)
'            FecVenci1 = CDate(DBLet(RS!fecfactu, "F")) + CCur(Ndias)
'            Fecvenci = CDate(FecVenci1)
'        End If
'        'fin 01/06/2007
        
        '--fin
        
        
        '++[Monica]05/08/2011: se añaden tantos vencimientos como nos indique la forma de pago
        
        'Obtener el Nº de Vencimientos de la forma de pago
        Sql = "SELECT numerove, diasvto primerve, restoven FROM sforpa WHERE codforpa=" & DBLet(Rs!Codforpa, "N")
        Set rsVenci = New ADODB.Recordset
        rsVenci.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not rsVenci.EOF Then
            If rsVenci!numerove > 0 And DBLet(Rs!TotalFac) <> 0 Then
        
                '++[Monica]05/08/2011: si no hay fecha de vencimiento ponemos la fecha de factura, si no los calculos se hacen con la
                '                    fecha de vencimiento
                If FechaVen = "" Then
                    FechaVen = DBLet(Rs!fecfactu, "F")
                    FechaVen = DateAdd("d", DBLet(rsVenci!primerve, "N"), FechaVen)
                End If
                
                fecvenci = CDate(FechaVen)
                '++fin
        
                '-------- Primer Vencimiento
                i = 1
                'FECHA VTO
                'FecVenci = CDate(FecVenci)
                'FecVenci = DateAdd("d", DBLet(RsVenci!primerve, "N"), FechaVen)
                '===
        
                '[Monica]17/01/2013: Calculamos la nueva fecha de vencimiento si el cliente tiene dia fijo de pago
                If vsocio.DiaPago <> "" Then
                    fecvenci = NuevaFechaVto(fecvenci, vsocio.DiaPago)
                End If
                
                
                '[Monica]24/01/2013: si la factura es de tpv y la cooperativa es Ribarrojala fecha de vencimiento es la fecha de factura
                If vParamAplic.Cooperativa = 5 Then
                    LetraS = DevuelveDesdeBDNew(cPTours, "stipom", "letraser", "codtipom", "FAT", "T")
                    If LetraS = DBLet(Rs!letraser, "T") Then
                        fecvenci = DBLet(Rs!fecfactu, "F")
                   End If
                End If
                
               
               'IMPORTE del Vencimiento
                TotalFactura2 = DBLet(Rs!TotalFac, "N")
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFactura2
                Else
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFactura2 Then
                        ImpVenci = Round(ImpVenci + (TotalFactura2 - ImpVenci * rsVenci!numerove), 2)
                    End If
                End If

        
                CadValuesAux2 = "(" & DBSet(Rs!letraser, "T") & ", " & DBSet(Rs!numfactu, "N") & ", " & DBSet(Rs!fecfactu, "F") & ", "
                      
                CadValues2 = CadValuesAux2 & "1," & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(Rs!Codforpa, "N") & "," & Format(DBSet(fecvenci, "F"), FormatoFecha) & ","
              

                CodmacBPr = ""
                CodmacBPr = DevuelveDesdeBD("codmacta", "sbanco", "codbanpr", CStr(BanPr), "N")
                
                '13/02/2007
                If vsocio.TipoFactu = 0 Then ' facturacion por tarjeta
                    Select Case vTabla
                        Case "schfac"
                            Sql3 = "select numtarje from slhfac where " & cadWHERE2
                        Case "schfacr"
                            Sql3 = "select numtarje from slhfacr where " & cadWHERE2
                        Case "schfac1"
                            Sql3 = "select numtarje from slhfac1 where " & cadWHERE2
                    End Select
                    Set Rs3 = New ADODB.Recordset
                    Rs3.Open Sql3, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If Not Rs3.EOF Then
                        '[Monica]22/11/2013: Tema iban
                        Sql4 = "select codbanco, codsucur, digcontr, cuentaba, iban from starje where codsocio = " & vsocio.Codigo & " and numtarje = " & DBSet(Rs3.Fields(0).Value, "N")
                        Set Rs4 = New ADODB.Recordset
                        Rs4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        If Not Rs4.EOF Then
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(Rs4!codbanco, "N") & ", " & DBSet(Rs4!codsucur, "N") & ", " & DBSet(Rs4!digcontr, "T") & ", " & DBSet(Rs4!cuentaba, "T") & ", " & DBSet(Rs4!Iban, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(Rs4!codbanco, "N") & ", " & DBSet(Rs4!codsucur, "N") & ", " & DBSet(Rs4!digcontr, "T") & ", " & DBSet(Rs4!cuentaba, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            End If
                        Else
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.Iban, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            End If
                        End If
                    Else
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                           CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.Iban, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        Else
                           CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        End If
                    End If
        
                Else    ' facturacion por cliente
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.Iban, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                    End If
        
                End If
                CadValues2 = CadValues2 & _
                             DBSet(textcsb42, "T") & "," & DBSet(textcsb43, "T") & "," & DBSet(textcsb51, "T") & "," & DBSet(textcsb52, "T") & "," & DBSet(textcsb53, "T") & "," & DBSet(textcsb61, "T") & "," & DBSet(textcsb62, "T") & "," & DBSet(textcsb63, "T") & "," & DBSet(textcsb71, "T") & "," & _
                             DBSet(textcsb72, "T") & "," & DBSet(textcsb73, "T") & "," & DBSet(textcsb81, "T") & "," & DBSet(textcsb82, "T") & "," & DBSet(textcsb83, "T") & ", 1),"
                             
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For i = 2 To rsVenci!numerove
                   'FECHA Resto Vencimientos
                    fecvenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), fecvenci)
                    '===
                
                    '[Monica]17/01/2013: Calculamos la nueva fecha de vencimiento si el cliente tiene dia fijo de pago
                    If vsocio.DiaPago <> "" Then
                        fecvenci = NuevaFechaVto(fecvenci, vsocio.DiaPago)
                    End If
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    
                    
                    CadValues2 = CadValues2 & CadValuesAux2 & DBSet(i, "N") & "," & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(Rs!Codforpa, "N") & "," & Format(DBSet(fecvenci, "F"), FormatoFecha) & ","
                    
                    
                    '13/02/2007
                    If vsocio.TipoFactu = 0 Then ' facturacion por tarjeta
                        Select Case vTabla
                            Case "schfac"
                                Sql3 = "select numtarje from slhfac where " & cadWHERE2
                            Case "schfacr"
                                Sql3 = "select numtarje from slhfacr where " & cadWHERE2
                            Case "schfac1"
                                Sql3 = "select numtarje from slhfac1 where " & cadWHERE2
                        End Select
                        Set Rs3 = New ADODB.Recordset
                        Rs3.Open Sql3, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        
                        If Not Rs3.EOF Then
                            Sql4 = "select codbanco, codsucur, digcontr, cuentaba, iban from starje where codsocio = " & vsocio.Codigo & " and numtarje = " & DBSet(Rs3.Fields(0).Value, "N")
                            Set Rs4 = New ADODB.Recordset
                            Rs4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            If Not Rs4.EOF Then
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(Rs4!codbanco, "N") & ", " & DBSet(Rs4!codsucur, "N") & ", " & DBSet(Rs4!digcontr, "T") & ", " & DBSet(Rs4!cuentaba, "T") & ", " & DBSet(Rs4!Iban, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.Iban, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
                            End If
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.Iban, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
                        End If
            
                    Else    ' facturacion por cliente
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.Iban, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
            
                    End If
                    CadValues2 = CadValues2 & _
                                 DBSet(textcsb42, "T") & "," & DBSet(textcsb43, "T") & "," & DBSet(textcsb51, "T") & "," & DBSet(textcsb52, "T") & "," & DBSet(textcsb53, "T") & "," & DBSet(textcsb61, "T") & "," & DBSet(textcsb62, "T") & "," & DBSet(textcsb63, "T") & "," & DBSet(textcsb71, "T") & "," & _
                                 DBSet(textcsb72, "T") & "," & DBSet(textcsb73, "T") & "," & DBSet(textcsb81, "T") & "," & DBSet(textcsb82, "T") & "," & DBSet(textcsb83, "T") & ", 1),"
                        
                Next i
                         

                If vsocio.CuentaConta <> "" Then
                    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
                    'forma de pago de la factura. Sino existe insertarla
                    'vemos si existe en la conta
                    CadValuesAux2 = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", DBLet(Rs!Codforpa), "N")
                    'si no existe la forma de pago en conta, insertamos la de ariges
                    If CadValuesAux2 = "" Then
                        cadValuesAux = "tipforpa"
                        CadValuesAux2 = DevuelveDesdeBDNew(cPTours, "sforpa", "nomforpa", "codforpa", DBLet(Rs!Codforpa), "N", cadValuesAux)
                        'insertamos e sforpa de la CONTA
                        Sql = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                        Sql = Sql & " VALUES(" & DBSet(Rs!Codforpa, "N") & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
                        ConnConta.Execute Sql
                    End If
        
                    'Insertamos en la tabla scobro de la CONTA
                    Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur, digcontr, cuentaba,"
                    '[Monica]22/11/2013: Tema Iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        Sql = Sql & "iban,text33csb , text41csb,"
                    Else
                        Sql = Sql & "text33csb , text41csb,"
                    End If
                    Sql = Sql & "text42csb, text43csb, text51csb, text52csb, text53csb, text61csb, text62csb, text63csb, text71csb, text72csb, text73csb, text81csb, text82csb, text83csb,agente) "
                    Sql = Sql & " VALUES " & Mid(CadValues2, 1, Len(CadValues2) - 1)
                    ConnConta.Execute Sql
                End If
            End If
        End If

    End If

    b = True

EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreria = b
End Function



Private Sub InsertarError(Cadena As String)
Dim Sql As String

    Sql = "insert into tmperrcomprob values ('" & Cadena & "')"
    Conn.Execute Sql

End Sub


Public Function InsertarCabAsientoDia(Diario As String, Asiento As String, Fecha As String, Obs As String, caderr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    
    cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
    cad = cad & "''," & ValorNulo & "," & DBSet(Obs, "T")
    cad = "(" & cad & ")"

    'Insertar en la contabilidad
    Sql = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) "
    Sql = Sql & " VALUES " & cad
    ConnConta.Execute Sql
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabAsientoDia = False
        caderr = Err.Description
    Else
        InsertarCabAsientoDia = True
    End If
End Function


Public Function InsertarLinAsientoDia(cad As String, caderr As String) As Boolean
' el Tipo me indica desde donde viene la llamada
' tipo = 0 srecau.codmacta
' tipo = 1 scaalb.codmacta

Dim Rs As ADODB.Recordset
Dim Aux As String
Dim Sql As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

 
    Sql = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
    Sql = Sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
    Sql = Sql & " VALUES " & cad
    
    ConnConta.Execute Sql

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinAsientoDia = False
        caderr = Err.Description
    Else
        InsertarLinAsientoDia = True
    End If
End Function

Public Function ActualizarRecaudacion(cadWHERE As String, caderr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    
    Sql = "UPDATE srecau SET intconta=1 "
    Sql = Sql & " WHERE " & cadWHERE

    Conn.Execute Sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarRecaudacion = False
        caderr = Err.Description
    Else
        ActualizarRecaudacion = True
    End If
End Function

Public Sub FechasEjercicioConta(FIni As String, FFin As String)
Dim Rs As ADODB.Recordset

    On Error GoTo EFechas

    FIni = "Select fechaini,fechafin From parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open FIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        FIni = DBLet(Rs!FechaIni, "F")
        FFin = DBLet(Rs!FechaFin, "F")
    End If
    Rs.Close
    Set Rs = Nothing

EFechas:
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function CrearTMPAsiento() As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPAsiento = False
    
    Sql = "CREATE TEMPORARY TABLE tmpasien ( "
    Sql = Sql & "fecalbar date NOT NULL default '0000-00-00',"
    Sql = Sql & "codturno tinyint(1) NOT NULL default '0',"
    Sql = Sql & "codmacta varchar(10) NOT NULL default ' ',"
    Sql = Sql & "importel decimal(10,2)  NOT NULL default '0.00')"
    Conn.Execute Sql

    CrearTMPAsiento = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPAsiento = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpasien;"
        Conn.Execute Sql
    End If
End Function


Public Function TarjetasInexistentes(Sql As String) As Boolean
Dim cadMen As String

    TarjetasInexistentes = False
    
    Sql = Sql & " and not (scaalb.codsocio, scaalb.numtarje) in (select codsocio, numtarje from starje) "
    
    If (RegistrosAListar(Sql) <> 0) Then
        cadMen = "Hay cargas en las que no es correcta la tarjeta para el socio." & vbCrLf & vbCrLf & _
                 "Revise en el mantenimiento de albaranes." & vbCrLf & vbCrLf
        MsgBox cadMen, vbExclamation
        TarjetasInexistentes = True
    End If
End Function

Public Function ComprobarNumFacturas_new(cadTABLA As String, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim Sql As String
Dim SQLconta As String
Dim Rs As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas_new = False
    
'    SQLconta = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
    SQLconta = "SELECT count(*) FROM cabfact WHERE "
'    SQLconta = SQLconta & " WHERE (" & cadWConta & ") "

    
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText

'    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        Sql = "SELECT DISTINCT " & cadTABLA & ".codtipom,letraser,facturas.numfactu,facturas.fecfactu "
        Sql = Sql & " FROM (" & cadTABLA & " INNER JOIN usuarios.stipom stipom ON " & cadTABLA & ".codtipom=stipom.codtipom) "
        Sql = Sql & " INNER JOIN tmpFactu ON facturas.codtipom=tmpFactu.codtipom AND facturas.numfactu=tmpFactu.numfactu AND facturas.fecfactu=tmpFactu.fecfactu "

        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        b = True
        While Not Rs.EOF And b
            Sql = "(numserie= " & DBSet(Rs!letraser, "T") & " AND codfaccl=" & DBSet(Rs!numfactu, "N") & " AND anofaccl=" & Year(Rs!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            Sql = SQLconta & Sql
            If RegistrosAListar(Sql, cConta) Then
                b = False
                Sql = "          Letra Serie: " & DBSet(Rs!letraser, "T") & vbCrLf
                Sql = Sql & "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                Sql = Sql & "          Fecha: " & Rs!fecfactu
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not b Then
            Sql = "Ya existe la factura: " & vbCrLf & Sql
            Sql = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & Sql
            
            MsgBox Sql, vbExclamation
            ComprobarNumFacturas_new = False
        Else
            ComprobarNumFacturas_new = True
        End If
'    Else
'        ComprobarNumFacturas_new = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
    Exit Function
    
ECompFactu:
     If Err.Number <> 0 Then
        ComprobarNumFacturas_new = False
        MuestraError Err.Number, "Comprobar Nº Facturas", Err.Description
    End If
End Function

Public Function ComprobarCtaContable_new(cadTABLA As String, Opcion As Byte) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim Sql As String
Dim Rs As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim SQLcuentas As String
Dim CadCampo1 As String
Dim numNivel As String
Dim NumDigit As String
Dim NumDigit3 As String


    On Error GoTo ECompCta

    ComprobarCtaContable_new = False
    
    cadG = ""
    If Opcion = 3 Or Opcion = 7 Or Opcion = 10 Or Opcion = 13 Then
        'si hay analitica comprobar que todas las cuentas
        'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
        cadG = "grupovta"
        Sql = DevuelveDesdeBDNew(cConta, "parametros", "grupogto", "", "", "", cadG)
        If Sql <> "" And cadG <> "" Then
            Sql = " AND (codmacta like '" & Sql & "%' OR codmacta like '" & cadG & "%')"
        ElseIf Sql <> "" Then
            Sql = " AND (codmacta like '" & Sql & "%')"
        ElseIf cadG <> "" Then
            Sql = " AND (codmacta like '" & cadG & "%')"
        End If
        cadG = Sql
    End If
    
    
'    SQL = "SELECT codmacta FROM cuentas "
'    SQL = SQL & " WHERE apudirec='S'"
'    If cadG <> "" Then SQL = SQL & cadG
    
    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG
    
    If Opcion = 1 Then
        If cadTABLA = "facturas" Then
            'Seleccionamos los distintos clientes,cuentas que vamos a facturar
            Sql = "SELECT DISTINCT facturas.codclien, clientes.codmacta "
            Sql = Sql & " FROM (facturas INNER JOIN clientes ON facturas.codclien=clientes.codclien) "
            Sql = Sql & " INNER JOIN tmpFactu ON facturas.codtipom=tmpFactu.codtipom AND facturas.numfactu=tmpFactu.numfactu AND facturas.fecfactu=tmpFactu.fecfactu "
        Else
            If cadTABLA = "scafpc" Then
                'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
                Sql = "SELECT DISTINCT scafpc.codprove, proveedor.codmacta "
                Sql = Sql & " FROM (scafpc INNER JOIN proveedor ON scafpc.codprove=proveedor.codprove) "
                Sql = Sql & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
            Else
                'Seleccionamos los distintos transportistas ,cuentas que vamos a facturar
                Sql = "SELECT DISTINCT tcafpc.codtrans, agencias.codmacta "
                Sql = Sql & " FROM (tcafpc INNER JOIN agencias ON tcafpc.codtrans=agencias.codtrans) "
                Sql = Sql & " INNER JOIN tmpFactu ON tcafpc.codtrans=tmpFactu.codtrans AND tcafpc.numfactu=tmpFactu.numfactu AND tcafpc.fecfactu=tmpFactu.fecfactu "
            
            End If
        End If
    ElseIf Opcion = 2 Or Opcion = 3 Or Opcion = 8 Then
        Sql = "SELECT distinct "
        If Opcion = 2 Then Sql = Sql & " sartic.codartic,"
        If cadTABLA = "facturas" Then
            If Opcion <> 8 Then
                Sql = Sql & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((facturas_envases "
                Sql = Sql & " INNER JOIN tmpFactu ON facturas_envases.codtipom=tmpFactu.codtipom AND facturas_envases.numfactu=tmpFactu.numfactu AND facturas_envases.fecfactu=tmpFactu.fecfactu) "
                Sql = Sql & "INNER JOIN sartic ON facturas_envases.codartic=sartic.codartic) "
            Else
                numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", vParamAplic.NumeroConta, "N")
                NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel, "codempre", vParamAplic.NumeroConta, "N")
                NumDigit3 = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
                
'                CadCampo1 = "concat(concat(variedades.raizctavtas,tipomer.digicont), right(concat('0000000000',albaran_variedad.codvarie)," & (CCur(NumDigit) - CCur(NumDigit3) - 1) & "))"
                CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                
                Sql = Sql & " albaran_variedad.codvarie, " & CadCampo1 & " as codmacta from ((((((facturas_variedad "
                Sql = Sql & " INNER JOIN tmpFactu ON facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu) "
                Sql = Sql & " inner join usuarios.stipom stipom on facturas_variedad.codtipom=stipom.codtipom) "
                Sql = Sql & " inner join albaran on facturas_variedad.numalbar = albaran.numalbar) "
                Sql = Sql & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
                Sql = Sql & " inner join albaran_variedad on facturas_variedad.numalbar = albaran_variedad.numalbar and facturas_variedad.numlinealbar = albaran_variedad.numlinea) "
                Sql = Sql & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
                
                
'                Sql = Sql & " INNER JOIN tmpFactu ON facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu) "
'                Sql = Sql & "INNER JOIN sartic ON facturas_envases.codartic=sartic.codartic) "
            End If
        Else
            Sql = Sql & " sartic.ctacompr as codmacta from ((slifpc "
            Sql = Sql & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
            Sql = Sql & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
        End If
'        If Opcion <> 8 Then Sql = Sql & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
    ElseIf Opcion = 4 Or Opcion = 6 Then
'        Sql = "select distinct " & DBSet(vParamAplic.CtaTraReten, "T") & " as codmacta from tcafpc "
    ElseIf Opcion = 5 Or Opcion = 7 Then
'        Sql = "select distinct " & DBSet(vParamAplic.CtaAboTrans, "T") & " as codmacta from tcafpc "
'       transporte
            Sql = " SELECT if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta "
            Sql = Sql & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer, tmpFactu, tcafpc  WHERE "
            Sql = Sql & " tcafpc.tipo = 0 and " ' transportista
            Sql = Sql & " tlifpc.codtrans=tmpFactu.codtrans and tlifpc.numfactu=tmpFactu.numfactu and tlifpc.fecfactu=tmpFactu.fecfactu and "
            Sql = Sql & " tlifpc.numalbar=albaran_variedad.numalbar and "
            Sql = Sql & " tlifpc.numlinea=albaran_variedad.numlinea and "
            Sql = Sql & " tlifpc.codtrans=tcafpc.codtrans and tlifpc.numfactu=tcafpc.numfactu and tlifpc.fecfactu=tcafpc.fecfactu and "
            Sql = Sql & " albaran_variedad.numalbar=albaran.numalbar and "
            Sql = Sql & " albaran_variedad.codvarie=variedades.codvarie and "
            Sql = Sql & " albaran.codtimer=tipomer.codtimer "
            Sql = Sql & " group by 1 "

    ElseIf Opcion = 12 Or Opcion = 13 Then
'       comisionista
            Sql = " SELECT variedades.ctacomisionista as cuenta, variedades.codvarie  "
            Sql = Sql & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer, tmpFactu, tcafpc  WHERE "
            Sql = Sql & " tcafpc.tipo = 1 and " ' comisionista
            Sql = Sql & " tlifpc.codtrans=tmpFactu.codtrans and tlifpc.numfactu=tmpFactu.numfactu and tlifpc.fecfactu=tmpFactu.fecfactu and "
            Sql = Sql & " tlifpc.numalbar=albaran_variedad.numalbar and "
            Sql = Sql & " tlifpc.numlinea=albaran_variedad.numlinea and "
            Sql = Sql & " tlifpc.codtrans=tcafpc.codtrans and tlifpc.numfactu=tcafpc.numfactu and tlifpc.fecfactu=tcafpc.fecfactu and "
            Sql = Sql & " albaran_variedad.numalbar=albaran.numalbar and "
            Sql = Sql & " albaran_variedad.codvarie=variedades.codvarie and "
            Sql = Sql & " albaran.codtimer=tipomer.codtimer "
            Sql = Sql & " group by 1 "
            
    ElseIf Opcion = 9 Or Opcion = 10 Then
            Sql = " select codmacta as cuenta "
            Sql = Sql & " from tcafpv, tmpFactu "
            Sql = Sql & " where tmpFactu.codtrans=tcafpv.codtrans and tmpFactu.numfactu=tcafpv.numfactu and tmpFactu.fecfactu=tcafpv.fecfactu "
            Sql = Sql & " group by 1 "
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    b = True

    While Not Rs.EOF And b
        If Opcion < 4 Or Opcion = 8 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Codmacta, "T")
        ElseIf Opcion = 4 Or Opcion = 6 Then
'            Sql = SQLcuentas & " AND codmacta= " & DBSet(vParamAplic.CtaTraReten, "T")
        ElseIf Opcion = 5 Or Opcion = 7 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!cuenta, "T")
        ElseIf Opcion = 12 Or Opcion = 13 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!cuenta, "T")
        ElseIf Opcion = 9 Or Opcion = 10 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!cuenta, "T")
        End If
            
        
        If Not (RegistrosAListar(Sql, cConta) > 0) Then
        'si no lo encuentra
            b = False 'no encontrado
            If Opcion = 1 Then
                If cadTABLA = "facturas" Then
                    Sql = Rs!Codmacta & " del Cliente " & Format(Rs!CodClien, "000000")
                Else
                    If cadTABLA = "scafpc" Then
                        Sql = Rs!Codmacta & " del Proveedor " & Format(Rs!CodProve, "000000")
                    Else
                        Sql = Rs!Codmacta & " del Transportista " & Format(Rs!codTrans, "000")
                    End If
                End If
            ElseIf Opcion = 2 Then
                Sql = Rs!Codmacta & " del articulo " & Format(Rs!codArtic, "000000")
            ElseIf Opcion = 3 Then
                Sql = Rs!Codmacta
            ElseIf Opcion = 4 Or Opcion = 6 Then
'                Sql = vParamAplic.CtaTraReten
            ElseIf Opcion = 5 Or Opcion = 7 Then
                Sql = DBLet(Rs!cuenta, "T") ' vParamAplic.CtaAboTrans
            ElseIf Opcion = 12 Or Opcion = 13 Then
                Sql = DBLet(Rs!cuenta, "T") & " de comisionista de la variedad " & Format(Rs!codvarie, "000000")
            ElseIf Opcion = 8 Then
                Sql = Rs!Codmacta & " de la variedad " & Format(Rs!codvarie, "0000")
            ElseIf Opcion = 9 Or Opcion = 10 Then
                Sql = DBLet(Rs!cuenta, "T") ' vParamAplic.CtaAboTrans
            End If
        End If
        
        
'        If Opcion = 2 Or Opcion = 3 Then
'            'Comprobar que ademas de existir la cuenta de ventas exista tambien
'            'la cuenta ABONO ventas (sfamia.aboventa)
'            '---------------------------------------------
'            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!ctaabono, "T")
''            RSconta.MoveFirst
''            RSconta.Find (SQL), , adSearchForward
''            If RSconta.EOF Then
'            If Not (RegistrosAListar(Sql, cConta) > 0) Then
'                b = False 'no encontrado
'                If Opcion = 2 Then
'                    Sql = Rs!ctaabono & " de la familia " & Format(Rs!codfamia, "0000")
'                ElseIf Opcion = 3 Then
'                    Sql = Rs!ctaabono
'                End If
'            End If
'
'
'            'comprobar cuentas alternativas solo para facturacion a CLIENTES
'            '----------------------------------------------------------------
'            If cadTABLA = "facturas" Then
'                ' Comprobar cuenta VENTA alternativa
'                If DBLet(Rs!ctavent1, "T") <> "" Then
'                    Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!ctavent1, "T")
''                    RSconta.MoveFirst
''                    RSconta.Find (SQL), , adSearchForward
''                    If RSconta.EOF Then
'                    If Not (RegistrosAListar(Sql, cConta) > 0) Then
'                        b = False 'no encontrado
'                        If Opcion = 2 Then
'                            Sql = Rs!ctavent1 & " de la familia " & Format(Rs!codfamia, "0000")
'                        ElseIf Opcion = 3 Then
'                            Sql = Rs!ctavent1
'                        End If
'                    End If
'                Else
'                    b = False
'                    Sql = " o la familia no tiene asignada cuenta venta alternativa."
'                End If
'
'                ' Comprobar cuenta de ABONO alternativa
'                If DBLet(Rs!abovent1, "T") <> "" Then
'                    Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!abovent1, "T")
''                    RSconta.MoveFirst
''                    RSconta.Find (SQL), , adSearchForward
''                    If RSconta.EOF Then
'                    If Not (RegistrosAListar(Sql, cConta) > 0) Then
'                        b = False 'no encontrado
'                        If Opcion = 2 Then
'                            Sql = Rs!abovent1 & " de la familia " & Format(Rs!codfamia, "0000")
'                        ElseIf Opcion = 3 Then
'                            Sql = Rs!abovent1
'                        End If
'                    End If
'                Else
'                    b = False
'                    Sql = " o la familia no tiene asignada cuenta abono alternativa."
'                End If
'            End If
'
'        End If
'
        Rs.MoveNext
    Wend
    
    
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

'    If Not RSconta.EOF Then
'        If Opcion = 1 Then
'            If cadTabla = "scafac" Then
'                'Seleccionamos los distintos clientes,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafac.codclien, sclien.codmacta "
'                SQL = SQL & " FROM (scafac INNER JOIN sclien ON scafac.codclien=sclien.codclien) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
'            Else
'                'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafpc.codprove, sprove.codmacta "
'                SQL = SQL & " FROM (scafpc INNER JOIN sprove ON scafpc.codprove=sprove.codprove) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
'            End If

'        ElseIf Opcion = 2 Or Opcion = 3 Then
'            SQL = "SELECT distinct "
'            If Opcion = 2 Then SQL = SQL & " sartic.codfamia,"
'            If cadTabla = "scafac" Then
'                SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((slifac "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifac.codtipom=tmpFactu.codtipom AND slifac.numfactu=tmpFactu.numfactu AND slifac.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifac.codartic=sartic.codartic) "
'            Else
'                SQL = SQL & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
'            End If
'            SQL = SQL & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
'        End If
        
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        b = True
'        While Not RS.EOF And b
'            SQL = "codmacta= " & DBSet(RS!Codmacta, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
'                b = False 'no encontrado
'                If Opcion = 1 Then
'                    If cadTabla = "scafac" Then
'                        SQL = RS!Codmacta & " del Cliente " & Format(RS!CodClien, "000000")
'                    Else
'                        SQL = RS!Codmacta & " del Proveedor " & Format(RS!codProve, "000000")
'                    End If
'                ElseIf Opcion = 2 Then
'                    SQL = RS!Codmacta & " de la familia " & Format(RS!codfamia, "0000")
'                ElseIf Opcion = 3 Then
'                    SQL = RS!Codmacta
'                End If
'            End If
            
'            If Opcion = 2 Then
'                'Comprobar que ademas de existir la cuenta de ventas exista tambien
'                'la cuenta ABONO ventas
'                SQL = "codmacta= " & DBSet(RS!ctaabono, "T")
'                RSconta.MoveFirst
'                RSconta.Find (SQL), , adSearchForward
'                If RSconta.EOF Then
'                    b = False 'no encontrado
'
'                    SQL = RS!ctaabono & " de la familia " & Format(RS!codfamia, "0000")
'                End If
'            End If
            
            'comprobar cuentas alternativas solo para facturacion a clientes
'            If cadTabla = "scafac" Then
'                If Opcion = 2 Then
'                    ' Comprobar cuenta venta alternativa
'                    If DBLet(RS!ctavent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!ctavent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctavent1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta venta alternativa."
'                    End If
'                End If
'                If Opcion = 2 Then
'                    ' Comprobar cuenta de abono alternativa
'                    If DBLet(RS!abovent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!abovent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctaabon1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta abono alternativa."
'                    End If
'                End If
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
        
        
        
        If Not b Then
            If Not (Opcion = 3 Or Opcion = 6 Or Opcion = 7) Then
                Sql = "No existe la cta contable " & Sql
            Else
                Sql = "La cuenta " & Sql & " no es del nivel correcto. "
                If Opcion = 3 Then Sql = Sql & "(Familias de artículos)."
            End If
            Sql = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & Sql
            
            MsgBox Sql, vbExclamation
            ComprobarCtaContable_new = False
        Else
            ComprobarCtaContable_new = True
        End If
'    Else
'        ComprobarCtaContable_new = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
    Exit Function
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function




Public Function ComprobarCCoste_new(cadCC As String, cadTABLA As String, Optional Opcion As Byte) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECCoste

    ComprobarCCoste_new = False
    Select Case cadTABLA
        Case "facturas" ' facturas de venta
            Select Case Opcion
                Case 1
                    Sql = "select distinct variedades.codccost from facturas_variedad, albaran_variedad, variedades, tmpFactu where "
                    Sql = Sql & " albaran_variedad.codvarie=variedades.codvarie and "
                    Sql = Sql & " facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu and  "
                    Sql = Sql & " albaran_variedad.numalbar = facturas_variedad.numalbar and "
                    Sql = Sql & " albaran_variedad.numlinea = facturas_variedad.numlinealbar "
                Case 2
                    Sql = " select distinct sfamia.codccost from facturas_envases, sartic, sfamia, tmpFactu where "
                    Sql = Sql & " facturas_envases.codtipom=tmpFactu.codtipom AND facturas_envases.numfactu=tmpFactu.numfactu AND facturas_envases.fecfactu=tmpFactu.fecfactu and  "
                    Sql = Sql & " facturas_envases.codartic = sartic.codartic and "
                    Sql = Sql & " sartic.codfamia = sfamia.codfamia "
                Case 3
'                    If HayFacturasACuenta Then
'                        Sql = " select '" & vParamAplic.CCosteFraACta & "' as codccost from tmpFactu where tmpfactu.codtipom = 'EAC' "
'                    Else
'                        ComprobarCCoste_new = True
'                        Exit Function
'                    End If
            End Select
        Case "scafpc" ' facturas de compra
            Sql = " select distinct sfamia.codccost from slifpc, sartic, sfamia, tmpFactu where "
            Sql = Sql & " slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu and  "
            Sql = Sql & " slifpc.codartic = sartic.codartic and "
            Sql = Sql & " sartic.codfamia = sfamia.codfamia "
        
        Case "tcafpc" ' facturas de transporte
            Sql = "select distinct variedades.codccost from tlifpc, albaran_variedad, variedades, tmpFactu where "
            Sql = Sql & " albaran_variedad.codvarie=variedades.codvarie and "
            Sql = Sql & " tlifpc.codtrans=tmpFactu.codtrans AND tlifpc.numfactu=tmpFactu.numfactu AND tlifpc.fecfactu=tmpFactu.fecfactu and  "
            Sql = Sql & " albaran_variedad.numalbar = tlifpc.numalbar and "
            Sql = Sql & " albaran_variedad.numlinea = tlifpc.numlinea "
    
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = True

    While Not Rs.EOF And b
        If DBLet(Rs.Fields(0).Value, "T") = "" Then
            b = False
        Else
            Sql = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", Rs.Fields(0).Value, "T")
            If Sql = "" Then
                b = False
                Sql2 = "Centro de Coste: " & Rs.Fields(0)
            End If
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    If Not b Then
        Sql = "No existe el " & Sql2
        Sql = "Comprobando Centros de Coste en contabilidad..." & vbCrLf & vbCrLf & Sql
    
        MsgBox Sql, vbExclamation
        ComprobarCCoste_new = False
        Exit Function
    Else
        ComprobarCCoste_new = True
    End If
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Centros de Coste", Err.Description
    End If
End Function

Public Function ComprobarFormadePago(cadCC As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ECCoste

    ComprobarFormadePago = False
    Sql = "select distinct facturas.codforpa from facturas, tmpFactu where "
    Sql = Sql & " facturas.codtipom=tmpFactu.codtipom AND facturas.numfactu=tmpFactu.numfactu AND facturas.fecfactu=tmpFactu.fecfactu  "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = True

    While Not Rs.EOF And b
        Sql = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", Rs.Fields(0).Value, "N")
        If Sql = "" Then
            b = False
            Sql2 = "Formas de Pago: " & Rs.Fields(0)
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    If Not b Then
        Sql = "No existe la " & Sql2
        Sql = "Comprobando Formas de Pago en contabilidad..." & vbCrLf & vbCrLf & Sql
    
        MsgBox Sql, vbExclamation
        ComprobarFormadePago = False
        Exit Function
    Else
        ComprobarFormadePago = True
    End If
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Formas de Pago", Err.Description
    End If
End Function



Public Function PasarFacturaProv(cadWHERE As String, CodCCost As String, FechaFin As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariges.scafpc --> conta.cabfactprov
' ariges.slifpc --> conta.linfactprov
'Actualizar la tabla ariges.scafpc.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim Sql As String
Dim Mc As Contadores


    On Error GoTo EContab

    ConnConta.BeginTrans
    Conn.BeginTrans
        
    
    Set Mc = New Contadores
    
    '---- Insertar en la conta Cabecera Factura
    b = InsertarCabFactProv(cadWHERE, cadMen, Mc, FechaFin)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If b Then
        CCoste = CodCCost
        '---- Insertar lineas de Factura en la Conta
        b = InsertarLinFact_new("scafpc", cadWHERE, cadMen, Mc.Contador)
        cadMen = "Insertando Lin. Factura: " & cadMen

        If b Then
            '---- Poner intconta=1 en ariges.scafac
            b = ActualizarCabFact("scafpc", cadWHERE, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
'    If Not b Then
'        SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'        SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'        Conn.Execute SQL
'    End If
    
EContab:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        Conn.CommitTrans
        PasarFacturaProv = True
    Else
        ConnConta.RollbackTrans
        Conn.RollbackTrans
        PasarFacturaProv = False
        If Not b Then
            InsertarTMPErrFac cadMen, cadWHERE
'            SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'            SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'            SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'            Conn.Execute SQL
        End If
    End If
End Function


Private Function InsertarCabFactProv(cadWHERE As String, caderr As String, ByRef Mc As Contadores, FechaFin As String) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Intracom As Integer

    On Error GoTo EInsertar
       
    
    Sql = " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,proveedor.codmacta,"
    Sql = Sql & "scafpc.dtoppago,scafpc.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "totalfac,tipoiva1,tipoiva2,tipoiva3,proveedor.codprove, proveedor.nomprove, proveedor.tipprove "
    Sql = Sql & " FROM " & "scafpc "
    Sql = Sql & "INNER JOIN " & "proveedor ON scafpc.codprove=proveedor.codprove "
    Sql = Sql & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
    
        If Mc.ConseguirContador("1", (Rs!FecRecep <= CDate(FechaFin) - 365), True) = 0 Then
        
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = Rs!DtoPPago
            DtoGnral = Rs!DtoGnral
            BaseImp = Rs!BaseIVA1 + CCur(DBLet(Rs!BaseIVA2, "N")) + CCur(DBLet(Rs!BaseIVA3, "N"))
            TotalFac = Rs!TotalFac
            AnyoFacPr = Rs!anofacpr
            
            Intracom = DBLet(Rs!tipprove, "N")
            If Intracom = 2 Then Intracom = 0
            
            Nulo2 = "N"
            Nulo3 = "N"
            If DBLet(Rs!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(Rs!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            Sql = ""
            Sql = Mc.Contador & "," & DBSet(Rs!fecfactu, "F") & "," & Rs!anofacpr & "," & DBSet(Rs!FecRecep, "F") & "," & DBSet(Rs!numfactu, "T") & "," & DBSet(Rs!Codmacta, "T") & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!BaseIVA1, "N") & "," & DBSet(Rs!BaseIVA2, "N", "S") & "," & DBSet(Rs!BaseIVA3, "N", "S") & ","
            Sql = Sql & DBSet(Rs!porciva1, "N") & "," & DBSet(Rs!porciva2, "N", Nulo2) & "," & DBSet(Rs!porciva3, "N", Nulo3) & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & DBSet(Rs!impoiva2, "N", Nulo2) & "," & DBSet(Rs!impoiva3, "N", Nulo3) & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!TipoIVA1, "N") & "," & DBSet(Rs!TipoIVA2, "N", Nulo2) & "," & DBSet(Rs!TipoIVA3, "N", Nulo3) & "," & DBSet(Intracom, "N") & ","
            Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Rs!FecRecep, "F") & ",0"
            cad = cad & "(" & Sql & ")"
            
            'Insertar en la contabilidad
            Sql = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
            Sql = Sql & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
            Sql = Sql & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
            Sql = Sql & " VALUES " & cad
            ConnConta.Execute Sql
            
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            Sql = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vSesion.Codigo & "," & Mc.Contador
            Sql = Sql & ",'" & DevNombreSQL(Rs!numfactu) & " @ " & Format(Rs!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(Rs!NomProve) & "'," & Rs!CodProve & ")"
            Conn.Execute Sql
            
            
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactProv = False
        caderr = Err.Description
    Else
        InsertarCabFactProv = True
    End If
End Function



Private Function InsertarLinFact_new(cadTABLA As String, cadWHERE As String, caderr As String, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim SqlAux As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim CadCampo1 As String
Dim numNivel As String
Dim NumDigit As String
Dim NumDigitAnt As String
Dim NumDigit3 As String
Dim Tipo As Byte
Dim TipoFact As String
Dim TieneAnalitica As String

    On Error GoTo EInLinea
    

    If cadTABLA = "scafpc" Then 'COMPRAS
        'utilizamos sfamia.ctaventa o sfamia.aboventa
        If TotalFac >= 0 Then
            cadCampo = "sartic.ctacompr"
        Else
            cadCampo = "sartic.ctacompr"
        End If
        TieneAnalitica = "0"
        TieneAnalitica = DevuelveDesdeBDNew(cConta, "parametros", "autocoste", "", "")
        If TieneAnalitica = "1" Then  'hay contab. analitica
            Sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sartic.codccost"
        Else
            Sql = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
        End If
        Sql = Sql & " FROM (slifpc  "
        Sql = Sql & " inner join sartic on slifpc.codartic=sartic.codartic) "
        Sql = Sql & " WHERE " & Replace(cadWHERE, "scafpc", "slifpc")
        Sql = Sql & " GROUP BY " & cadCampo
        
        If TieneAnalitica = "1" Then
            Sql = Sql & ", sartic.codccost "
        End If
    End If
  
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    i = 1
    totimp = 0
    SqlAux = ""
    While Not Rs.EOF
        SqlAux = cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = Rs!Importe - CCur(CalcularPorcentaje(Rs!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(Rs!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql2 = ""
        
        If cadTABLA = "facturas" Then 'VENTAS a clientes
            Sql = "'" & Rs!letraser & "'," & Rs!numfactu & "," & Year(Rs!fecfactu) & "," & i & ","
            Sql = Sql & DBSet(Rs!cuenta, "T")
'            If Not conCtaAlt Then 'cliente no tiene cuenta alternativa
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctaventa, "T")
'                Else
'                    SQL = SQL & DBSet(RS!aboventa, "T")
'                End If
'            Else
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctavent1, "T")
'                Else
'                    SQL = SQL & DBSet(RS!abovent1, "T")
'                End If
'            End If
        Else
            If cadTABLA = "scafpc" Then 'COMPRAS
                'Laura 24/10/2006
                'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
                Sql = numRegis & "," & AnyoFacPr & "," & i & ","
                
    '            If ImpLinea >= 0 Then
                    Sql = Sql & DBSet(Rs!cuenta, "T")
    '            Else
    '                SQL = SQL & DBSet(RS!abocompr, "T")
    '            End If
            Else 'TRANSPORTE
                Sql = numRegis & "," & AnyoFacPr & "," & i & ","
                Sql = Sql & DBSet(Rs!cuenta, "T")
            End If
        End If
        
        Sql2 = Sql & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        Sql = Sql & "," & DBSet(ImpLinea, "N") & ","
        
        If TieneAnalitica = "1" Then
            If cadTABLA = "tcafpc" Then
                If DBLet(Rs!CodCCost, "T") = "----" Then
                    Sql = Sql & DBSet(CCoste, "T")
                Else
                    Sql = Sql & DBSet(Rs!CodCCost, "T")
                    CCoste = DBLet(Rs!CodCCost, "T")
                End If
            Else
                Sql = Sql & DBSet(Rs!CodCCost, "T")
                CCoste = DBLet(Rs!CodCCost, "T")
            End If
        Else
            Sql = Sql & ValorNulo
            CCoste = ValorNulo
        End If
        
        cad = cad & "(" & Sql & ")" & ","
        
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Sql2 = Sql2 & DBSet(totimp, "N") & ","
        If CCoste = "" Or CCoste = ValorNulo Then
            Sql2 = Sql2 & ValorNulo
        Else
            Sql2 = Sql2 & DBSet(CCoste, "T")
        End If
        If SqlAux <> "" Then 'hay mas de una linea
            cad = SqlAux & "(" & Sql2 & ")" & ","
        Else 'solo una linea
            cad = "(" & Sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        If cadTABLA = "facturas" Then
            Sql = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            Sql = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        Sql = Sql & " VALUES " & cad
        ConnConta.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_new = False
        caderr = Err.Description
    Else
        InsertarLinFact_new = True
    End If
End Function


Private Sub InsertarTMPErrFac(MenError As String, cadWHERE As String)
Dim Sql As String

    On Error Resume Next
    Sql = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    Sql = Sql & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    Sql = Sql & " WHERE " & Replace(cadWHERE, "scafpc", "tmpFactu")
    Conn.Execute Sql
    
    If Err.Number <> 0 Then Err.Clear
End Sub



' ### [Monica] 02/10/2006
' copiado de la clase de laura cfactura
Public Function InsertarEnTesoreriaDB(db As BaseDatos, cadWHERE As String, ByVal fecvenci As String, BanPr As String, MenError As String, ByRef vsocio As CSocio, vTabla As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim RSx As ADODB.Recordset
Dim Sql As String, textcsb33 As String, textcsb41 As String
Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Sql5 As String
Dim Rs3 As ADODB.Recordset
Dim Rs4 As ADODB.Recordset
Dim Rs5 As ADODB.Recordset

Dim textcsb42 As String, textcsb43 As String
Dim textcsb51 As String, textcsb52 As String, textcsb53 As String
Dim textcsb61 As String, textcsb62 As String, textcsb63 As String
Dim textcsb71 As String, textcsb72 As String, textcsb73 As String
Dim textcsb81 As String, textcsb82 As String, textcsb83 As String
Dim n_linea As Integer
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci1 As Date
Dim ImpVenci As Single
Dim i As Byte
Dim CodmacBPr As String
Dim cadWHERE2 As String

Dim FacturaFP As String

Dim ForPago As String
Dim Ndias As String

    On Error GoTo EInsertarTesoreria

    b = False
    InsertarEnTesoreriaDB = False
    CadValues = ""
    CadValues2 = ""

    Sql = "select * from " & vTabla & " where  " & cadWHERE
    
    Set Rs = db.cursor(Sql)
    
    If Not Rs.EOF Then
    
        textcsb33 = "'FACT: " & DBLet(Rs!letraser, "T") & "-" & Format(DBLet(Rs!numfactu, "N"), "0000000") & " " & Format(DBLet(Rs!fecfactu, "F"), "dd/mm/yy")
        textcsb33 = textcsb33 & " de " & DBSet(Rs!TotalFac, "N") & "'"
        ' añadido 07022007
'        textcsb41 = "'B.IMP " & DBSet(RS!baseimp1, "N") & " IVA " & DBSet(RS!impoiva1, "N") & " TOTAL " & DBSet(RS!TOTALFAC, "N") & "',"
        ' end del añadido
        
        ' añadido 08022007
        textcsb41 = ""
        textcsb42 = ""
        textcsb43 = ""
        textcsb51 = ""
        textcsb52 = ""
        textcsb53 = ""
        textcsb61 = ""
        textcsb62 = ""
        textcsb63 = ""
        textcsb71 = ""
        textcsb72 = ""
        textcsb73 = ""
        textcsb81 = ""
        textcsb82 = ""
        textcsb83 = ""
        
        
        
        
'[Monica]22/11/2013: quitamos los csbs
        If vTabla = "schfac" Then
            cadWHERE2 = Replace(cadWHERE, "schfac", "slhfac")
        Else
            cadWHERE2 = Replace(cadWHERE, "schfacr", "slhfacr")
        End If


'        Sql = "select count(distinct numalbar) from " & vTabla & " where " & cadWHERE
'        If vTabla = "schfac" Then
'            Sql = Replace(Sql, "schfac", "slhfac")
'            cadWHERE2 = Replace(cadWHERE, "schfac", "slhfac")
'        Else
'            Sql = Replace(Sql, "schfacr", "slhfacr")
'            cadWHERE2 = Replace(cadWHERE, "schfacr", "slhfacr")
'        End If
'        If TotalRegistros(Sql) <= 15 Then
''            cadwhere2 = Replace(cadwhere, "schfac", "slhfac")
'            Sql2 = "select numalbar, fecalbar, sum(implinea) "
'            If vTabla = "schfac" Then
'                Sql2 = Sql2 & " from slhfac where " & cadWHERE2
'            Else
'                Sql2 = Sql2 & " from slhfacr where " & cadWHERE2
'            End If
'
'            Sql2 = Sql2 & " group by numalbar, fecalbar order by numalbar, fecalbar "
'            Set RSx = New ADODB.Recordset
'            RSx.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            n_linea = 0
'            While Not RSx.EOF
'                n_linea = n_linea + 1
'                Select Case n_linea
'                    Case 1
'                        textcsb41 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 2
'                        textcsb42 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 3
'                        textcsb43 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 4
'                        textcsb51 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 5
'                        textcsb52 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 6
'                        textcsb53 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 7
'                        textcsb61 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 8
'                        textcsb62 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 9
'                        textcsb63 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 10
'                        textcsb71 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 11
'                        textcsb72 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 12
'                        textcsb73 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 13
'                        textcsb81 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 14
'                        textcsb82 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                    Case 15
'                        textcsb83 = "TICKET " & DBLet(RSx.Fields(0).Value, "T") & " DE " & Format(RSx.Fields(1).Value, "DD-MM") & " IMPORTE " & DBSet(RSx.Fields(2).Value, "N")
'                End Select
'
'
'                RSx.MoveNext
'            Wend
'        End If
'        ' end del añadido 08022007
        
'[Monica]08/01/2014: lo cambiamos rellenando lo maximo que podemos
        If vParamAplic.Cooperativa = 5 Then
            Dim cad1 As String
            Dim cad2 As String
            Dim cad22 As String
            
            Sql = "select count(distinct numalbar) from " & vTabla & " where " & cadWHERE
            cad1 = ""
            Sql2 = "select numalbar, fecalbar, sum(implinea) "
            Select Case vTabla
                Case "schfac"
                    Sql2 = Sql2 & " from slhfac where " & cadWHERE2
                Case "schfacr"
                    Sql2 = Sql2 & " from slhfacr where " & cadWHERE2
                Case "schfac1"
                    Sql2 = Sql2 & " from slhfac1 where " & cadWHERE2
            End Select

            Sql2 = Sql2 & " group by numalbar, fecalbar order by numalbar, fecalbar "
            
            Set RSx = New ADODB.Recordset
            RSx.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            n_linea = 0
            cad2 = " "
            cad22 = ""
            While Not RSx.EOF
                n_linea = n_linea + 1
            
                cad1 = "T-" & Right("        " & DBLet(RSx.Fields(0).Value, "T"), 8) & " " & Format(DBLet(RSx.Fields(2).Value, "N"), "##0.00") & "€ "
                                
                If n_linea <= 2 Then
                    cad2 = cad2 & cad1
                Else
                    cad22 = cad22 & cad1
                End If
                RSx.MoveNext
            Wend
            If cad2 <> "" Then textcsb33 = textcsb33 & cad2
            textcsb41 = Mid(cad22, 1, InStrRev(Mid(cad22, 1, 40), "€"))
            If Len(cad22) > 40 Then textcsb41 = textcsb41 & "..."
        End If
        
        'monica 01/06/2007
        FacturaFP = ""
        FacturaFP = DevuelveDesdeBDNew(cPTours, "ssocio", "facturafp", "codsocio", Rs!codsocio, "N")
        If CInt(FacturaFP) = 1 Then
            Ndias = ""
            Ndias = DevuelveDesdeBDNew(cPTours, "sforpa", "diasvto", "codforpa", Rs!Codforpa, "N")
            Ndias = ComprobarCero(Ndias)
            FecVenci1 = CDate(DBLet(Rs!fecfactu, "F")) + CCur(Ndias)
            fecvenci = CDate(FecVenci1)
        End If
        'fin 01/06/2007
        
        CadValuesAux2 = "(" & DBSet(Rs!letraser, "T") & ", " & DBSet(Rs!numfactu, "N") & ", " & DBSet(Rs!fecfactu, "F") & ", "
              
        CadValues2 = CadValuesAux2 & "1," & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(Rs!Codforpa, "N") & "," & Format(DBSet(fecvenci, "F"), FormatoFecha) & ","
              
' 01/06/2006 he quitado esta instruccion
'        'FECHA VTO
'        FecVenci = CDate(FecVenci)

        ImpVenci = DBLet(Rs!TotalFac, "N")
        CodmacBPr = ""
        CodmacBPr = DevuelveDesdeBD("codmacta", "sbanco", "codbanpr", CStr(BanPr), "N")
        
        '13/02/2007
        If vsocio.TipoFactu = 0 Then ' facturacion por tarjeta
            If vTabla = "schfac" Then
                Sql3 = "select numtarje from slhfac where " & cadWHERE2
            Else
                Sql3 = "select numtarje from slhfacr where " & cadWHERE2
            End If
            Set Rs3 = New ADODB.Recordset
            Rs3.Open Sql3, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs3.EOF Then
                '[Monica]22/11/2013: Iban
                Sql4 = "select codbanco, codsucur, digcontr, cuentaba, iban from starje where codsocio = " & vsocio.Codigo & " and numtarje = " & DBSet(Rs3.Fields(0).Value, "N")
                Set Rs4 = New ADODB.Recordset
                Rs4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not Rs4.EOF Then
                    If vEmpresa.HayNorma19_34Nueva Then
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(Rs4!codbanco, "N") & ", " & DBSet(Rs4!codsucur, "N") & ", " & DBSet(Rs4!digcontr, "T") & ", " & DBSet(Rs4!cuentaba, "T") & ", " & DBSet(Rs4!Iban, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(Rs4!codbanco, "N") & ", " & DBSet(Rs4!codsucur, "N") & ", " & DBSet(Rs4!digcontr, "T") & ", " & DBSet(Rs4!cuentaba, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
                    End If
                Else
                    If vEmpresa.HayNorma19_34Nueva Then
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.Iban, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
                    End If
                End If
            Else
                If vEmpresa.HayNorma19_34Nueva Then
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.Iban, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
                Else
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
                End If
            End If

        Else    ' facturacion por cliente
            If vEmpresa.HayNorma19_34Nueva Then
                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.Iban, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
            Else
                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.Digcontrol, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & textcsb33 & "," & DBSet(textcsb41, "T") & ","
            End If

        End If
        CadValues2 = CadValues2 & _
                     DBSet(textcsb42, "T") & "," & DBSet(textcsb43, "T") & "," & DBSet(textcsb51, "T") & "," & DBSet(textcsb52, "T") & "," & DBSet(textcsb53, "T") & "," & DBSet(textcsb61, "T") & "," & DBSet(textcsb62, "T") & "," & DBSet(textcsb63, "T") & "," & DBSet(textcsb71, "T") & "," & _
                     DBSet(textcsb72, "T") & "," & DBSet(textcsb73, "T") & "," & DBSet(textcsb81, "T") & "," & DBSet(textcsb82, "T") & "," & DBSet(textcsb83, "T") & ", 1)"

        If vsocio.CuentaConta <> "" Then
            'antes de grabar en la scobro comprobar que existe en conta.sforpa la
            'forma de pago de la factura. Sino existe insertarla
            'vemos si existe en la conta
            CadValuesAux2 = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", DBLet(Rs!Codforpa), "N")
            'si no existe la forma de pago en conta, insertamos la de ariges
            If CadValuesAux2 = "" Then
                cadValuesAux = "tipforpa"
                CadValuesAux2 = DevuelveDesdeBDNew(cPTours, "sforpa", "nomforpa", "codforpa", DBLet(Rs!Codforpa), "N", cadValuesAux)
                'insertamos e sforpa de la CONTA
                Sql = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                Sql = Sql & " VALUES(" & DBSet(Rs!Codforpa, "N") & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
                ConnConta.Execute Sql
            End If

            'Insertamos en la tabla scobro de la CONTA
            Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur, digcontr, cuentaba,"
                '[Monica]22/11/2013: Iban
            If vEmpresa.HayNorma19_34Nueva Then
                Sql = Sql & "iban, text33csb , text41csb, "
            Else
                Sql = Sql & "text33csb , text41csb, "
            End If
            Sql = Sql & "text42csb, text43csb, text51csb, text52csb, text53csb, text61csb, text62csb, text63csb, text71csb, text72csb, text73csb, text81csb, text82csb, text83csb,agente) "
            Sql = Sql & " VALUES " & CadValues2
            ConnConta.Execute Sql
        End If


    End If

    b = True

EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaDB = b
End Function


Private Function NuevaFechaVto(vFecVenci As Date, DiaPago As Integer) As Date
Dim NewFec As String
Dim Dia As Integer
Dim Mes As Integer
Dim Anyo As Integer
    
    On Error Resume Next
    
    
    NuevaFechaVto = vFecVenci
    
    Dia = Day(vFecVenci)
    Mes = Month(vFecVenci)
    Anyo = Year(vFecVenci)
    
    If DiaPago <= Dia Then
        Mes = Mes + 1
        If Mes > 12 Then
            Mes = 1
            Anyo = Anyo + 1
        End If
        Dia = CInt(DiaPago)
    Else
        Dia = CInt(DiaPago)
    End If
    
    NewFec = Format(Dia, "00") & "/" & Format(Mes, "00") & "/" & Format(Anyo, "0000")
    
    '31
    If Not EsFechaOK(NewFec) Then
        Dia = Dia - 1
        NewFec = Format(Dia, "00") & "/" & Format(Mes, "00") & "/" & Format(Anyo, "0000")
    End If
    '30
    If Not IsDate(NewFec) Then
        Dia = Dia - 1
        NewFec = Format(Dia, "00") & "/" & Format(Mes, "00") & "/" & Format(Anyo, "0000")
    End If
    '29
    If Not IsDate(NewFec) Then
        Dia = Dia - 1
        NewFec = Format(Dia, "00") & "/" & Format(Mes, "00") & "/" & Format(Anyo, "0000")
    End If
    NuevaFechaVto = CDate(NewFec)

End Function
