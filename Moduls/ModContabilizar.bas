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
Private IvaImp As Currency
Private TotalFac As Currency
Private CCoste As String
Private conCtaAlt As Boolean 'el cliente utiliza cuentas alternativas

Private AnyoFacPr As Integer 'año factura proveedor, es el ano de fecha_recepcion

Dim vvIban As String

Private vTipoIva(2) As Currency
Private vPorcIva(2) As Currency
Private vPorcRec(2) As Currency
Private vBaseIva(2) As Currency
Private vImpIva(2) As Currency
Private vImpRec(2) As Currency


Public Function CrearTMPFacturas(cadTabla As String, cadWhere As String) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    If cadTabla = "scafpc" Then
        SQL = "CREATE TEMPORARY TABLE tmpfactu ( "
        SQL = SQL & "codprove int(6) NOT NULL default '0',"
        SQL = SQL & "numfactu varchar(10) NOT NULL default '',"
        SQL = SQL & "fecfactu date NOT NULL default '0000-00-00') "
        Conn.Execute SQL
         
         
        SQL = "SELECT codprove, numfactu, fecfactu"
        SQL = SQL & " FROM " & cadTabla
        SQL = SQL & " WHERE " & cadWhere
        SQL = " INSERT INTO tmpfactu " & SQL
        Conn.Execute SQL
    
        CrearTMPFacturas = True
    
    
    Else
    
        SQL = "CREATE TEMPORARY TABLE tmpfactu ( "
        SQL = SQL & "letraser char(3) NOT NULL default '',"
        SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
        SQL = SQL & "fecfactu date NOT NULL default '0000-00-00') "
        Conn.Execute SQL
         
         
        SQL = "SELECT letraser, numfactu, fecfactu"
        SQL = SQL & " FROM " & cadTabla
        SQL = SQL & " WHERE " & cadWhere
        SQL = " INSERT INTO tmpfactu " & SQL
        Conn.Execute SQL
    
        CrearTMPFacturas = True
        
    End If
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpfactu;"
        Conn.Execute SQL
    End If
End Function


Public Sub BorrarTMPFacturas()
On Error Resume Next

    Conn.Execute " DROP TABLE IF EXISTS tmpfactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function CrearTMPErrFact(cadTabla As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    SQL = "CREATE TEMPORARY TABLE tmperrfac ( "
    If cadTabla = "schfac" Or cadTabla = "schfacr" Then
        SQL = SQL & "codtipom char(1) NOT NULL default '',"
        SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    End If
    SQL = SQL & "fecfactu date NOT NULL default '0000-00-00', "
    SQL = SQL & "error varchar(100) NULL )"
    Conn.Execute SQL
     
    CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmperrfac;"
        Conn.Execute SQL
    End If
End Function


Public Function CrearTMPErrComprob() As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPErrComprob = False
    
    SQL = "CREATE TEMPORARY TABLE tmperrcomprob ( "
    SQL = SQL & "error varchar(100) NULL )"
    Conn.Execute SQL
     
    CrearTMPErrComprob = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrComprob = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmperrcomprob;"
        Conn.Execute SQL
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


Public Function ComprobarLetraSerie(cadTabla As String) As Boolean
'Para Facturas VENTA a clientes
'Comprueba que la letra del serie del tipo de movimiento es  correcta
Dim SQL As String
Dim RS As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetra

    ComprobarLetraSerie = False
    
    'Comprobar que existe la letra de serie en contabilidad
    If cadTabla = "schfac" Then
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
        SQL = "Select distinct tiporegi from contadores"
        Set RSconta = New ADODB.Recordset
        RSconta.Open SQL, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
        If RSconta.EOF Then
            RSconta.Close
            Set RSconta = Nothing
            Exit Function
        End If
            
    
        'obtenemos los distintos tipos de movimiento que vamos a contabilizar
        'de las facturas seleccionadas
        SQL = "select distinct letraser from tmpfactu "

        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        B = True
        While Not RS.EOF 'And b
            'comprobar que todas las letras serie existen en Arigasol
            SQL = "letraser"
            devuelve = DevuelveDesdeBD("letraser", "stipom", "letraser", DBLet(RS!letraser), "T", SQL)
            If devuelve = "" Then
                B = False
                Cad = RS!letraser & " en BD de Arigasol."
                InsertarError "No existe la letra de serie " & Cad
            Else
                'comprobar que todas las letras serie existen en la contabilidad
                devuelve = "tiporegi= '" & devuelve & "'"
                RSconta.MoveFirst
                RSconta.Find (devuelve), , adSearchForward
                If RSconta.EOF Then
                    'no encontrado
                    B = False
                    Cad = SQL & " en BD de Contabilidad."
                    InsertarError "No existe la letra de serie " & Cad
                End If
            End If
            If B Then Cad = Cad & DBSet(RS!letraser, "T") & ","
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        RSconta.Close
        Set RSconta = Nothing
        
        If Not B Then 'Hay algun movimiento que no existe
            devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
            devuelve = devuelve & "Consulte con el administrador."
'            MsgBox devuelve, vbExclamation
            Exit Function
        End If
        
        'Todos los Tipo de movimiento existen
        If Cad <> "" Then
            Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitamos ult. coma
        
            'miramos si hay algun movimiento de factura que la letra serie sea nulo
            SQL = "select count(*) from stipom "
            SQL = SQL & "where letraser IN (" & Cad & ") and (isnull(letraser) or letraser='')"
            If RegistrosAListar(SQL) > 0 Then
                SQL = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                SQL = SQL & "Comprobar en la tabla de tipos de movimiento: " & Cad
                InsertarError SQL
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


Public Function ComprobarNumFacturas(cadTabla As String, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim SQL As String
Dim RS As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas = False
    
    If vParamAplic.ContabilidadNueva Then
        SQL = "SELECT numserie,numfactu,anofactu FROM factcli "
        SQL = SQL & " WHERE " & cadWConta
    Else
        SQL = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
        SQL = SQL & " WHERE " & cadWConta
    End If
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        SQL = "SELECT DISTINCT tmpfactu.letraser,tmpfactu.numfactu,tmpfactu.fecfactu "
        SQL = SQL & " FROM tmpfactu "
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        B = True
        While Not RS.EOF 'And b
' quitado el 12022007
'            SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND codfaccl=" & DBSet(RS!numfactu, "N") & " AND anofaccl=" & Year(RS!fecfactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
            SQL = ""
            If vParamAplic.ContabilidadNueva Then
                SQL = DevuelveDesdeBDNew(cConta, "factcli", "numfactu", "numfactu", RS!numfactu, "N", , "numserie", RS!letraser, "T", "anofactu", Year(RS!fecfactu), "N")
            Else
                SQL = DevuelveDesdeBDNew(cConta, "cabfact", "codfaccl", "codfaccl", RS!numfactu, "N", , "numserie", RS!letraser, "T", "anofaccl", Year(RS!fecfactu), "N")
            End If
            If SQL <> "" Then
                B = False
                SQL = "          Nº Fac.: " & Format(RS!numfactu, "0000000") & vbCrLf
                SQL = SQL & "          Fecha: " & RS!fecfactu
                
                SQL = "Ya existe la factura: " & vbCrLf & SQL
                InsertarError SQL
            
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        
        If Not B Then
            SQL = "Ya existe la factura: " & vbCrLf & SQL
            SQL = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
            
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



Public Function ComprobarCtaContable(cadTabla As String, Opcion As Byte, Optional cadWhere As String) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim RS As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim cadG As String
Dim enc As String
    
    On Error GoTo ECompCta

    ComprobarCtaContable = False
    
    SQL = "SELECT codmacta FROM cuentas "
    SQL = SQL & " WHERE apudirec='S'"
    If cadG <> "" Then SQL = SQL & cadG
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        If Opcion = 1 Then
            Select Case cadTabla
                Case "schfac"
                    'Seleccionamos los distintos clientes,cuentas que vamos a facturar
                    SQL = "SELECT DISTINCT schfac.codsocio, ssocio.codmacta "
                    SQL = SQL & " FROM (schfac INNER JOIN ssocio ON schfac.codsocio=ssocio.codsocio) "
                    SQL = SQL & " INNER JOIN tmpfactu ON schfac.letraser=tmpfactu.letraser AND schfac.numfactu=tmpfactu.numfactu AND schfac.fecfactu=tmpfactu.fecfactu "
                Case "ssocio"
                    SQL = "SELECT DISTINCT scaalb.codsocio, ssocio.codmacta "
                    SQL = SQL & " FROM scaalb, ssocio, sforpa  "
                    SQL = SQL & " where " & cadWhere & " and scaalb.codsocio=ssocio.codsocio and scaalb.codforpa = sforpa.codforpa "
                Case "schfacr"
                    SQL = "SELECT DISTINCT schfacr.codsocio, ssocio.codmacta "
                    SQL = SQL & " FROM (schfacr INNER JOIN ssocio ON schfacr.codsocio=ssocio.codsocio) "
                    SQL = SQL & " INNER JOIN tmpfactu ON schfacr.letraser=tmpfactu.letraser AND schfacr.numfactu=tmpfactu.numfactu AND schfacr.fecfactu=tmpfactu.fecfactu "
            End Select
        ElseIf Opcion = 2 Then
                SQL = "SELECT distinct sartic.codartic "
                SQL = SQL & ", sartic.codmacta, sartic.codmaccl"
                SQL = SQL & " from ((slhfac "
                SQL = SQL & " INNER JOIN tmpfactu ON slhfac.letraser=tmpfactu.letraser AND slhfac.numfactu=tmpfactu.numfactu AND slhfac.fecfactu=tmpfactu.fecfactu) "
                SQL = SQL & "INNER JOIN sartic ON slhfac.codartic=sartic.codartic) "
                SQL = SQL & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
        ElseIf Opcion = 3 Then
                'si hay analitica comprobar que todas las cuentas
                'empiezan por el digito que hay en conta.parametros.grupovta
                cadG = DevuelveDesdeBDNew(cConta, "parametros", "grupovta", "", "", "")
        
                SQL = "SELECT distinct sartic.codartic "
                SQL = SQL & ", sartic.codmacta, sartic.codmaccl"
                SQL = SQL & " from ((slhfac "
                SQL = SQL & " INNER JOIN tmpfactu ON slhfac.letraser=tmpfactu.letraser AND slhfac.numfactu=tmpfactu.numfactu AND slhfac.fecfactu=tmpfactu.fecfactu) "
                SQL = SQL & "INNER JOIN sartic ON slhfac.codartic=sartic.codartic) "
                SQL = SQL & " where sartic.codmacta "
                If cadG <> "" Then
                     SQL = SQL & " AND not ((sartic.codmacta like '" & cadG & "%') and (sartic.codmaccl like '" & cadG & "%'))"
                End If
        ElseIf Opcion = 4 Then
            SQL = "select codmacta from sbanco where codbanpr = " & cadTabla
        ElseIf Opcion = 5 Then
            SQL = "select codmacta from sforpa where cuadresn = 1 and not codmacta is null and mid(codmacta,1,1) <> ' '"
        ElseIf Opcion = 6 Then
            SQL = "select ctaposit from sparam"
        ElseIf Opcion = 7 Then
            SQL = "select ctanegtat from sparam"
        End If
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        B = True
        While Not RS.EOF 'And b
            If Opcion = 3 Then
                B = False
                SQL = DBLet(RS!codmacta, "T") & " o " & DBLet(RS!CodmacCl, "T")
                SQL = "La cuenta " & SQL & " del articulo " & RS!codArtic & " no es del grupo correcto."
                InsertarError SQL
            Else
                If Opcion = 6 Or Opcion = 7 Then
                    SQL = "codmacta= " & DBSet(RS.Fields(0).Value, "T") '& " and apudirec='S' "
                Else
                    SQL = "codmacta= " & DBSet(RS!codmacta, "T") '& " and apudirec='S' "
                End If
' comentado 12022007
'                RSconta.MoveFirst
'                RSconta.Find (SQL), , adSearchForward
'                If RSconta.EOF Then
                 enc = ""
                 If Opcion = 6 Or Opcion = 7 Then
                    enc = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", DBLet(RS.Fields(0).Value, "T"), "T")
                 Else
                    enc = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", DBLet(RS!codmacta, "T"), "T")
                 End If
                 
                 If enc = "" Then
                    B = False 'no encontrado
                    If Opcion = 1 Then
                        If cadTabla = "schfac" Or cadTabla = "ssocio" Or cadTabla = "schfacr" Then
                            SQL = DBLet(RS!codmacta, "T") & " del Cliente " & Format(RS!codsocio, "000000")
                            SQL = "No existe la cta contable " & SQL
                            InsertarError SQL
                        End If
                    End If
                    If Opcion = 2 Then
                        SQL = DBLet(RS!codmacta, "T") & " del Artículo " & Format(RS!codArtic, "000000")
                        SQL = "No existe la cta contable " & SQL
                        InsertarError SQL
                    End If
                    If Opcion = 4 Then
                        SQL = DBLet(RS!codmacta, "T") & " del Banco " & Format(CCur(cadTabla), "000")
                        SQL = "No existe la cta contable " & SQL
                        InsertarError SQL
                    End If
                    If Opcion = 6 Or Opcion = 7 Then
                        SQL = "No existe la cta contable " & SQL
                        InsertarError SQL
                    End If
                End If
                
                ' en caso de que estemos comprobando las cuentas contables del articulo
                ' comprobamos tb la cuenta contable socio del articulo
                '---------------------------------------------------------------------
                If Opcion = 2 Then
                    If Not IsNull(RS!CodmacCl) Then
                        SQL = "codmacta= " & DBSet(RS!CodmacCl, "T") '& " and apudirec='S' "
                        enc = ""
                        enc = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", DBLet(RS!CodmacCl, "T"), "T")
                        If enc = "" Then
' comentado el 12022007
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
                            B = False 'no encontrado
                            SQL = DBLet(RS!CodmacCl, "T") & " del artículo " & Format(RS!codArtic, "000000")
                            SQL = "No existe la cta contable cliente " & SQL
                            InsertarError SQL
                        End If
                    Else
                        B = False 'no encontrado
                        SQL = DBLet(RS!CodmacCl, "T") & " del artículo " & Format(RS!codArtic, "000000")
                        SQL = "No existe la cta contable cliente " & SQL
                        InsertarError SQL
                    End If
                End If
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        
        If Not B Then
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





Public Function ComprobarTiposIVA(cadTabla As String) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim RS As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    SQL = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For i = 1 To 3
            If cadTabla = "schfac" Then
                SQL = "SELECT DISTINCT schfac.tipoiva" & i
                SQL = SQL & " FROM schfac "
                SQL = SQL & " INNER JOIN tmpfactu ON schfac.letraser=tmpfactu.letraser AND schfac.numfactu=tmpfactu.numfactu AND schfac.fecfactu=tmpfactu.fecfactu "
                SQL = SQL & " WHERE not isnull(tipoiva" & i & ")"
            Else
                If cadTabla = "scafpc" Then
                    SQL = "SELECT DISTINCT scafpc.tipoiva" & i
                    SQL = SQL & " FROM scafpc "
                    SQL = SQL & " INNER JOIN tmpfactu ON scafpc.codprove=tmpfactu.codprove AND scafpc.numfactu=tmpfactu.numfactu AND scafpc.fecfactu=tmpfactu.fecfactu "
                    SQL = SQL & " WHERE not isnull(tipoiva" & i & ")"
                End If
            End If

            Set RS = New ADODB.Recordset
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            B = True
            While Not RS.EOF 'And b
                If RS.Fields(0) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                    SQL = "codigiva= " & DBSet(RS.Fields(0), "N")
                    RSconta.MoveFirst
                    RSconta.Find (SQL), , adSearchForward
                    If RSconta.EOF Then
                        B = False 'no encontrado
                        SQL = "No existe el " & SQL
                        SQL = "Tipo de IVA: " & RS.Fields(0)
                        InsertarError SQL
                    End If
                End If
                RS.MoveNext
            Wend
            RS.Close
            Set RS = Nothing
        
            If Not B Then
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


Public Function PasarFactura(cadWhere As String, FecVenci As String, Banpr As String, CodCCost As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim SQL As String
Dim vsocio As CSocio
Dim codsoc As Long

Dim LetraInt As String  ' letra de serie de las facturas internas

Dim RS As ADODB.Recordset

Dim RSx As ADODB.Recordset
Dim sql2 As String
Dim codfor As Integer
Dim TipForpa As String
Dim Mc As CContadorContab
Dim Obs As String





    On Error GoTo EContab

    ConnConta.BeginTrans
    Conn.BeginTrans
    
    'seleccionamos el socio de la factura
    '[Monica]04/03/2011: Facturas internas añado en el select la letra de serie
    SQL = "select codsocio, letraser, fecfactu from schfac where " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenStatic, adLockPessimistic, adCmdText
    
    codsoc = 0
    
    If Not RS.EOF Then
        codsoc = RS.Fields(0).Value
        LetraInt = RS.Fields(1).Value
    End If
    
    Set vsocio = New CSocio
    If vsocio.LeerDatos(CStr(codsoc)) Then
'[Monica]25/07/2013: serie internas
'        '[Monica]04/03/2011: Facturas internas añado en el select la letra de serie
'        If LetraInt = vParamAplic.LetraInt Then
        If EsInterna(LetraInt) Then
            Set Mc = New CContadorContab
            
            If Mc.ConseguirContador("0", (RS!fecfactu <= CDate(FFin)), True) = 0 Then
            
                Obs = "Contabilización Factura Interna de Fecha " & Format(RS!fecfactu, "dd/mm/yyyy")
            
                'Insertar en la conta Cabecera Asiento
                B = InsertarCabAsientoDia(vEmpresa.NumDiarioInt, Mc.Contador, RS!fecfactu, Obs, cadMen)
                cadMen = "Insertando Cab. Asiento: " & cadMen
            Else
                B = False
            End If
        Else
            'Insertar en la conta Cabecera Factura
            B = InsertarCabFact(cadWhere, cadMen)
            cadMen = "Insertando Cab. Factura: " & cadMen
        End If
            
        ' insertar en tesoreria
        If B Then
            sql2 = "select codforpa from schfac where " & cadWhere
            Set RSx = New ADODB.Recordset
            RSx.Open sql2, Conn, adOpenStatic, adLockPessimistic, adCmdText
            
            If Not RSx.EOF Then codfor = RSx.Fields(0).Value
            TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", DBSet(RSx.Fields(0).Value, "N"), "N")
            
'[Monica]16/12/2010: solo se inserta en tesoreria si no hacen la contabilizacion de cierre de turno
            '[Monica]04/01/2013: Efectivos
            '[Monica]11/01/2013: En Ribarroja se inserta en Tesoreria
            If (TipForpa <> "0" And TipForpa <> "6") Or vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 5 Then
            
                B = InsertarEnTesoreria(cadWhere, FecVenci, Banpr, cadMen, vsocio, "schfac")
                cadMen = "Insertando en Tesoreria: " & cadMen
            End If
            
            Set RSx = Nothing
            
        End If
    
        If B Then
'[Monica]25/07/2013: serie internas
'            If LetraInt = vParamAplic.LetraInt Then
            If EsInterna(LetraInt) Then
                B = InsertarLinAsientoFactInt("schfac", cadWhere, cadMen, vsocio, Mc.Contador)
                cadMen = "Insertando Lin. Factura Interna: " & cadMen
            
                Set Mc = Nothing
            Else
        '        CCoste = CodCCost
                'Insertar lineas de Factura en la Conta
                '21032007
                '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
                If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 5 Then ' si Alzira o Pobla del Duc o Ribarroja
                    B = InsertarLinFact("schfac", cadWhere, cadMen, vsocio)
                Else
                    B = InsertarLinFactReg("schfac", cadWhere, cadMen, vsocio)
                End If
                cadMen = "Insertando Lin. Factura: " & cadMen
                
            End If
            If B Then
                'Poner intconta=1 en arigasol.scafac
                B = ActualizarCabFact("schfac", cadWhere, cadMen)
                cadMen = "Actualizando Factura: " & cadMen
            End If
        End If
        
        If Not B Then
            SQL = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
            SQL = SQL & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
            SQL = SQL & " WHERE " & Replace(cadWhere, "schfac", "tmpfactu")
            Conn.Execute SQL
        End If
    End If
    
    Set vsocio = Nothing
    
EContab:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        Conn.CommitTrans
        PasarFactura = True
    Else
        ConnConta.RollbackTrans
        Conn.RollbackTrans
        PasarFactura = False
    End If
End Function

Public Function PasarFactura2(cadWhere As String, ByRef vsocio As CSocio, vTabla As String) As Boolean   ' , CodCCost As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim SQL As String

    On Error GoTo EContab
    
    'Insertar en la conta Cabecera Factura
    B = InsertarCabFact(cadWhere, cadMen, vTabla)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If B Then
'        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        B = InsertarLinFact("schfac", cadWhere, cadMen, vsocio)
        cadMen = "Insertando Lin. Factura: " & cadMen

        If B Then
            'Poner intconta=1 en arigasol.scafac
            B = ActualizarCabFact("schfac", cadWhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    If Not B Then
        SQL = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
        SQL = SQL & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
        SQL = SQL & " WHERE " & Replace(cadWhere, "scafac", "tmpfactu")
        Conn.Execute SQL
    End If
    
EContab:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If B Then
        PasarFactura2 = True
    Else
        PasarFactura2 = False
    End If
End Function

Public Function PasarFactura3(cadWhere As String, FecVenci As String, Banpr As String, CodCCost As String, ByRef cadMen As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' arigasol.schfac --> conta.cabfact
' arigasol.slhfac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim B As Boolean
'Dim cadMen As String
Dim SQL As String
Dim vsocio As CSocio
Dim codsoc As Long
Dim RS As ADODB.Recordset

Dim RSx As ADODB.Recordset
Dim sql2 As String
Dim codfor As Integer
Dim TipForpa As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    Conn.BeginTrans
    
    'seleccionamos el socio de la factura
    SQL = "select codsocio from schfacr where " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenStatic, adLockPessimistic, adCmdText
    
    codsoc = 0
    
    If Not RS.EOF Then codsoc = RS.Fields(0).Value
    
    
    Set vsocio = New CSocio
    If vsocio.LeerDatos(CStr(codsoc)) Then
    
        
        ' insertar en tesoreria
        sql2 = "select codforpa from schfacr where " & cadWhere
        Set RSx = New ADODB.Recordset
        RSx.Open sql2, Conn, adOpenStatic, adLockPessimistic, adCmdText
        
        If Not RSx.EOF Then codfor = RSx.Fields(0).Value
        TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", DBSet(RSx.Fields(0).Value, "N"), "N")
        '[Monica]04/01/2013: efectivos
        If TipForpa <> "0" And TipForpa <> "6" Then
            B = InsertarEnTesoreriaAjenas(cadWhere, FecVenci, Banpr, cadMen, vsocio, "schfacr")
            cadMen = "Insertando en Tesoreria: " & cadMen
        End If
        
        Set RSx = Nothing
        
        If B Then
            'Poner intconta=1 en arigasol.scafac
            B = ActualizarCabFact("schfacr", cadWhere, cadMen)
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
        B = False
        MuestraError Err.Number, "Contabilizando Factura Ajena en Tesorería", Err.Description
    End If
    If B Then
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
Dim B As Boolean
Dim cadMen As String
Dim SQL As String
Dim sql2 As String
Dim RSx As ADODB.Recordset
Dim cadWhere As String
Dim codfor As String
Dim TipForpa As String
Dim ctabancl As String
Dim FecVenci As String
Dim Banpr As String

    On Error GoTo EContab
    
    'Insertar en la conta Cabecera Factura
    cadWhere = "letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(NueNumfact, "N") & " and fecfactu = " & DBSet(NueFecFactu, "F")
    
    B = InsertarCabFact(cadWhere, cadMen)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If B Then
'        CCoste = CodCCost
        'Insertar lineas de Factura en la Conta
        B = InsertarLinFact("schfac", cadWhere, cadMen, vsocio)
        cadMen = "Insertando Lin. Factura: " & cadMen

        sql2 = "select codforpa from schfac where " & cadWhere
        Set RSx = New ADODB.Recordset
        RSx.Open sql2, Conn, adOpenStatic, adLockPessimistic, adCmdText
        
        If Not RSx.EOF Then codfor = RSx.Fields(0).Value
        TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", DBSet(RSx.Fields(0).Value, "N"), "N")
        
        ctabancl = "ctabanc1"
        '[Monica]04/01/2013 : efectivos
        If TipForpa <> "0" And TipForpa <> "6" Then
            FecVenci = ""
            If vParamAplic.ContabilidadNueva Then
                FecVenci = DevuelveDesdeBDNew(cConta, "cobros", "fecvenci", "numserie", letraser, "T", ctabancl, "numfactu", numfactu, "N", "fecfactu", fecfactu, "F")
            Else
                FecVenci = DevuelveDesdeBDNew(cConta, "scobro", "fecvenci", "numserie", letraser, "T", ctabancl, "codfaccl", numfactu, "N", "fecfaccl", fecfactu, "F")
            End If
            If FecVenci <> "" Then
                Banpr = DevuelveDesdeBDNew(cPTours, "sbanco", "codbanpr", "codmacta", ctabancl, "T")
            
                B = InsertarEnTesoreria(cadWhere, FecVenci, Banpr, cadMen, vsocio, "schfac")
                cadMen = "Insertando en Tesoreria: " & cadMen
            End If
        End If
        
        Set RSx = Nothing
    End If
    
EContab:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, "Contabilizando Factura", Err.Description & cadMen
        PasarFactura4 = False
    Else
        PasarFactura4 = True
    End If
End Function



Private Function InsertarCabFact(cadWhere As String, caderr As String, Optional vTabla As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Cad As String
Dim CadenaInsertFaclin2 As String
Dim sql2 As String

    On Error GoTo EInsertar
    
    SQL = " SELECT letraser,numfactu,fecfactu, ssocio.codmacta, year(fecfactu) as anofaccl,"
    SQL = SQL & "baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    SQL = SQL & "totalfac,tipoiva1,tipoiva2,tipoiva3, "
    
    If vTabla <> "" Then
        SQL = SQL & vTabla & ".codforpa,"
    Else
        SQL = SQL & "schfac.codforpa,"
    End If
        
    SQL = SQL & "ssocio.nomsocio, ssocio.domsocio, ssocio.codposta, ssocio.pobsocio, ssocio.prosocio, ssocio.nifsocio"
    
    '[Monica]24/07/2013:
    If vTabla <> "" Then
        SQL = SQL & " FROM " & vTabla
        SQL = SQL & "INNER JOIN " & "ssocio ON " & vTabla & ".codsocio=ssocio.codsocio"
    Else
        SQL = SQL & " FROM " & "schfac "
        SQL = SQL & "INNER JOIN " & "ssocio ON schfac.codsocio=ssocio.codsocio"
    End If
    SQL = SQL & " WHERE " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not RS.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        BaseImp = DBLet(RS!baseimp1, "N") + DBLet(RS!baseimp2, "N") + DBLet(RS!baseimp3, "N")
        IvaImp = DBLet(RS!impoiva1, "N") + DBLet(RS!impoiva2, "N") + DBLet(RS!impoiva3, "N")
        
        If vParamAplic.ContabilidadNueva Then
            SQL = ""
            SQL = DBSet(RS!letraser, "T") & "," & DBSet(RS!numfactu, "N") & "," & DBSet(RS!fecfactu, "F") & "," & DBSet(RS!codmacta, "T") & "," & Year(RS!fecfactu) & ",0,'FACTURACION',0,"
            SQL = SQL & DBSet(RS!codfopa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
            SQL = SQL & ValorNulo & "," & DBSet(RS!TotalFac, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0," & DBSet(RS!fecfactu, "F") & ","
            SQL = SQL & DBSet(RS!NomSocio, "T") & "," & DBSet(RS!domsocio, "T") & "," & DBSet(RS!codPosta, "T") & "," & DBSet(RS!pobsocio, "T") & ","
            SQL = SQL & DBSet(RS!prosocio, "T") & "," & DBSet(RS!nifsocio, "T") & ",'ES',1)"
            
            Cad = Cad & "(" & SQL & ")"
        
        Else
        
            SQL = ""
            SQL = DBSet(RS!letraser, "T") & "," & DBSet(RS!numfactu, "N") & "," & DBSet(RS!fecfactu, "F") & "," & DBSet(RS!codmacta, "T") & "," & Year(RS!fecfactu) & ",'FACTURACION',"
            SQL = SQL & DBSet(RS!baseimp1, "N") & "," & DBSet(RS!baseimp2, "N") & "," & DBSet(RS!baseimp3, "N") & "," & DBSet(RS!porciva1, "N") & "," & DBSet(RS!porciva2, "N") & "," & DBSet(RS!porciva3, "N") & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(RS!impoiva1, "N", "N") & "," & DBSet(RS!impoiva2, "N") & "," & DBSet(RS!impoiva3, "N") & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(RS!TotalFac, "N") & "," & DBSet(RS!TipoIVA1, "N") & "," & DBSet(RS!TipoIVA2, "N", "S") & "," & DBSet(RS!TipoIVA3, "N", "S") & ",0,"
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(RS!fecfactu, "F")
            Cad = Cad & "(" & SQL & ")"
        End If
    End If
    RS.Close
    Set RS = Nothing
    
    If vParamAplic.ContabilidadNueva Then
        SQL = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
        SQL = SQL & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
        SQL = SQL & "codpais,codagente)"
        SQL = SQL & " VALUES " & Cad
        ConnConta.Execute SQL
    '$$$
'***
        CadenaInsertFaclin2 = ""
            
        
        'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
        'IVA 1, siempre existe
        sql2 = "'" & RS!letraser & "'," & RS!numfactu & "," & DBSet(RS!fecfactu, "F") & "," & Year(RS!fecfactu) & ","
        sql2 = sql2 & "1," & DBSet(RS!baseimp1, "N") & "," & RS!TipoIVA1 & "," & DBSet(RS!porciva1, "N") & ","
        sql2 = sql2 & ValorNulo & "," & DBSet(RS!impoiva1, "N") & "," & ValorNulo
        CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & sql2 & ")"
        
        'para las lineas
        vTipoIva(0) = RS!TipoIVA1
        vPorcIva(0) = RS!porciva1
        vPorcRec(0) = 0
        vImpIva(0) = RS!impoiva1
        vImpRec(0) = 0
        vBaseIva(0) = RS!baseimp1
        
        vTipoIva(1) = 0: vTipoIva(2) = 0
        
        If Not IsNull(RS!porcva2) Then
            sql2 = "'" & RS!letraser & "'," & RS!numfactu & "," & DBSet(RS!fecfactu, "F") & "," & Year(RS!fecfactu) & ","
            sql2 = sql2 & "2," & DBSet(RS!baseimp2, "N") & "," & RS!TipoIVA2 & "," & DBSet(RS!porciva2, "N") & ","
            sql2 = sql2 & ValorNulo & "," & DBSet(RS!impoiva2, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & sql2 & ")"
            vTipoIva(1) = RS!TipoIVA2
            vPorcIva(1) = RS!porciva2
            vPorcRec(1) = 0
            vImpIva(1) = RS!impoiva2
            vImpRec(1) = 0
            vBaseIva(1) = RS!baseimp2
        End If
        If Not IsNull(RS!porciva3) Then
            sql2 = "'" & RS!letraser & "'," & RS!numfactu & "," & DBSet(RS!fecfactu, "F") & "," & Year(RS!fecfactu) & ","
            sql2 = sql2 & "3," & DBSet(RS!baseimp3, "N") & "," & RS!codigiv3 & "," & DBSet(RS!porciva3, "N") & ","
            sql2 = sql2 & DBSet(RS!porciva3re, "N") & "," & DBSet(RS!imporiv3, "N") & "," & DBSet(RS!imporiv3re, "N")
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & sql2 & ")"
            vTipoIva(2) = RS!TipoIVA3
            vPorcIva(2) = RS!porciva3
            vPorcRec(2) = 0
            vImpIva(2) = RS!impoiva3
            vImpRec(2) = 0
            vBaseIva(2) = RS!baseimp3
        End If


        SQL = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
        SQL = SQL & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
        ConnConta.Execute SQL

    
'***
    
    Else
        'Insertar en la contabilidad
        SQL = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
        SQL = SQL & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
        SQL = SQL & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
        SQL = SQL & " VALUES " & Cad
        ConnConta.Execute SQL
    End If
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFact = False
        caderr = Err.Description
    Else
        InsertarCabFact = True
    End If
End Function


Private Function InsertarLinAsientoFactInt(cadTabla As String, cadWhere As String, caderr As String, ByRef vsocio As CSocio, Optional Contador As Long) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim numdocum As String
Dim ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim i As Long
Dim B As Boolean
Dim Cad As String
Dim cadMen As String
Dim FeFact As Date

    On Error GoTo eInsertarLinAsientoFactInt

    InsertarLinAsientoFactInt = False
    
    '[Monica]25/09/2014: cambiado tipoconta = 1 indica sobre cuenta contable del socio, 0 = cuenta contable del cliente
    If vsocio.TipoConta = 1 Then
        SQL = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmacta, " ' sartic.codmaccl, "
        SQL = SQL & " sum(implinea) as importe FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        SQL = SQL & " WHERE " & Replace(cadWhere, "schfac", "slhfac")
        SQL = SQL & " GROUP BY 1,2,3,5"
    Else
        SQL = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmaccl codmacta, "
        SQL = SQL & " sum(implinea) as importe FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        SQL = SQL & " WHERE " & Replace(cadWhere, "schfac", "slhfac")
        SQL = SQL & " GROUP BY 1,2,3,5"
    End If

    
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
            
    i = 0
    ImporteD = 0
    ImporteH = 0
    
    numdocum = Format(RS!numfactu, "0000000")
    '[Monica]25/07/2013: letra de serie
'    ampliacion = vParamAplic.LetraInt & "-" & Format(Rs!numfactu, "0000000")
    ampliacion = Trim(RS!letraser) & "-" & Format(RS!numfactu, "0000000")
    ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & ampliacion
    ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", vEmpresa.ConceptoInt, "N")) & " " & ampliacion
    
    If Not RS.EOF Then RS.MoveFirst
    
    B = True
    
    While Not RS.EOF And B
        i = i + 1
        
        FeFact = RS!fecfactu
        
        Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(RS!fecfactu, "F") & "," & DBSet(Contador, "N") & ","
        Cad = Cad & DBSet(i, "N") & "," & DBSet(RS!codmacta, "T") & "," & DBSet(numdocum, "T") & ","
        
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If RS.Fields(5).Value < 0 Then
            ' importe al debe en positivo
            Cad = Cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(RS.Fields(5).Value * (-1), "N") & ","
            Cad = Cad & ValorNulo & "," & ValorNulo & "," & DBSet(vsocio.CuentaConta, "T") & "," & ValorNulo & ",0"
        
            ImporteD = ImporteD + (CCur(RS.Fields(5).Value) * (-1))
        Else
            ' importe al haber en positivo, cambiamos el signo
            Cad = Cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
            Cad = Cad & DBSet((RS.Fields(5).Value), "N") & "," & ValorNulo & "," & DBSet(vsocio.CuentaConta, "T") & "," & ValorNulo & ",0"
        
            ImporteH = ImporteH + CCur(RS.Fields(5).Value)
        End If
        
        Cad = "(" & Cad & ")"
        
        B = InsertarLinAsientoDia(Cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & i

        RS.MoveNext
    Wend
    
    If B And i > 0 Then
        i = i + 1
                
        ' el Total es sobre la cuenta del cliente
        Cad = DBSet(vEmpresa.NumDiarioInt, "N") & "," & DBSet(FeFact, "F") & "," & DBSet(Contador, "N") & ","
        Cad = Cad & DBSet(i, "N") & ","
        Cad = Cad & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(numdocum, "T") & ","
            
        ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
        If ImporteD - ImporteH > 0 Then
            ' importe al debe en positivo
            Cad = Cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliaciond, "T") & "," & ValorNulo & ","
            Cad = Cad & DBSet(ImporteD - ImporteH, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        Else
            ' importe al haber en positivo, cambiamos el signo
            Cad = Cad & DBSet(vEmpresa.ConceptoInt, "N") & "," & DBSet(ampliacionh, "T") & "," & DBSet(((ImporteD - ImporteH) * -1), "N") & ","
            Cad = Cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
        End If
        
        Cad = "(" & Cad & ")"
        
        B = InsertarLinAsientoDia(Cad, cadMen)
        cadMen = "Insertando Lin. Asiento: " & i
        
    End If
        
    Set RS = Nothing
    InsertarLinAsientoFactInt = B
    Exit Function
    
eInsertarLinAsientoFactInt:
    caderr = "Insertar Linea Asiento Factura Interna: " & Err.Description
    caderr = caderr & cadMen
    InsertarLinAsientoFactInt = False
End Function


Private Function InsertarLinFact(cadTabla As String, cadWhere As String, caderr As String, ByRef vsocio As CSocio, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SqlAux As String
Dim sql2 As String

Dim RS As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim iva As String
Dim vIva As Currency


    On Error GoTo EInLinea

    If cadTabla = "schfac" Then
        '[Monica]25/09/2014: cambiado tipoconta = 1 indica sobre cuenta contable del socio, 0 = cuenta contable del cliente
        If vsocio.TipoConta = 1 Then
            SQL = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmacta, " ' sartic.codmaccl, "
            SQL = SQL & " sum(implinea) as importe FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            SQL = SQL & " WHERE " & Replace(cadWhere, "schfac", "slhfac")
            SQL = SQL & " GROUP BY 1,2,3,5"
        Else
            SQL = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmaccl, "
            SQL = SQL & " sum(implinea) as importe FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
            SQL = SQL & " WHERE " & Replace(cadWhere, "schfac", "slhfac")
            SQL = SQL & " GROUP BY 1,2,3,5"
        End If
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    SqlAux = ""
    While Not RS.EOF
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        'de multibase
        'Let v_base = Round(basesfac / (1 + (porc_iva / 100)), 2)
'        Implinea = CCur(CalcularBase(CStr(RS!Importe), CStr(RS!codartic)))
        SqlAux = Cad
        
        ImpLinea = CCur(CalcularBase(CStr(RS.Fields(5).Value), CStr(RS!codArtic)))
        
        ImpLinea = Round2(ImpLinea, 2)
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        sql2 = ""
        
        SQL = "'" & RS!letraser & "'," & RS!numfactu & "," & Year(RS!fecfactu) & "," & i & ","
        
        '[Monica]25/09/2014: cambiado tipoconta = 1 indica sobre cuenta contable del socio, 0 = cuenta contable del cliente
        If vsocio.TipoConta = 1 Then
            SQL = SQL & DBSet(RS!codmacta, "T")
        Else
            SQL = SQL & DBSet(RS!CodmacCl, "T")
        End If
        
        sql2 = SQL & ","
        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
        
        If CCoste = "" Then
            SQL = SQL & ValorNulo
        Else
            SQL = SQL & DBSet(CCoste, "T")
        End If
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        sql2 = sql2 & DBSet(totimp, "N") & ","
        
        If CCoste = "" Then
            sql2 = sql2 & ValorNulo
        Else
            sql2 = sql2 & DBSet(CCoste, "T")
        End If
        If SqlAux <> "" Then 'hay mas de una linea
            Cad = SqlAux & "(" & sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & sql2 & ")" & ","
        End If
        
        
        
'        Aux = Replace(sql, DBSet(Implinea, "N"), DBSet(totimp, "N"))
'        cad = Replace(cad, sql, Aux)
    End If


    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTabla = "schfac" Then
            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        End If
        SQL = SQL & " VALUES " & Cad
        ConnConta.Execute SQL
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact = False
        caderr = Err.Description
    Else
        InsertarLinFact = True
    End If
End Function



Private Function InsertarLinFactReg(cadTabla As String, cadWhere As String, caderr As String, ByRef vsocio As CSocio, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim Sql1 As String
Dim RS As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim iva As String
Dim vIva As Currency
Dim Impuesto As Currency
Dim Impue As Currency
Dim TotalImpuesto As Currency

Dim numfactu As Long
Dim letraser As String
Dim fecfactu As Date

    On Error GoTo EInLinea

    '[Monica]25/09/2014: cambiado tipoconta = 1 indica sobre cuenta contable del socio, 0 = cuenta contable del cliente
    If vsocio.TipoConta = 1 Then
        SQL = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmacta, " ' sartic.codmaccl, "
        SQL = SQL & " sum(implinea) as importe, sum(cantidad) as cantidad FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        SQL = SQL & " WHERE " & Replace(cadWhere, "schfac", "slhfac")
        SQL = SQL & " GROUP BY 1,2,3,5"
    Else
        SQL = " SELECT slhfac.letraser,numfactu,fecfactu,sartic.codartic,sartic.codmaccl, "
        SQL = SQL & " sum(implinea) as importe, sum(cantidad) as cantidad FROM slhfac inner join sartic on slhfac.codartic=sartic.codartic "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        SQL = SQL & " WHERE " & Replace(cadWhere, "schfac", "slhfac")
        SQL = SQL & " GROUP BY 1,2,3,5"
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    
    totimp = 0
    TotalImpuesto = 0
    
    While Not RS.EOF
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        'de multibase
        'Let v_base = Round(basesfac / (1 + (porc_iva / 100)), 2)
'        Implinea = CCur(CalcularBase(CStr(RS!Importe), CStr(RS!codartic)))
        
        numfactu = RS!numfactu
        letraser = RS!letraser
        fecfactu = RS!fecfactu
        
        
        ' se quita el impuesto por linea
        Sql1 = ""
        Sql1 = DevuelveDesdeBD("impuesto", "sartic", "codartic", DBLet(RS!codArtic), "N")
        If Sql1 = "" Then
            Impuesto = 0
        Else
            Impuesto = CCur(Sql1) ' Comprueba si es nulo y lo pone a 0 o ""
        End If
        
        If EsArticuloCombustible(RS!codArtic) Then
            Impue = Round2((RS.Fields(6).Value * Impuesto), 2)
            TotalImpuesto = TotalImpuesto + Impue
        End If
        
        
        ImpLinea = CCur(CalcularBase(CStr(RS.Fields(5).Value), CStr(RS!codArtic))) - Impue
        ImpLinea = Round2(ImpLinea, 2)
        
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL = "'" & RS!letraser & "'," & RS!numfactu & "," & Year(RS!fecfactu) & "," & i & ","
        
        '[Monica]25/09/2014: cambiado tipoconta = 1 indica sobre cuenta contable del socio, 0 = cuenta contable del cliente
        If vsocio.TipoConta = 1 Then
            SQL = SQL & DBSet(RS!codmacta, "T")
        Else
            SQL = SQL & DBSet(RS!CodmacCl, "T")
        End If
        
        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
        
        If CCoste = "" Then
            SQL = SQL & ValorNulo
        Else
            SQL = SQL & DBSet(CCoste, "T")
        End If
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    totimp = totimp + TotalImpuesto
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(totimp, "N"))
        Cad = Replace(Cad, SQL, Aux)
    End If

    ' insertamos la linea de base de impuesto
    '20/12/2012: dependiendo de la fecha de cambio
    If fecfactu < CDate(vParamAplic.FechaCam) Then
        SQL = ""
        SQL = "'" & letraser & "'," & numfactu & "," & Year(fecfactu) & "," & i & ","
        SQL = SQL & DBSet(vParamAplic.CtaImpuesto, "T")
        SQL = SQL & "," & DBSet(TotalImpuesto, "N") & ","
        If CCoste = "" Then
            SQL = SQL & ValorNulo
        Else
            SQL = SQL & DBSet(CCoste, "T")
        End If
        Cad = Cad & "(" & SQL & "),"
    End If
    
    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTabla = "schfac" Then
            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        End If
        SQL = SQL & " VALUES " & Cad
        ConnConta.Execute SQL
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactReg = False
        caderr = Err.Description
    Else
        InsertarLinFactReg = True
    End If
End Function







Private Function ActualizarCabFact(cadTabla As String, cadWhere As String, caderr As String) As Boolean
'Poner la factura como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    
    SQL = "UPDATE " & cadTabla & " SET intconta=1 "
    SQL = SQL & " WHERE " & cadWhere

    Conn.Execute SQL
    
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
Public Function InsertarEnTesoreria(cadWhere As String, ByVal FechaVen As String, Banpr As String, MenError As String, ByRef vsocio As CSocio, vTabla As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim B As Boolean
Dim RS As ADODB.Recordset
Dim RSx As ADODB.Recordset
Dim SQL As String, textcsb33 As String, textcsb41 As String
Dim sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Sql5 As String
Dim Rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
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
Dim FecVenci As Date
Dim rsVenci As ADODB.Recordset
Dim TotalFactura2 As Currency

Dim LetraS As String



    On Error GoTo EInsertarTesoreria

    B = False
    InsertarEnTesoreria = False
    CadValues = ""
    CadValues2 = ""

    SQL = "select * from " & vTabla & " where  " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
    
        textcsb33 = "FACT: " & DBLet(RS!letraser, "T") & "-" & Format(DBLet(RS!numfactu, "N"), "0000000") & " " & Format(DBLet(RS!fecfactu, "F"), "dd/mm/yy")
        textcsb33 = textcsb33 & " de " & DBSet(RS!TotalFac, "N")
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
                cadWHERE2 = Replace(cadWhere, "schfac", "slhfac")
            Case "schfacr"
                cadWHERE2 = Replace(cadWhere, "schfacr", "slhfacr")
            Case "schfac1"
                cadWHERE2 = Replace(cadWhere, "schfac1", "slhfac1")
        End Select

        
'[Monica]08/01/2014: lo cambiamos rellenando lo maximo que podemos
        If vParamAplic.Cooperativa = 5 Then
            Dim cad1 As String
            Dim cad2 As String
            Dim cad22 As String
            
            SQL = "select count(distinct numalbar) from " & vTabla & " where " & cadWhere
            cad1 = ""
            sql2 = "select numalbar, fecalbar, sum(implinea) "
            Select Case vTabla
                Case "schfac"
                    sql2 = sql2 & " from slhfac where " & cadWHERE2
                Case "schfacr"
                    sql2 = sql2 & " from slhfacr where " & cadWHERE2
                Case "schfac1"
                    sql2 = sql2 & " from slhfac1 where " & cadWHERE2
            End Select

            sql2 = sql2 & " group by numalbar, fecalbar order by numalbar, fecalbar "
            
            Set RSx = New ADODB.Recordset
            RSx.Open sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        SQL = "SELECT numerove, diasvto primerve, restoven FROM sforpa WHERE codforpa=" & DBLet(RS!CodForpa, "N")
        Set rsVenci = New ADODB.Recordset
        rsVenci.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not rsVenci.EOF Then
            If rsVenci!numerove > 0 And DBLet(RS!TotalFac) <> 0 Then
        
                '++[Monica]05/08/2011: si no hay fecha de vencimiento ponemos la fecha de factura, si no los calculos se hacen con la
                '                    fecha de vencimiento
                If FechaVen = "" Then
                    FechaVen = DBLet(RS!fecfactu, "F")
                    FechaVen = DateAdd("d", DBLet(rsVenci!primerve, "N"), FechaVen)
                End If
                
                FecVenci = CDate(FechaVen)
                '++fin
        
                '-------- Primer Vencimiento
                i = 1
                'FECHA VTO
                'FecVenci = CDate(FecVenci)
                'FecVenci = DateAdd("d", DBLet(RsVenci!primerve, "N"), FechaVen)
                '===
        
                '[Monica]17/01/2013: Calculamos la nueva fecha de vencimiento si el cliente tiene dia fijo de pago
                If vsocio.DiaPago <> "" Then
                    FecVenci = NuevaFechaVto(FecVenci, vsocio.DiaPago)
                End If
                
                
                '[Monica]24/01/2013: si la factura es de tpv y la cooperativa es Ribarrojala fecha de vencimiento es la fecha de factura
                If vParamAplic.Cooperativa = 5 Then
                    LetraS = DevuelveDesdeBDNew(cPTours, "stipom", "letraser", "codtipom", "FAT", "T")
                    If LetraS = DBLet(RS!letraser, "T") Then
                        FecVenci = DBLet(RS!fecfactu, "F")
                   End If
                End If
                
                '[Monica]28/12/2015: quitamos lo que hay en el vale
                Dim SqlNuevo As String
                Dim ImporteVale As Currency
                
                SqlNuevo = "select sum(coalesce(importevale,0)) from "
                Select Case vTabla
                    Case "schfac"
                        SqlNuevo = SqlNuevo & " slhfac where " & cadWHERE2
                    Case "schfacr"
                        SqlNuevo = SqlNuevo & " slhfacr where " & cadWHERE2
                    Case "schfac1"
                        SqlNuevo = SqlNuevo & " slhfac1 where " & cadWHERE2
                End Select
                ImporteVale = DevuelveValor(SqlNuevo)
               
               'IMPORTE del Vencimiento
                                                        '[Monica]28/12/2015: le quitamos el importe del vale que va en un registro aparte
                TotalFactura2 = DBLet(RS!TotalFac, "N") - ImporteVale
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFactura2
                Else
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFactura2 Then
                        ImpVenci = Round(ImpVenci + (TotalFactura2 - ImpVenci * rsVenci!numerove), 2)
                    End If
                End If

        
                CadValuesAux2 = "(" & DBSet(RS!letraser, "T") & ", " & DBSet(RS!numfactu, "N") & ", " & DBSet(RS!fecfactu, "F") & ", "
                      
                CadValues2 = CadValuesAux2 & "1," & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(RS!CodForpa, "N") & "," & Format(DBSet(FecVenci, "F"), FormatoFecha) & ","
              

                CodmacBPr = ""
                CodmacBPr = DevuelveDesdeBD("codmacta", "sbanco", "codbanpr", CStr(Banpr), "N")
                
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
                        Set rs4 = New ADODB.Recordset
                        rs4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        
                        If vParamAplic.ContabilidadNueva Then
                            
                            If Not rs4.EOF Then
                                vvIban = MiFormat(DBLet(rs4!IBAN, "T"), "") & MiFormat(DBLet(rs4!codbanco, "N"), "0000") & MiFormat(DBLet(rs4!codsucur, "N"), "0000") & MiFormat(DBLet(rs4!digcontr, "T"), "00") & MiFormat(DBLet(rs4!cuentaba, "T"), "0000000000")
                            
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            Else
                                vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                                
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            End If
                        
                        Else
                            If Not rs4.EOF Then
                                If vEmpresa.HayNorma19_34Nueva = 1 Then
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(rs4!codbanco, "N") & ", " & DBSet(rs4!codsucur, "N") & ", " & DBSet(rs4!digcontr, "T") & ", " & DBSet(rs4!cuentaba, "T") & ", " & DBSet(rs4!IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                Else
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(rs4!codbanco, "N") & ", " & DBSet(rs4!codsucur, "N") & ", " & DBSet(rs4!digcontr, "T") & ", " & DBSet(rs4!cuentaba, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                End If
                            Else
                                If vEmpresa.HayNorma19_34Nueva = 1 Then
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                Else
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                End If
                            End If
                        End If
                    Else
                        If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                        
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        
                        Else
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                               CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            Else
                               CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            End If
                        End If
                    End If
        
                Else    ' facturacion por cliente
                
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                        
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                    
                    Else
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        End If
                    End If
                End If
                
                If vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & ", 1),"
                
                Else
                    CadValues2 = CadValues2 & _
                                 DBSet(textcsb42, "T") & "," & DBSet(textcsb43, "T") & "," & DBSet(textcsb51, "T") & "," & DBSet(textcsb52, "T") & "," & DBSet(textcsb53, "T") & "," & DBSet(textcsb61, "T") & "," & DBSet(textcsb62, "T") & "," & DBSet(textcsb63, "T") & "," & DBSet(textcsb71, "T") & "," & _
                                 DBSet(textcsb72, "T") & "," & DBSet(textcsb73, "T") & "," & DBSet(textcsb81, "T") & "," & DBSet(textcsb82, "T") & "," & DBSet(textcsb83, "T") & ", 1),"
                End If
                
                '[Monica]28/12/2015: dejamos en el ultimo lo correspondiente al importevale
                Dim J As Integer
                J = 2
                             
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For i = 2 To rsVenci!numerove
                    '[Monica]28/12/2015: dejamos en el ultimo lo correspondiente al importevale
                    J = i
                   
                   
                   'FECHA Resto Vencimientos
                    FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                
                    '[Monica]17/01/2013: Calculamos la nueva fecha de vencimiento si el cliente tiene dia fijo de pago
                    If vsocio.DiaPago <> "" Then
                        FecVenci = NuevaFechaVto(FecVenci, vsocio.DiaPago)
                    End If
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    
                    
                    CadValues2 = CadValues2 & CadValuesAux2 & DBSet(i, "N") & "," & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(RS!CodForpa, "N") & "," & Format(DBSet(FecVenci, "F"), FormatoFecha) & ","
                    
                    
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
                            Set rs4 = New ADODB.Recordset
                            rs4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            If vParamAplic.ContabilidadNueva Then
                                If Not rs4.EOF Then
                                    vvIban = MiFormat(DBLet(rs4!IBAN, "T"), "") & MiFormat(DBLet(rs4!codbanco, "N"), "0000") & MiFormat(DBLet(rs4!codsucur, "N"), "0000") & MiFormat(DBLet(rs4!digcontr, "T"), "00") & MiFormat(DBLet(rs4!cuentaba, "T"), "0000000000")
                                
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                Else
                                    vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                                    
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                End If
                            
                            Else
                                If Not rs4.EOF Then
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(rs4!codbanco, "N") & ", " & DBSet(rs4!codsucur, "N") & ", " & DBSet(rs4!digcontr, "T") & ", " & DBSet(rs4!cuentaba, "T") & ", " & DBSet(rs4!IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                Else
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                End If
                            End If
                        Else
                            If vParamAplic.ContabilidadNueva Then
                                vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                                
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            End If
                        End If
            
                    Else    ' facturacion por cliente
                        If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                            
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        End If
            
                    End If
                    
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & ", 1),"
                    Else
                        CadValues2 = CadValues2 & _
                                     DBSet(textcsb42, "T") & "," & DBSet(textcsb43, "T") & "," & DBSet(textcsb51, "T") & "," & DBSet(textcsb52, "T") & "," & DBSet(textcsb53, "T") & "," & DBSet(textcsb61, "T") & "," & DBSet(textcsb62, "T") & "," & DBSet(textcsb63, "T") & "," & DBSet(textcsb71, "T") & "," & _
                                     DBSet(textcsb72, "T") & "," & DBSet(textcsb73, "T") & "," & DBSet(textcsb81, "T") & "," & DBSet(textcsb82, "T") & "," & DBSet(textcsb83, "T") & ", 1),"
                    End If
                
                Next i
                         
                '[Monica]28/12/2015: dejamos en el ultimo lo correspondiente al importevale
                If ImporteVale <> 0 Then
                    J = i
                   
                    FecVenci = CDate(FechaVen)
            
            
                    '[Monica]17/01/2013: Calculamos la nueva fecha de vencimiento si el cliente tiene dia fijo de pago
                    If vsocio.DiaPago <> "" Then
                        FecVenci = NuevaFechaVto(FecVenci, vsocio.DiaPago)
                    End If
                
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = ImporteVale
                    
                    
                    ' la forma de pago es la del vale
                    Dim SqlNuevo1 As String
                    Dim CodForpaVale As Integer
                    CodForpaVale = DevuelveValor("select codforpa from sforpa where tipovale = 1")
                    If vParamAplic.ContabilidadNueva Then
                        SqlNuevo = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", DBLet(CodForpaVale), "N")
                    Else
                        SqlNuevo = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", DBLet(CodForpaVale), "N")
                    End If
                    'si no existe la forma de pago en conta, insertamos la de ariges
                    If SqlNuevo = "" Then
                        'insertamos e sforpa de la CONTA
                        If vParamAplic.ContabilidadNueva Then
                            SQL = "INSERT INTO formapago(codforpa,nomforpa,tipforpa,numerove,primerve,restoven) "
                            SQL = SQL & " select codforpa, nomforpa, tipforpa, numerove, diasvto, restoven "
                            SQL = SQL & " from " & vSesion.CadenaConexion & ".sforpa where codforpa = " & DBSet(CodForpaVale, "N")
                        Else
                            SqlNuevo1 = "tipforpa"
                            SqlNuevo = DevuelveDesdeBDNew(cPTours, "sforpa", "nomforpa", "codforpa", DBLet(CodForpaVale), "N", SqlNuevo1)
                            SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                            SQL = SQL & " VALUES(" & DBSet(CodForpaVale, "N") & ", " & DBSet(SqlNuevo, "T") & ", " & SqlNuevo1 & ")"
                        End If
                        ConnConta.Execute SQL
                    End If
                    
                    CadValues2 = CadValues2 & CadValuesAux2 & DBSet(J, "N") & "," & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(CodForpaVale, "N") & "," & Format(DBSet(FecVenci, "F"), FormatoFecha) & ","
                    
                    
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
                            Set rs4 = New ADODB.Recordset
                            rs4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            
                            If vParamAplic.ContabilidadNueva Then
                                If Not rs4.EOF Then
                                    vvIban = MiFormat(DBLet(rs4!IBAN, "T"), "") & MiFormat(DBLet(rs4!codbanco, "N"), "0000") & MiFormat(DBLet(rs4!codsucur, "N"), "0000") & MiFormat(DBLet(rs4!digcontr, "T"), "00") & MiFormat(DBLet(rs4!cuentaba, "T"), "0000000000")
                                
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                Else
                                    vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                                    
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                End If
                                
                            Else
                                If Not rs4.EOF Then
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(rs4!codbanco, "N") & ", " & DBSet(rs4!codsucur, "N") & ", " & DBSet(rs4!digcontr, "T") & ", " & DBSet(rs4!cuentaba, "T") & ", " & DBSet(rs4!IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                Else
                                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                                End If
                            End If
                        Else
                            If vParamAplic.ContabilidadNueva Then
                                vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                                
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            End If
                        End If
            
                    Else    ' facturacion por cliente
                        If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                            
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                    
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        End If
            
                    End If
                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & ", 1),"
                    
                    Else
                        CadValues2 = CadValues2 & _
                                     DBSet(textcsb42, "T") & "," & DBSet(textcsb43, "T") & "," & DBSet(textcsb51, "T") & "," & DBSet(textcsb52, "T") & "," & DBSet(textcsb53, "T") & "," & DBSet(textcsb61, "T") & "," & DBSet(textcsb62, "T") & "," & DBSet(textcsb63, "T") & "," & DBSet(textcsb71, "T") & "," & _
                                     DBSet(textcsb72, "T") & "," & DBSet(textcsb73, "T") & "," & DBSet(textcsb81, "T") & "," & DBSet(textcsb82, "T") & "," & DBSet(textcsb83, "T") & ", 1),"
                    End If
                End If
                         
                         

                If vsocio.CuentaConta <> "" Then
                    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
                    'forma de pago de la factura. Sino existe insertarla
                    'vemos si existe en la conta
                    If vParamAplic.ContabilidadNueva Then
                        CadValuesAux2 = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", DBLet(RS!CodForpa), "N")
                    Else
                        CadValuesAux2 = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", DBLet(RS!CodForpa), "N")
                    End If
                    'si no existe la forma de pago en conta, insertamos la de ariges
                    If CadValuesAux2 = "" Then
                        If vParamAplic.ContabilidadNueva Then
                            'insertamos e sforpa de la CONTA
                            SQL = "INSERT INTO formapago(codforpa,nomforpa,tipforpa,numerove,primerve,restoven) "
                            SQL = SQL & " select codforpa, nomforpa, tipforpa, numerove, diasvto, restoven "
                            SQL = SQL & " from " & vSesion.CadenaConexion & ".sforpa where codforpa = " & DBSet(RS!CodForpa, "N")
                        
                        Else
                            cadValuesAux = "tipforpa"
                            CadValuesAux2 = DevuelveDesdeBDNew(cPTours, "sforpa", "nomforpa", "codforpa", DBLet(RS!CodForpa), "N", cadValuesAux)
                            'insertamos e sforpa de la CONTA
                            SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                            SQL = SQL & " VALUES(" & DBSet(RS!CodForpa, "N") & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
                        End If
                        ConnConta.Execute SQL
                    End If
        
                    'Insertamos en la tabla scobro de la CONTA
                    If vParamAplic.ContabilidadNueva Then
                        '$$$
                        SQL = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1,"
                        '[Monica]22/11/2013: Tema Iban
                        SQL = SQL & "iban,text33csb , text41csb,"
                        SQL = SQL & "agente) "
                        SQL = SQL & " VALUES " & Mid(CadValues2, 1, Len(CadValues2) - 1)
                    
                    Else
                        SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur, digcontr, cuentaba,"
                        '[Monica]22/11/2013: Tema Iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            SQL = SQL & "iban,text33csb , text41csb,"
                        Else
                            SQL = SQL & "text33csb , text41csb,"
                        End If
                        SQL = SQL & "text42csb, text43csb, text51csb, text52csb, text53csb, text61csb, text62csb, text63csb, text71csb, text72csb, text73csb, text81csb, text82csb, text83csb,agente) "
                        SQL = SQL & " VALUES " & Mid(CadValues2, 1, Len(CadValues2) - 1)
                    End If
                    ConnConta.Execute SQL
                End If
            End If
        End If

    End If

    B = True

EInsertarTesoreria:
    If Err.Number <> 0 Then
        B = False
        MenError = Err.Description
    End If
    InsertarEnTesoreria = B
End Function



Private Sub InsertarError(Cadena As String)
Dim SQL As String

    SQL = "insert into tmperrcomprob values ('" & Cadena & "')"
    Conn.Execute SQL

End Sub


Public Function InsertarCabAsientoDia(Diario As String, Asiento As String, Fecha As String, Obs As String, caderr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    
    Cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
    Cad = Cad & "''," & ValorNulo & "," & DBSet(Obs, "T")
    Cad = "(" & Cad & ")"

    'Insertar en la contabilidad
    SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) "
    SQL = SQL & " VALUES " & Cad
    ConnConta.Execute SQL
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabAsientoDia = False
        caderr = Err.Description
    Else
        InsertarCabAsientoDia = True
    End If
End Function


Public Function InsertarLinAsientoDia(Cad As String, caderr As String) As Boolean
' el Tipo me indica desde donde viene la llamada
' tipo = 0 srecau.codmacta
' tipo = 1 scaalb.codmacta

Dim RS As ADODB.Recordset
Dim Aux As String
Dim SQL As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

 
    SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
    SQL = SQL & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
    SQL = SQL & " VALUES " & Cad
    
    ConnConta.Execute SQL

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinAsientoDia = False
        caderr = Err.Description
    Else
        InsertarLinAsientoDia = True
    End If
End Function

Public Function ActualizarRecaudacion(cadWhere As String, caderr As String) As Boolean
'Poner la factura como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    
    SQL = "UPDATE srecau SET intconta=1 "
    SQL = SQL & " WHERE " & cadWhere

    Conn.Execute SQL
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarRecaudacion = False
        caderr = Err.Description
    Else
        ActualizarRecaudacion = True
    End If
End Function

Public Sub FechasEjercicioConta(FIni As String, FFin As String)
Dim RS As ADODB.Recordset

    On Error GoTo EFechas

    FIni = "Select fechaini,fechafin From parametros"
    Set RS = New ADODB.Recordset
    RS.Open FIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        FIni = DBLet(RS!FechaIni, "F")
        FFin = DBLet(RS!FechaFin, "F")
    End If
    RS.Close
    Set RS = Nothing

EFechas:
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function CrearTMPAsiento() As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPAsiento = False
    
    SQL = "CREATE TEMPORARY TABLE tmpasien ( "
    SQL = SQL & "fecalbar date NOT NULL default '0000-00-00',"
    SQL = SQL & "codturno tinyint(1) NOT NULL default '0',"
    SQL = SQL & "codmacta varchar(10) NOT NULL default ' ',"
    SQL = SQL & "importel decimal(10,2)  NOT NULL default '0.00')"
    Conn.Execute SQL

    CrearTMPAsiento = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPAsiento = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpasien;"
        Conn.Execute SQL
    End If
End Function


Public Function TarjetasInexistentes(SQL As String) As Boolean
Dim cadMen As String

    TarjetasInexistentes = False
    
    SQL = SQL & " and not (scaalb.codsocio, scaalb.numtarje) in (select codsocio, numtarje from starje) "
    
    If (RegistrosAListar(SQL) <> 0) Then
        cadMen = "Hay cargas en las que no es correcta la tarjeta para el socio." & vbCrLf & vbCrLf & _
                 "Revise en el mantenimiento de albaranes." & vbCrLf & vbCrLf
        MsgBox cadMen, vbExclamation
        TarjetasInexistentes = True
    End If
End Function

Public Function ComprobarNumFacturas_new(cadTabla As String, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim SQL As String
Dim SQLconta As String
Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim B As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas_new = False
    
    If vParamAplic.ContabilidadNueva Then
        SQLconta = "SELECT count(*) FROM factcli WHERE "
    Else
        SQLconta = "SELECT count(*) FROM cabfact WHERE "
    End If
        

    
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText

'    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        SQL = "SELECT DISTINCT " & cadTabla & ".codtipom,letraser,facturas.numfactu,facturas.fecfactu "
        SQL = SQL & " FROM (" & cadTabla & " INNER JOIN usuarios.stipom stipom ON " & cadTabla & ".codtipom=stipom.codtipom) "
        SQL = SQL & " INNER JOIN tmpFactu ON facturas.codtipom=tmpFactu.codtipom AND facturas.numfactu=tmpFactu.numfactu AND facturas.fecfactu=tmpFactu.fecfactu "

        
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        B = True
        While Not RS.EOF And B
            If vParamAplic.ContabilidadNueva Then
                SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND numfactu=" & DBSet(RS!numfactu, "N") & " AND anofactu=" & Year(RS!fecfactu) & ")"
            
            Else
                SQL = "(numserie= " & DBSet(RS!letraser, "T") & " AND codfaccl=" & DBSet(RS!numfactu, "N") & " AND anofaccl=" & Year(RS!fecfactu) & ")"
            End If
'            If SituarRSetMULTI(RSconta, SQL) Then
            SQL = SQLconta & SQL
            If RegistrosAListar(SQL, cConta) Then
                B = False
                SQL = "          Letra Serie: " & DBSet(RS!letraser, "T") & vbCrLf
                SQL = SQL & "          Nº Fac.: " & Format(RS!numfactu, "0000000") & vbCrLf
                SQL = SQL & "          Fecha: " & RS!fecfactu
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        
        If Not B Then
            SQL = "Ya existe la factura: " & vbCrLf & SQL
            SQL = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
            
            MsgBox SQL, vbExclamation
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

Public Function ComprobarCtaContable_new(cadTabla As String, Opcion As Byte) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim B As Boolean
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
        SQL = DevuelveDesdeBDNew(cConta, "parametros", "grupogto", "", "", "", cadG)
        If SQL <> "" And cadG <> "" Then
            SQL = " AND (codmacta like '" & SQL & "%' OR codmacta like '" & cadG & "%')"
        ElseIf SQL <> "" Then
            SQL = " AND (codmacta like '" & SQL & "%')"
        ElseIf cadG <> "" Then
            SQL = " AND (codmacta like '" & cadG & "%')"
        End If
        cadG = SQL
    End If
    
    
'    SQL = "SELECT codmacta FROM cuentas "
'    SQL = SQL & " WHERE apudirec='S'"
'    If cadG <> "" Then SQL = SQL & cadG
    
    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG
    
    If Opcion = 1 Then
        If cadTabla = "facturas" Then
            'Seleccionamos los distintos clientes,cuentas que vamos a facturar
            SQL = "SELECT DISTINCT facturas.codclien, clientes.codmacta "
            SQL = SQL & " FROM (facturas INNER JOIN clientes ON facturas.codclien=clientes.codclien) "
            SQL = SQL & " INNER JOIN tmpFactu ON facturas.codtipom=tmpFactu.codtipom AND facturas.numfactu=tmpFactu.numfactu AND facturas.fecfactu=tmpFactu.fecfactu "
        Else
            If cadTabla = "scafpc" Then
                'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
                SQL = "SELECT DISTINCT scafpc.codprove, proveedor.codmacta "
                SQL = SQL & " FROM (scafpc INNER JOIN proveedor ON scafpc.codprove=proveedor.codprove) "
                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
            Else
                'Seleccionamos los distintos transportistas ,cuentas que vamos a facturar
                SQL = "SELECT DISTINCT tcafpc.codtrans, agencias.codmacta "
                SQL = SQL & " FROM (tcafpc INNER JOIN agencias ON tcafpc.codtrans=agencias.codtrans) "
                SQL = SQL & " INNER JOIN tmpFactu ON tcafpc.codtrans=tmpFactu.codtrans AND tcafpc.numfactu=tmpFactu.numfactu AND tcafpc.fecfactu=tmpFactu.fecfactu "
            
            End If
        End If
    ElseIf Opcion = 2 Or Opcion = 3 Or Opcion = 8 Then
        SQL = "SELECT distinct "
        If Opcion = 2 Then SQL = SQL & " sartic.codartic,"
        If cadTabla = "facturas" Then
            If Opcion <> 8 Then
                SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((facturas_envases "
                SQL = SQL & " INNER JOIN tmpFactu ON facturas_envases.codtipom=tmpFactu.codtipom AND facturas_envases.numfactu=tmpFactu.numfactu AND facturas_envases.fecfactu=tmpFactu.fecfactu) "
                SQL = SQL & "INNER JOIN sartic ON facturas_envases.codartic=sartic.codartic) "
            Else
                numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", vParamAplic.NumeroConta, "N")
                NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel, "codempre", vParamAplic.NumeroConta, "N")
                NumDigit3 = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
                
'                CadCampo1 = "concat(concat(variedades.raizctavtas,tipomer.digicont), right(concat('0000000000',albaran_variedad.codvarie)," & (CCur(NumDigit) - CCur(NumDigit3) - 1) & "))"
                CadCampo1 = "CASE tipomer.tiptimer WHEN 0 THEN ctavtasinterior WHEN 1 THEN ctavtasexportacion WHEN 2 THEN ctavtasindustria WHEN 3 THEN ctavtasretirada WHEN 4 THEN ctavtasotros END"
                
                SQL = SQL & " albaran_variedad.codvarie, " & CadCampo1 & " as codmacta from ((((((facturas_variedad "
                SQL = SQL & " INNER JOIN tmpFactu ON facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu) "
                SQL = SQL & " inner join usuarios.stipom stipom on facturas_variedad.codtipom=stipom.codtipom) "
                SQL = SQL & " inner join albaran on facturas_variedad.numalbar = albaran.numalbar) "
                SQL = SQL & " inner join tipomer on albaran.codtimer = tipomer.codtimer) "
                SQL = SQL & " inner join albaran_variedad on facturas_variedad.numalbar = albaran_variedad.numalbar and facturas_variedad.numlinealbar = albaran_variedad.numlinea) "
                SQL = SQL & " inner join variedades on albaran_variedad.codvarie=variedades.codvarie) "
                
                
'                Sql = Sql & " INNER JOIN tmpFactu ON facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu) "
'                Sql = Sql & "INNER JOIN sartic ON facturas_envases.codartic=sartic.codartic) "
            End If
        Else
            SQL = SQL & " sartic.ctacompr as codmacta from ((slifpc "
            SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
            SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
        End If
'        If Opcion <> 8 Then Sql = Sql & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
    ElseIf Opcion = 4 Or Opcion = 6 Then
'        Sql = "select distinct " & DBSet(vParamAplic.CtaTraReten, "T") & " as codmacta from tcafpc "
    ElseIf Opcion = 5 Or Opcion = 7 Then
'        Sql = "select distinct " & DBSet(vParamAplic.CtaAboTrans, "T") & " as codmacta from tcafpc "
'       transporte
            SQL = " SELECT if(tipomer.tiptimer = 1,variedades.ctatraexporta,variedades.ctatrainterior) as cuenta "
            SQL = SQL & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer, tmpFactu, tcafpc  WHERE "
            SQL = SQL & " tcafpc.tipo = 0 and " ' transportista
            SQL = SQL & " tlifpc.codtrans=tmpFactu.codtrans and tlifpc.numfactu=tmpFactu.numfactu and tlifpc.fecfactu=tmpFactu.fecfactu and "
            SQL = SQL & " tlifpc.numalbar=albaran_variedad.numalbar and "
            SQL = SQL & " tlifpc.numlinea=albaran_variedad.numlinea and "
            SQL = SQL & " tlifpc.codtrans=tcafpc.codtrans and tlifpc.numfactu=tcafpc.numfactu and tlifpc.fecfactu=tcafpc.fecfactu and "
            SQL = SQL & " albaran_variedad.numalbar=albaran.numalbar and "
            SQL = SQL & " albaran_variedad.codvarie=variedades.codvarie and "
            SQL = SQL & " albaran.codtimer=tipomer.codtimer "
            SQL = SQL & " group by 1 "

    ElseIf Opcion = 12 Or Opcion = 13 Then
'       comisionista
            SQL = " SELECT variedades.ctacomisionista as cuenta, variedades.codvarie  "
            SQL = SQL & " FROM tlifpc, albaran, albaran_variedad, variedades, tipomer, tmpFactu, tcafpc  WHERE "
            SQL = SQL & " tcafpc.tipo = 1 and " ' comisionista
            SQL = SQL & " tlifpc.codtrans=tmpFactu.codtrans and tlifpc.numfactu=tmpFactu.numfactu and tlifpc.fecfactu=tmpFactu.fecfactu and "
            SQL = SQL & " tlifpc.numalbar=albaran_variedad.numalbar and "
            SQL = SQL & " tlifpc.numlinea=albaran_variedad.numlinea and "
            SQL = SQL & " tlifpc.codtrans=tcafpc.codtrans and tlifpc.numfactu=tcafpc.numfactu and tlifpc.fecfactu=tcafpc.fecfactu and "
            SQL = SQL & " albaran_variedad.numalbar=albaran.numalbar and "
            SQL = SQL & " albaran_variedad.codvarie=variedades.codvarie and "
            SQL = SQL & " albaran.codtimer=tipomer.codtimer "
            SQL = SQL & " group by 1 "
            
    ElseIf Opcion = 9 Or Opcion = 10 Then
            SQL = " select codmacta as cuenta "
            SQL = SQL & " from tcafpv, tmpFactu "
            SQL = SQL & " where tmpFactu.codtrans=tcafpv.codtrans and tmpFactu.numfactu=tcafpv.numfactu and tmpFactu.fecfactu=tcafpv.fecfactu "
            SQL = SQL & " group by 1 "
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    B = True

    While Not RS.EOF And B
        If Opcion < 4 Or Opcion = 8 Then
            SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!codmacta, "T")
        ElseIf Opcion = 4 Or Opcion = 6 Then
'            Sql = SQLcuentas & " AND codmacta= " & DBSet(vParamAplic.CtaTraReten, "T")
        ElseIf Opcion = 5 Or Opcion = 7 Then
            SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!Cuenta, "T")
        ElseIf Opcion = 12 Or Opcion = 13 Then
            SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!Cuenta, "T")
        ElseIf Opcion = 9 Or Opcion = 10 Then
            SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!Cuenta, "T")
        End If
            
        
        If Not (RegistrosAListar(SQL, cConta) > 0) Then
        'si no lo encuentra
            B = False 'no encontrado
            If Opcion = 1 Then
                If cadTabla = "facturas" Then
                    SQL = RS!codmacta & " del Cliente " & Format(RS!CodClien, "000000")
                Else
                    If cadTabla = "scafpc" Then
                        SQL = RS!codmacta & " del Proveedor " & Format(RS!CodProve, "000000")
                    Else
                        SQL = RS!codmacta & " del Transportista " & Format(RS!codTrans, "000")
                    End If
                End If
            ElseIf Opcion = 2 Then
                SQL = RS!codmacta & " del articulo " & Format(RS!codArtic, "000000")
            ElseIf Opcion = 3 Then
                SQL = RS!codmacta
            ElseIf Opcion = 4 Or Opcion = 6 Then
'                Sql = vParamAplic.CtaTraReten
            ElseIf Opcion = 5 Or Opcion = 7 Then
                SQL = DBLet(RS!Cuenta, "T") ' vParamAplic.CtaAboTrans
            ElseIf Opcion = 12 Or Opcion = 13 Then
                SQL = DBLet(RS!Cuenta, "T") & " de comisionista de la variedad " & Format(RS!codvarie, "000000")
            ElseIf Opcion = 8 Then
                SQL = RS!codmacta & " de la variedad " & Format(RS!codvarie, "0000")
            ElseIf Opcion = 9 Or Opcion = 10 Then
                SQL = DBLet(RS!Cuenta, "T") ' vParamAplic.CtaAboTrans
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
        RS.MoveNext
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
        
        
        
        If Not B Then
            If Not (Opcion = 3 Or Opcion = 6 Or Opcion = 7) Then
                SQL = "No existe la cta contable " & SQL
            Else
                SQL = "La cuenta " & SQL & " no es del nivel correcto. "
                If Opcion = 3 Then SQL = SQL & "(Familias de artículos)."
            End If
            SQL = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & SQL
            
            MsgBox SQL, vbExclamation
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




Public Function ComprobarCCoste_new(cadCC As String, cadTabla As String, Optional Opcion As Byte) As Boolean
Dim SQL As String
Dim sql2 As String
Dim RS As ADODB.Recordset
Dim B As Boolean

    On Error GoTo ECCoste

    ComprobarCCoste_new = False
    Select Case cadTabla
        Case "facturas" ' facturas de venta
            Select Case Opcion
                Case 1
                    SQL = "select distinct variedades.codccost from facturas_variedad, albaran_variedad, variedades, tmpFactu where "
                    SQL = SQL & " albaran_variedad.codvarie=variedades.codvarie and "
                    SQL = SQL & " facturas_variedad.codtipom=tmpFactu.codtipom AND facturas_variedad.numfactu=tmpFactu.numfactu AND facturas_variedad.fecfactu=tmpFactu.fecfactu and  "
                    SQL = SQL & " albaran_variedad.numalbar = facturas_variedad.numalbar and "
                    SQL = SQL & " albaran_variedad.numlinea = facturas_variedad.numlinealbar "
                Case 2
                    SQL = " select distinct sfamia.codccost from facturas_envases, sartic, sfamia, tmpFactu where "
                    SQL = SQL & " facturas_envases.codtipom=tmpFactu.codtipom AND facturas_envases.numfactu=tmpFactu.numfactu AND facturas_envases.fecfactu=tmpFactu.fecfactu and  "
                    SQL = SQL & " facturas_envases.codartic = sartic.codartic and "
                    SQL = SQL & " sartic.codfamia = sfamia.codfamia "
                Case 3
'                    If HayFacturasACuenta Then
'                        Sql = " select '" & vParamAplic.CCosteFraACta & "' as codccost from tmpFactu where tmpfactu.codtipom = 'EAC' "
'                    Else
'                        ComprobarCCoste_new = True
'                        Exit Function
'                    End If
            End Select
        Case "scafpc" ' facturas de compra
            SQL = " select distinct sfamia.codccost from slifpc, sartic, sfamia, tmpFactu where "
            SQL = SQL & " slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu and  "
            SQL = SQL & " slifpc.codartic = sartic.codartic and "
            SQL = SQL & " sartic.codfamia = sfamia.codfamia "
        
        Case "tcafpc" ' facturas de transporte
            SQL = "select distinct variedades.codccost from tlifpc, albaran_variedad, variedades, tmpFactu where "
            SQL = SQL & " albaran_variedad.codvarie=variedades.codvarie and "
            SQL = SQL & " tlifpc.codtrans=tmpFactu.codtrans AND tlifpc.numfactu=tmpFactu.numfactu AND tlifpc.fecfactu=tmpFactu.fecfactu and  "
            SQL = SQL & " albaran_variedad.numalbar = tlifpc.numalbar and "
            SQL = SQL & " albaran_variedad.numlinea = tlifpc.numlinea "
    
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    B = True

    While Not RS.EOF And B
        If DBLet(RS.Fields(0).Value, "T") = "" Then
            B = False
        Else
            SQL = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", RS.Fields(0).Value, "T")
            If SQL = "" Then
                B = False
                sql2 = "Centro de Coste: " & RS.Fields(0)
            End If
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
    If Not B Then
        SQL = "No existe el " & sql2
        SQL = "Comprobando Centros de Coste en contabilidad..." & vbCrLf & vbCrLf & SQL
    
        MsgBox SQL, vbExclamation
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
Dim SQL As String
Dim sql2 As String
Dim RS As ADODB.Recordset
Dim B As Boolean

    On Error GoTo ECCoste

    ComprobarFormadePago = False
    SQL = "select distinct facturas.codforpa from facturas, tmpFactu where "
    SQL = SQL & " facturas.codtipom=tmpFactu.codtipom AND facturas.numfactu=tmpFactu.numfactu AND facturas.fecfactu=tmpFactu.fecfactu  "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    B = True

    While Not RS.EOF And B
        If vParamAplic.ContabilidadNueva Then
            SQL = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", RS.Fields(0).Value, "N")
        Else
            SQL = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", RS.Fields(0).Value, "N")
        End If
        If SQL = "" Then
            B = False
            sql2 = "Formas de Pago: " & RS.Fields(0)
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
    If Not B Then
        SQL = "No existe la " & sql2
        SQL = "Comprobando Formas de Pago en contabilidad..." & vbCrLf & vbCrLf & SQL
    
        MsgBox SQL, vbExclamation
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



Public Function PasarFacturaProv(cadWhere As String, CodCCost As String, FechaFin As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariges.scafpc --> conta.cabfactprov
' ariges.slifpc --> conta.linfactprov
'Actualizar la tabla ariges.scafpc.inconta=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim SQL As String
Dim Mc As Contadores


    On Error GoTo EContab

    ConnConta.BeginTrans
    Conn.BeginTrans
        
    
    Set Mc = New Contadores
    
    '---- Insertar en la conta Cabecera Factura
    B = InsertarCabFactProv(cadWhere, cadMen, Mc, FechaFin)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If B Then
        CCoste = CodCCost
        '---- Insertar lineas de Factura en la Conta
        B = InsertarLinFact_new("scafpc", cadWhere, cadMen, Mc.Contador)
        cadMen = "Insertando Lin. Factura: " & cadMen

        If B Then
            '---- Poner intconta=1 en ariges.scafac
            B = ActualizarCabFact("scafpc", cadWhere, cadMen)
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
        B = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        Conn.CommitTrans
        PasarFacturaProv = True
    Else
        ConnConta.RollbackTrans
        Conn.RollbackTrans
        PasarFacturaProv = False
        If Not B Then
            InsertarTMPErrFac cadMen, cadWhere
'            SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
'            SQL = SQL & " Select *," & DBSet(Mid(cadMen, 1, 200), "T") & " as error From tmpFactu "
'            SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
'            Conn.Execute SQL
        End If
    End If
End Function


Private Function InsertarCabFactProv(cadWhere As String, caderr As String, ByRef Mc As Contadores, FechaFin As String) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim Intracom As Integer

    On Error GoTo EInsertar
       
    
    SQL = " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,proveedor.codmacta,"
    SQL = SQL & "scafpc.dtoppago,scafpc.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    SQL = SQL & "totalfac,tipoiva1,tipoiva2,tipoiva3,proveedor.codprove, proveedor.nomprove, proveedor.tipprove "
    SQL = SQL & " FROM " & "scafpc "
    SQL = SQL & "INNER JOIN " & "proveedor ON scafpc.codprove=proveedor.codprove "
    SQL = SQL & " WHERE " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not RS.EOF Then
    
        If Mc.ConseguirContador("1", (RS!FecRecep <= CDate(FechaFin) - 365), True) = 0 Then
        
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = RS!DtoPPago
            DtoGnral = RS!DtoGnral
            BaseImp = RS!BaseIVA1 + CCur(DBLet(RS!BaseIVA2, "N")) + CCur(DBLet(RS!BaseIVA3, "N"))
            TotalFac = RS!TotalFac
            AnyoFacPr = RS!anofacpr
            
            Intracom = DBLet(RS!tipprove, "N")
            If Intracom = 2 Then Intracom = 0
            
            Nulo2 = "N"
            Nulo3 = "N"
            If DBLet(RS!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(RS!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            SQL = ""
            SQL = Mc.Contador & "," & DBSet(RS!fecfactu, "F") & "," & RS!anofacpr & "," & DBSet(RS!FecRecep, "F") & "," & DBSet(RS!numfactu, "T") & "," & DBSet(RS!codmacta, "T") & "," & ValorNulo & ","
            SQL = SQL & DBSet(RS!BaseIVA1, "N") & "," & DBSet(RS!BaseIVA2, "N", "S") & "," & DBSet(RS!BaseIVA3, "N", "S") & ","
            SQL = SQL & DBSet(RS!porciva1, "N") & "," & DBSet(RS!porciva2, "N", Nulo2) & "," & DBSet(RS!porciva3, "N", Nulo3) & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(RS!impoiva1, "N") & "," & DBSet(RS!impoiva2, "N", Nulo2) & "," & DBSet(RS!impoiva3, "N", Nulo3) & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            SQL = SQL & DBSet(RS!TotalFac, "N") & "," & DBSet(RS!TipoIVA1, "N") & "," & DBSet(RS!TipoIVA2, "N", Nulo2) & "," & DBSet(RS!TipoIVA3, "N", Nulo3) & "," & DBSet(Intracom, "N") & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(RS!FecRecep, "F") & ",0"
            Cad = Cad & "(" & SQL & ")"
            
            'Insertar en la contabilidad
            SQL = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
            SQL = SQL & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
            SQL = SQL & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
            SQL = SQL & " VALUES " & Cad
            ConnConta.Execute SQL
            
            'añadido como david para saber que numero de registro corresponde a cada factura
            'Para saber el numreo de registro que le asigna a la factrua
            SQL = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vSesion.Codigo & "," & Mc.Contador
            SQL = SQL & ",'" & DevNombreSQL(RS!numfactu) & " @ " & Format(RS!fecfactu, "dd/mm/yyyy") & "','" & DevNombreSQL(RS!NomProve) & "'," & RS!CodProve & ")"
            Conn.Execute SQL
            
            
        End If
    End If
    RS.Close
    Set RS = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactProv = False
        caderr = Err.Description
    Else
        InsertarCabFactProv = True
    End If
End Function



Private Function InsertarLinFact_new(cadTabla As String, cadWhere As String, caderr As String, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SqlAux As String
Dim sql2 As String
Dim RS As ADODB.Recordset
Dim Cad As String, Aux As String
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
    

    If cadTabla = "scafpc" Then 'COMPRAS
        'utilizamos sfamia.ctaventa o sfamia.aboventa
        If TotalFac >= 0 Then
            cadCampo = "sartic.ctacompr"
        Else
            cadCampo = "sartic.ctacompr"
        End If
        TieneAnalitica = "0"
        TieneAnalitica = DevuelveDesdeBDNew(cConta, "parametros", "autocoste", "", "")
        If TieneAnalitica = "1" Then  'hay contab. analitica
            SQL = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe, sartic.codccost"
        Else
            SQL = " SELECT slifpc.codprove,numfactu,fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
        End If
        SQL = SQL & " FROM (slifpc  "
        SQL = SQL & " inner join sartic on slifpc.codartic=sartic.codartic) "
        SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "slifpc")
        SQL = SQL & " GROUP BY " & cadCampo
        
        If TieneAnalitica = "1" Then
            SQL = SQL & ", sartic.codccost "
        End If
    End If
  
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    totimp = 0
    SqlAux = ""
    While Not RS.EOF
        SqlAux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        totimp = totimp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        sql2 = ""
        
        If cadTabla = "facturas" Then 'VENTAS a clientes
            SQL = "'" & RS!letraser & "'," & RS!numfactu & "," & Year(RS!fecfactu) & "," & i & ","
            SQL = SQL & DBSet(RS!Cuenta, "T")
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
            If cadTabla = "scafpc" Then 'COMPRAS
                'Laura 24/10/2006
                'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
                SQL = numRegis & "," & AnyoFacPr & "," & i & ","
                
    '            If ImpLinea >= 0 Then
                    SQL = SQL & DBSet(RS!Cuenta, "T")
    '            Else
    '                SQL = SQL & DBSet(RS!abocompr, "T")
    '            End If
            Else 'TRANSPORTE
                SQL = numRegis & "," & AnyoFacPr & "," & i & ","
                SQL = SQL & DBSet(RS!Cuenta, "T")
            End If
        End If
        
        sql2 = SQL & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
        
        If TieneAnalitica = "1" Then
            If cadTabla = "tcafpc" Then
                If DBLet(RS!CodCCost, "T") = "----" Then
                    SQL = SQL & DBSet(CCoste, "T")
                Else
                    SQL = SQL & DBSet(RS!CodCCost, "T")
                    CCoste = DBLet(RS!CodCCost, "T")
                End If
            Else
                SQL = SQL & DBSet(RS!CodCCost, "T")
                CCoste = DBLet(RS!CodCCost, "T")
            End If
        Else
            SQL = SQL & ValorNulo
            CCoste = ValorNulo
        End If
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If totimp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        totimp = BaseImp - totimp
        totimp = ImpLinea + totimp '(+- diferencia)
        sql2 = sql2 & DBSet(totimp, "N") & ","
        If CCoste = "" Or CCoste = ValorNulo Then
            sql2 = sql2 & ValorNulo
        Else
            sql2 = sql2 & DBSet(CCoste, "T")
        End If
        If SqlAux <> "" Then 'hay mas de una linea
            Cad = SqlAux & "(" & sql2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & sql2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If


    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTabla = "facturas" Then
            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            SQL = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        SQL = SQL & " VALUES " & Cad
        ConnConta.Execute SQL
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_new = False
        caderr = Err.Description
    Else
        InsertarLinFact_new = True
    End If
End Function


Private Sub InsertarTMPErrFac(MenError As String, cadWhere As String)
Dim SQL As String

    On Error Resume Next
    SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    SQL = SQL & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
    Conn.Execute SQL
    
    If Err.Number <> 0 Then Err.Clear
End Sub



' ### [Monica] 02/10/2006
' copiado de la clase de laura cfactura
Public Function InsertarEnTesoreriaDB(db As BaseDatos, cadWhere As String, ByVal FecVenci As String, Banpr As String, MenError As String, ByRef vsocio As CSocio, vTabla As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim B As Boolean
Dim RS As ADODB.Recordset
Dim RSx As ADODB.Recordset
Dim SQL As String, textcsb33 As String, textcsb41 As String
Dim sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Sql5 As String
Dim Rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
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

    B = False
    InsertarEnTesoreriaDB = False
    CadValues = ""
    CadValues2 = ""

    SQL = "select * from " & vTabla & " where  " & cadWhere
    
    Set RS = db.cursor(SQL)
    
    If Not RS.EOF Then
    
        textcsb33 = "'FACT: " & DBLet(RS!letraser, "T") & "-" & Format(DBLet(RS!numfactu, "N"), "0000000") & " " & Format(DBLet(RS!fecfactu, "F"), "dd/mm/yy")
        textcsb33 = textcsb33 & " de " & DBSet(RS!TotalFac, "N") & "'"
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
            cadWHERE2 = Replace(cadWhere, "schfac", "slhfac")
        Else
            cadWHERE2 = Replace(cadWhere, "schfacr", "slhfacr")
        End If


        
'[Monica]08/01/2014: lo cambiamos rellenando lo maximo que podemos
        If vParamAplic.Cooperativa = 5 Then
            Dim cad1 As String
            Dim cad2 As String
            Dim cad22 As String
            
            SQL = "select count(distinct numalbar) from " & vTabla & " where " & cadWhere
            cad1 = ""
            sql2 = "select numalbar, fecalbar, sum(implinea) "
            Select Case vTabla
                Case "schfac"
                    sql2 = sql2 & " from slhfac where " & cadWHERE2
                Case "schfacr"
                    sql2 = sql2 & " from slhfacr where " & cadWHERE2
                Case "schfac1"
                    sql2 = sql2 & " from slhfac1 where " & cadWHERE2
            End Select

            sql2 = sql2 & " group by numalbar, fecalbar order by numalbar, fecalbar "
            
            Set RSx = New ADODB.Recordset
            RSx.Open sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        FacturaFP = DevuelveDesdeBDNew(cPTours, "ssocio", "facturafp", "codsocio", RS!codsocio, "N")
        If CInt(FacturaFP) = 1 Then
            Ndias = ""
            Ndias = DevuelveDesdeBDNew(cPTours, "sforpa", "diasvto", "codforpa", RS!CodForpa, "N")
            Ndias = ComprobarCero(Ndias)
            FecVenci1 = CDate(DBLet(RS!fecfactu, "F")) + CCur(Ndias)
            FecVenci = CDate(FecVenci1)
        End If
        'fin 01/06/2007
        
        CadValuesAux2 = "(" & DBSet(RS!letraser, "T") & ", " & DBSet(RS!numfactu, "N") & ", " & DBSet(RS!fecfactu, "F") & ", "
              
        CadValues2 = CadValuesAux2 & "1," & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(RS!CodForpa, "N") & "," & Format(DBSet(FecVenci, "F"), FormatoFecha) & ","
              
' 01/06/2006 he quitado esta instruccion
'        'FECHA VTO
'        FecVenci = CDate(FecVenci)

        ImpVenci = DBLet(RS!TotalFac, "N")
        CodmacBPr = ""
        CodmacBPr = DevuelveDesdeBD("codmacta", "sbanco", "codbanpr", CStr(Banpr), "N")
        
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
                Set rs4 = New ADODB.Recordset
                rs4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If vParamAplic.ContabilidadNueva Then
                    If Not rs4.EOF Then
                        vvIban = MiFormat(rs4!IBAN, "") & MiFormat(DBLet(rs4!codbanco), "0000") & MiFormat(DBLet(rs4!codsucur), "0000") & MiFormat(DBLet(rs4!digcontr), "00") & MiFormat(DBLet(rs4!cuentaba), "0000000000")
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                    Else
                        vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                    End If
                Else
                    If Not rs4.EOF Then
                        If vEmpresa.HayNorma19_34Nueva Then
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(rs4!codbanco, "N") & ", " & DBSet(rs4!codsucur, "N") & ", " & DBSet(rs4!digcontr, "T") & ", " & DBSet(rs4!cuentaba, "T") & ", " & DBSet(rs4!IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(rs4!codbanco, "N") & ", " & DBSet(rs4!codsucur, "N") & ", " & DBSet(rs4!digcontr, "T") & ", " & DBSet(rs4!cuentaba, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        End If
                    Else
                        If vEmpresa.HayNorma19_34Nueva Then
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        End If
                    End If
                End If
            Else
                If vParamAplic.ContabilidadNueva Then
                    vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                Else
                    If vEmpresa.HayNorma19_34Nueva Then
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                    Else
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                    End If
                End If
            End If

        Else    ' facturacion por cliente
            If vParamAplic.ContabilidadNueva Then
                vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
            Else
                If vEmpresa.HayNorma19_34Nueva Then
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                Else
                    CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                End If
            End If
        End If
        
        If vParamAplic.ContabilidadNueva Then
            CadValues2 = CadValues2 & ", 1)"
        
        Else
            CadValues2 = CadValues2 & _
                         DBSet(textcsb42, "T") & "," & DBSet(textcsb43, "T") & "," & DBSet(textcsb51, "T") & "," & DBSet(textcsb52, "T") & "," & DBSet(textcsb53, "T") & "," & DBSet(textcsb61, "T") & "," & DBSet(textcsb62, "T") & "," & DBSet(textcsb63, "T") & "," & DBSet(textcsb71, "T") & "," & _
                         DBSet(textcsb72, "T") & "," & DBSet(textcsb73, "T") & "," & DBSet(textcsb81, "T") & "," & DBSet(textcsb82, "T") & "," & DBSet(textcsb83, "T") & ", 1)"
        End If
        If vsocio.CuentaConta <> "" Then
            'antes de grabar en la scobro comprobar que existe en conta.sforpa la
            'forma de pago de la factura. Sino existe insertarla
            'vemos si existe en la conta
            If vParamAplic.ContabilidadNueva Then
                CadValuesAux2 = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", DBLet(RS!CodForpa), "N")
            Else
                CadValuesAux2 = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", DBLet(RS!CodForpa), "N")
            End If
            'si no existe la forma de pago en conta, insertamos la de ariges
            If CadValuesAux2 = "" Then
                If vParamAplic.ContabilidadNueva Then
                    SQL = "INSERT INTO formapago(codforpa,nomforpa,tipforpa,numerove, primerve,restoven) "
                    SQL = SQL & " select codforpa, nomforpa, tipforpa, numerove, diasvto, restoven "
                    SQL = SQL & " from " & vSesion.CadenaConexion & ".sforpa where codforpa = " & DBSet(RS!CodForpa, "N")
                Else
                    cadValuesAux = "tipforpa"
                    CadValuesAux2 = DevuelveDesdeBDNew(cPTours, "sforpa", "nomforpa", "codforpa", DBLet(RS!CodForpa), "N", cadValuesAux)
                    'insertamos e sforpa de la CONTA
                    SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                    SQL = SQL & " VALUES(" & DBSet(RS!CodForpa, "N") & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
                End If
                ConnConta.Execute SQL
            End If

            If vParamAplic.ContabilidadNueva Then
                'Insertamos en la tabla scobro de la CONTA
                SQL = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, "
                SQL = SQL & "iban, text33csb , text41csb, agente) "
                SQL = SQL & " VALUES " & CadValues2
            
            Else
                'Insertamos en la tabla scobro de la CONTA
                SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur, digcontr, cuentaba,"
                    '[Monica]22/11/2013: Iban
                If vEmpresa.HayNorma19_34Nueva Then
                    SQL = SQL & "iban, text33csb , text41csb, "
                Else
                    SQL = SQL & "text33csb , text41csb, "
                End If
                SQL = SQL & "text42csb, text43csb, text51csb, text52csb, text53csb, text61csb, text62csb, text63csb, text71csb, text72csb, text73csb, text81csb, text82csb, text83csb,agente) "
                SQL = SQL & " VALUES " & CadValues2
                
            End If
                
            ConnConta.Execute SQL
        End If


    End If

    B = True

EInsertarTesoreria:
    If Err.Number <> 0 Then
        B = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaDB = B
End Function


Private Function NuevaFechaVto(vFecVenci As Date, DiaPago As Integer) As Date
Dim NewFec As String
Dim Dia As Integer
Dim mes As Integer
Dim Anyo As Integer
    
    On Error Resume Next
    
    
    NuevaFechaVto = vFecVenci
    
    Dia = Day(vFecVenci)
    mes = Month(vFecVenci)
    Anyo = Year(vFecVenci)
    
    If DiaPago <= Dia Then
        mes = mes + 1
        If mes > 12 Then
            mes = 1
            Anyo = Anyo + 1
        End If
        Dia = CInt(DiaPago)
    Else
        Dia = CInt(DiaPago)
    End If
    
    NewFec = Format(Dia, "00") & "/" & Format(mes, "00") & "/" & Format(Anyo, "0000")
    
    '31
    If Not EsFechaOK(NewFec) Then
        Dia = Dia - 1
        NewFec = Format(Dia, "00") & "/" & Format(mes, "00") & "/" & Format(Anyo, "0000")
    End If
    '30
    If Not IsDate(NewFec) Then
        Dia = Dia - 1
        NewFec = Format(Dia, "00") & "/" & Format(mes, "00") & "/" & Format(Anyo, "0000")
    End If
    '29
    If Not IsDate(NewFec) Then
        Dia = Dia - 1
        NewFec = Format(Dia, "00") & "/" & Format(mes, "00") & "/" & Format(Anyo, "0000")
    End If
    NuevaFechaVto = CDate(NewFec)

End Function


'[Monica]29/06/2016: He creado una nueva funcion partiendo de InsertarEnTesoreria para las ajenas de Regaixo pq quieren un cobro por tarjeta
'                    ya no jugamos con el total de la factura - importe de vale

Public Function InsertarEnTesoreriaAjenas(cadWhere As String, ByVal FechaVen As String, Banpr As String, MenError As String, ByRef vsocio As CSocio, vTabla As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim B As Boolean
Dim RS As ADODB.Recordset
Dim RSx As ADODB.Recordset
Dim SQL As String, textcsb33 As String, textcsb41 As String
Dim sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Sql5 As String
Dim Rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
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
Dim FecVenci As Date
Dim rsVenci As ADODB.Recordset
Dim TotalFactura2 As Currency

Dim LetraS As String
Dim CodForpa As Integer

    On Error GoTo EInsertarTesoreria

    B = False
    InsertarEnTesoreriaAjenas = False
    CadValues = ""
    CadValues2 = ""

'    SQL = "select * from " & vTabla & " where  " & cadWhere
' ahora
    SQL = "select letraser, numfactu, fecfactu, iban, codbanco, codsucur, digcontr, cuentaba, "
    SQL = SQL & " sum(implinea - coalesce(importevale,0)) totalfac, sum(coalesce(importevale,0)) importevale "
    SQL = SQL & " from starje, slhfacr "
    SQL = SQL & " where " & Replace(cadWhere, "schfacr", "slhfacr")
    SQL = SQL & " and starje.numtarje = slhfacr.numtarje "
    SQL = SQL & " group by 1,2,3,4,5,6,7,8 "
    SQL = SQL & " order by 1,2,3,4,5,6,7,8 "
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    i = 0
    
    While Not RS.EOF
    
        CodForpa = DevuelveValor("select codforpa from schfacr where " & cadWhere)
    
    
        textcsb33 = "FACT: " & DBLet(RS!letraser, "T") & "-" & Format(DBLet(RS!numfactu, "N"), "0000000") & " " & Format(DBLet(RS!fecfactu, "F"), "dd/mm/yy")
        textcsb33 = textcsb33 & " de " & DBSet(RS!TotalFac, "N")
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
                cadWHERE2 = Replace(cadWhere, "schfac", "slhfac")
            Case "schfacr"
                cadWHERE2 = Replace(cadWhere, "schfacr", "slhfacr")
            Case "schfac1"
                cadWHERE2 = Replace(cadWhere, "schfac1", "slhfac1")
        End Select


        
        '++[Monica]05/08/2011: se añaden tantos vencimientos como nos indique la forma de pago
        
        'Obtener el Nº de Vencimientos de la forma de pago
        SQL = "SELECT numerove, diasvto primerve, restoven FROM sforpa WHERE codforpa=" & CodForpa
        Set rsVenci = New ADODB.Recordset
        rsVenci.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not rsVenci.EOF Then
            If rsVenci!numerove > 0 And DBLet(RS!TotalFac) <> 0 Then
        
        
                i = i + 1
        
                '++[Monica]05/08/2011: si no hay fecha de vencimiento ponemos la fecha de factura, si no los calculos se hacen con la
                '                    fecha de vencimiento
                If FechaVen = "" Then
                    FechaVen = DBLet(RS!fecfactu, "F")
                    FechaVen = DateAdd("d", DBLet(rsVenci!primerve, "N"), FechaVen)
                End If
                
                FecVenci = CDate(FechaVen)
                '++fin
        
                '-------- Primer Vencimiento
                'FECHA VTO
                'FecVenci = CDate(FecVenci)
                'FecVenci = DateAdd("d", DBLet(RsVenci!primerve, "N"), FechaVen)
                '===
        
                '[Monica]17/01/2013: Calculamos la nueva fecha de vencimiento si el cliente tiene dia fijo de pago
                If vsocio.DiaPago <> "" Then
                    FecVenci = NuevaFechaVto(FecVenci, vsocio.DiaPago)
                End If
                
                '[Monica]28/12/2015: quitamos lo que hay en el vale
                Dim SqlNuevo As String
                Dim ImporteVale As Currency
'
'                SqlNuevo = "select sum(coalesce(importevale,0)) from "
'                Select Case vTabla
'                    Case "schfac"
'                        SqlNuevo = SqlNuevo & " slhfac where " & cadWHERE2
'                    Case "schfacr"
'                        SqlNuevo = SqlNuevo & " slhfacr where " & cadWHERE2
'                    Case "schfac1"
'                        SqlNuevo = SqlNuevo & " slhfac1 where " & cadWHERE2
'                End Select
'                ImporteVale = DevuelveValor(SqlNuevo)

                ' el importevale ya no lo llevamos aqui pq no jugamos con el totalfac de schfacr sino con la suma de lineas de la misma ccc
                ' se pone en
               ImporteVale = 0
               'IMPORTE del Vencimiento
                                                        '[Monica]28/12/2015: le quitamos el importe del vale que va en un registro aparte
                TotalFactura2 = DBLet(RS!TotalFac, "N") - ImporteVale
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFactura2
                Else
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFactura2 Then
                        ImpVenci = Round(ImpVenci + (TotalFactura2 - ImpVenci * rsVenci!numerove), 2)
                    End If
                End If

        
                CadValuesAux2 = "(" & DBSet(RS!letraser, "T") & ", " & DBSet(RS!numfactu, "N") & ", " & DBSet(RS!fecfactu, "F") & ", "
                      
                CadValues2 = CadValuesAux2 & DBSet(i, "N") & "," & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(CodForpa, "N") & "," & Format(DBSet(FecVenci, "F"), FormatoFecha) & ","
              

                CodmacBPr = ""
                CodmacBPr = DevuelveDesdeBD("codmacta", "sbanco", "codbanpr", CStr(Banpr), "N")
                
                
                ' en lugar de por socio o por tarjeta nos fijamos si tiene o no ccc
                If DBLet(RS!codbanco, "N") = 0 Or DBLet(RS!codsucur, "N") = 0 Or DBLet(RS!digcontr, "N") = 0 Or DBLet(RS!cuentaba, "N") = 0 Then
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                        
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                    
                    Else
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        End If
                    End If
                Else
                    If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(DBLet(RS!IBAN), "") & MiFormat(DBLet(RS!codbanco), "0000") & MiFormat(DBLet(RS!codsucur), "0000") & MiFormat(DBLet(RS!DigControl), "00") & MiFormat(DBLet(RS!cuentaba), "0000000000")
                        
                        CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                    
                    Else
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(RS!codbanco, "N") & ", " & DBSet(RS!codsucur, "N") & ", " & DBSet(RS!digcontr, "T") & ", " & DBSet(RS!cuentaba, "T") & ", " & DBSet(RS!IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        Else
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(RS!codbanco, "N") & ", " & DBSet(RS!codsucur, "N") & ", " & DBSet(RS!digcontr, "T") & ", " & DBSet(RS!cuentaba, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        End If
                    End If
                End If
                
                If vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & ", 1),"
                
                Else
                    
                    CadValues2 = CadValues2 & _
                                 DBSet(textcsb42, "T") & "," & DBSet(textcsb43, "T") & "," & DBSet(textcsb51, "T") & "," & DBSet(textcsb52, "T") & "," & DBSet(textcsb53, "T") & "," & DBSet(textcsb61, "T") & "," & DBSet(textcsb62, "T") & "," & DBSet(textcsb63, "T") & "," & DBSet(textcsb71, "T") & "," & _
                                 DBSet(textcsb72, "T") & "," & DBSet(textcsb73, "T") & "," & DBSet(textcsb81, "T") & "," & DBSet(textcsb82, "T") & "," & DBSet(textcsb83, "T") & ", 1),"
                End If
                
                '[Monica]28/12/2015: dejamos en el ultimo lo correspondiente al importevale
                Dim J As Integer
                             
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For J = 2 To rsVenci!numerove
                    '[Monica]28/12/2015: dejamos en el ultimo lo correspondiente al importevale
                    i = i + 1
                   
                   
                   'FECHA Resto Vencimientos
                    FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                
                    '[Monica]17/01/2013: Calculamos la nueva fecha de vencimiento si el cliente tiene dia fijo de pago
                    If vsocio.DiaPago <> "" Then
                        FecVenci = NuevaFechaVto(FecVenci, vsocio.DiaPago)
                    End If
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                    
                    
                    CadValues2 = CadValues2 & CadValuesAux2 & DBSet(i, "N") & "," & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(CodForpa, "N") & "," & Format(DBSet(FecVenci, "F"), FormatoFecha) & ","
                    
                    

                    ' en lugar de por socio o por tarjeta nos fijamos si tiene o no ccc
                    If DBLet(RS!codbanco, "N") = 0 Or DBLet(RS!codsucur, "N") = 0 Or DBLet(RS!digcontr, "N") = 0 Or DBLet(RS!cuentaba, "N") = 0 Then
                        If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                            
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        
                        Else
                    
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            End If
                        End If
                    Else
                        If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(DBLet(RS!IBAN), "") & MiFormat(DBLet(RS!codbanco), "0000") & MiFormat(DBLet(RS!codsucur), "0000") & MiFormat(DBLet(RS!DigControl), "00") & MiFormat(DBLet(RS!cuentaba), "0000000000")
                            
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        
                        Else
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(RS!codbanco, "N") & ", " & DBSet(RS!codsucur, "N") & ", " & DBSet(RS!digcontr, "T") & ", " & DBSet(RS!cuentaba, "T") & ", " & DBSet(RS!IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(RS!codbanco, "N") & ", " & DBSet(RS!codsucur, "N") & ", " & DBSet(RS!digcontr, "T") & ", " & DBSet(RS!cuentaba, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            End If
                        End If
                    End If

                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & ", 1),"
                    
                    Else
    
                        CadValues2 = CadValues2 & _
                                     DBSet(textcsb42, "T") & "," & DBSet(textcsb43, "T") & "," & DBSet(textcsb51, "T") & "," & DBSet(textcsb52, "T") & "," & DBSet(textcsb53, "T") & "," & DBSet(textcsb61, "T") & "," & DBSet(textcsb62, "T") & "," & DBSet(textcsb63, "T") & "," & DBSet(textcsb71, "T") & "," & _
                                     DBSet(textcsb72, "T") & "," & DBSet(textcsb73, "T") & "," & DBSet(textcsb81, "T") & "," & DBSet(textcsb82, "T") & "," & DBSet(textcsb83, "T") & ", 1),"
                            
                    End If
                
                Next J
                         
                '[Monica]28/12/2015: dejamos en el ultimo lo correspondiente al importevale
                If DBLet(RS!ImporteVale, "N") <> 0 Then
                    i = i + 1
                   
                    FecVenci = CDate(FechaVen)
            
            
                    '[Monica]17/01/2013: Calculamos la nueva fecha de vencimiento si el cliente tiene dia fijo de pago
                    If vsocio.DiaPago <> "" Then
                        FecVenci = NuevaFechaVto(FecVenci, vsocio.DiaPago)
                    End If
                
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = DBLet(RS!ImporteVale, "N")
                    
                    
                    ' la forma de pago es la del vale
                    Dim SqlNuevo1 As String
                    Dim CodForpaVale As Integer
                    CodForpaVale = DevuelveValor("select codforpa from sforpa where tipovale = 1")
                    If vParamAplic.ContabilidadNueva Then
                        SqlNuevo = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", DBLet(CodForpaVale), "N")
                        'si no existe la forma de pago en conta, insertamos la de ariges
                        If SqlNuevo = "" Then
                            'insertamos e sforpa de la CONTA
                            SQL = "INSERT INTO formapago(codforpa,nomforpa,tipforpa,numerove,primerve,restoven) "
                            SQL = SQL & " select codforpa, nomforpa, tipforpa, numerove, diasvto, restoven "
                            SQL = SQL & " from " & vSesion.CadenaConexion & ".sforpa where codforpa = " & DBSet(CodForpaVale, "N")
                            
                            ConnConta.Execute SQL
                        End If
                    
                    Else
                        SqlNuevo = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", DBLet(CodForpaVale), "N")
                        'si no existe la forma de pago en conta, insertamos la de ariges
                        If SqlNuevo = "" Then
                            SqlNuevo1 = "tipforpa"
                            SqlNuevo = DevuelveDesdeBDNew(cPTours, "sforpa", "nomforpa", "codforpa", DBLet(CodForpaVale), "N", SqlNuevo1)
                            'insertamos e sforpa de la CONTA
                            SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                            SQL = SQL & " VALUES(" & DBSet(CodForpaVale, "N") & ", " & DBSet(SqlNuevo, "T") & ", " & SqlNuevo1 & ")"
                            ConnConta.Execute SQL
                        End If
                    End If
                    CadValues2 = CadValues2 & CadValuesAux2 & DBSet(i, "N") & "," & DBSet(vsocio.CuentaConta, "T") & "," & DBSet(CodForpaVale, "N") & "," & Format(DBSet(FecVenci, "F"), FormatoFecha) & ","
                    
                    
                    ' en lugar de por socio o por tarjeta nos fijamos si tiene o no ccc
                    If DBLet(RS!codbanco, "N") = 0 Or DBLet(RS!codsucur, "N") = 0 Or DBLet(RS!digcontr, "N") = 0 Or DBLet(RS!cuentaba, "N") = 0 Then
                        If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(vsocio.IBAN, "") & MiFormat(vsocio.Banco, "0000") & MiFormat(vsocio.Sucursal, "0000") & MiFormat(vsocio.DigControl, "00") & MiFormat(vsocio.CuentaBan, "0000000000")
                            
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        Else
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(vsocio.IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vsocio.Banco, "N") & ", " & DBSet(vsocio.Sucursal, "N") & ", " & DBSet(vsocio.DigControl, "T") & ", " & DBSet(vsocio.CuentaBan, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            End If
                        End If
                    Else
                        If vParamAplic.ContabilidadNueva Then
                            vvIban = MiFormat(DBLet(RS!IBAN), "") & MiFormat(DBLet(RS!codbanco), "0000") & MiFormat(DBLet(RS!codsucur), "0000") & MiFormat(DBLet(RS!DigControl), "00") & MiFormat(DBLet(RS!cuentaba), "0000000000")
                            
                            CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(vvIban, "T", "S") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                        
                        Else
                            If vEmpresa.HayNorma19_34Nueva = 1 Then
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(RS!codbanco, "N") & ", " & DBSet(RS!codsucur, "N") & ", " & DBSet(RS!digcontr, "T") & ", " & DBSet(RS!cuentaba, "T") & ", " & DBSet(RS!IBAN, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            Else
                                CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CodmacBPr, "T") & ", " & DBSet(RS!codbanco, "N") & ", " & DBSet(RS!codsucur, "N") & ", " & DBSet(RS!digcontr, "T") & ", " & DBSet(RS!cuentaba, "T") & ", " & DBSet(textcsb33, "T") & "," & DBSet(textcsb41, "T") & ","
                            End If
                        End If
                    End If

                    If vParamAplic.ContabilidadNueva Then
                        CadValues2 = CadValues2 & ", 1),"
                    
                    Else
                        CadValues2 = CadValues2 & _
                                     DBSet(textcsb42, "T") & "," & DBSet(textcsb43, "T") & "," & DBSet(textcsb51, "T") & "," & DBSet(textcsb52, "T") & "," & DBSet(textcsb53, "T") & "," & DBSet(textcsb61, "T") & "," & DBSet(textcsb62, "T") & "," & DBSet(textcsb63, "T") & "," & DBSet(textcsb71, "T") & "," & _
                                     DBSet(textcsb72, "T") & "," & DBSet(textcsb73, "T") & "," & DBSet(textcsb81, "T") & "," & DBSet(textcsb82, "T") & "," & DBSet(textcsb83, "T") & ", 1),"
                    End If
                End If
                         
                         

                If vsocio.CuentaConta <> "" Then
                    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
                    'forma de pago de la factura. Sino existe insertarla
                    'vemos si existe en la conta
                    If vParamAplic.ContabilidadNueva Then
                        CadValuesAux2 = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", CStr(CodForpa), "N")
                        'si no existe la forma de pago en conta, insertamos la de ariges
                        If CadValuesAux2 = "" Then
                            SQL = "INSERT INTO formapago(codforpa,nomforpa,tipforpa,numerove,primerve,restoven) "
                            SQL = SQL & " select codforpa, nomforpa, tipforpa, numerove, diasvto, restoven "
                            SQL = SQL & " from " & vSesion.CadenaConexion & ".sforpa where codforpa = " & DBSet(CodForpa, "N")
                            ConnConta.Execute SQL
                        End If
                    Else
                        CadValuesAux2 = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", CStr(CodForpa), "N")
                        'si no existe la forma de pago en conta, insertamos la de ariges
                        If CadValuesAux2 = "" Then
                            cadValuesAux = "tipforpa"
                            CadValuesAux2 = DevuelveDesdeBDNew(cPTours, "sforpa", "nomforpa", "codforpa", CStr(CodForpa), "N", cadValuesAux)
                            'insertamos e sforpa de la CONTA
                            SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                            SQL = SQL & " VALUES(" & DBSet(CodForpa, "N") & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
                            ConnConta.Execute SQL
                        End If
                    End If
                    
                    If vParamAplic.ContabilidadNueva Then
                        'Insertamos en la tabla scobro de la CONTA
                        SQL = "INSERT INTO scobro (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1,"
                        SQL = SQL & "iban,text33csb , text41csb,agente) "
                        SQL = SQL & " VALUES " & Mid(CadValues2, 1, Len(CadValues2) - 1)
                    Else
                        'Insertamos en la tabla scobro de la CONTA
                        SQL = "INSERT INTO cobros (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur, digcontr, cuentaba,"
                        '[Monica]22/11/2013: Tema Iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            SQL = SQL & "iban,text33csb , text41csb,"
                        Else
                            SQL = SQL & "text33csb , text41csb,"
                        End If
                        SQL = SQL & "text42csb, text43csb, text51csb, text52csb, text53csb, text61csb, text62csb, text63csb, text71csb, text72csb, text73csb, text81csb, text82csb, text83csb,agente) "
                        SQL = SQL & " VALUES " & Mid(CadValues2, 1, Len(CadValues2) - 1)
                    
                    End If
                        
                    ConnConta.Execute SQL
                End If
            End If
        End If
        
        RS.MoveNext
    Wend
    Set RS = Nothing

    B = True

EInsertarTesoreria:
    If Err.Number <> 0 Then
        B = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaAjenas = B
End Function



