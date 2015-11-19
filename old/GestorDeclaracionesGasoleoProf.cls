VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GestorDeclaracionesGasoleoProf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private sql As String
Private RS As ADODB.Recordset
'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent procesando[(arg1, arg2, ... , argn)]
Public Event procesando(registro As Integer)


Public Function declaraGasoleoProfesional(Conn As ADODB.Connection, Optional test As Boolean) As Boolean
    Dim mDeclaracion As Declaracion
    Dim mColSuministro As ColSuministro
    Dim mSuministro As Suministro
    Dim i As Integer
On Error GoTo err_declaraGasoleoProfesional
    Conn.Execute "START TRANSACTION"
        sql = "select a.numtarje, a.fecalbar, a.horalbar, a.cantidad, b.matricul, c.nifsocio, d.porcbd" & _
                " from scaalb as a, starje as b, ssocio as c, sartic as d" & _
                " where a.numtarje in (select numtarje from starje where tiptarje = 2)" & _
                " and a.codartic in (select codartic from sartic where gp = 1)" & _
                " and b.numtarje = a.numtarje" & _
                " and c.codsocio = a.codsocio" & _
                " and d.codartic = a.codartic" & _
                " and a.declaradogp = 0"
        Set RS = New ADODB.Recordset
        RS.Open sql, Conn, adOpenForwardOnly
        If Not RS.EOF Then
            i = 0
            Set mDeclaracion = New Declaracion
            RS.MoveFirst
            '-- el id se montará al guardar la Declaracion
            '-- el codEE se obtiene automáticamente
            If test Then mDeclaracion.test = "S" Else mDeclaracion.test = "N"
            Set mColSuministro = New ColSuministro
            While Not RS.EOF
                i = i + 1
                RaiseEvent procesando(i)
                Set mSuministro = New Suministro
                With mSuministro
                    '-- el id se montará al guardar la Declaracion
                    '-- el idmovcont se genera automáticamente
                    .Fecha = RS!fecAlbar
                    .Hora = RS!horalbar
                    .codpro = RS!porcbd
                    .lit = RS!cantidad
                    .NIF = RS!nifsocio
                    .Matricula = RS!matricul
                End With
                mColSuministro.Add mSuministro
                RS.MoveNext
            Wend
            Set mDeclaracion.suministros = mColSuministro
            If Not mDeclaracion.guardar(Conn) Then
                Conn.Execute "ROLLBACK"
                Exit Function
            End If
        End If
        '-- Ya lo hemos pasado a la base de datos y lo actualizamos como declarado
        sql = "update scaalb set declaradogp = 1" & _
                " where numtarje in (select numtarje from starje where tiptarje = 2)" & _
                " and codartic in (select codartic from sartic where gp = 1)"
        Conn.Execute sql
    Conn.Execute "COMMIT"
    declaraGasoleoProfesional = True
    Exit Function
err_declaraGasoleoProfesional:
    Conn.Execute "ROLLBACK"
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, "declaraGasoleoProfesional"
End Function

Public Function enviaDeclaracionesPendientes(Conn As ADODB.Connection) As Boolean
    Dim mDeclaracion As Declaracion
    Dim i As Integer
On Error GoTo err_enviaDeclaracionesPendientes
    Conn.Execute "START TRANSACTION"
        sql = "select * from gp_suministrv2ent where situacion = 0"
        Set RS = New ADODB.Recordset
        RS.Open sql, Conn, adOpenForwardOnly
        If Not RS.EOF Then
            i = 0
            RS.MoveFirst
            While Not RS.EOF
                i = i + 1
                RaiseEvent procesando(i)
                Set mDeclaracion = New Declaracion
                If Not mDeclaracion.Leer(RS!ID, Conn) Then
                    Conn.Execute "ROLLBACK"
                    Exit Function
                End If
                If Not mDeclaracion.exportarXML Then
                    Conn.Execute "ROLLBACK"
                    Exit Function
                End If
                RS.MoveNext
            Wend
        End If
        If Not enviaAEAT() Then
            Conn.Execute "ROLLBACK"
            Exit Function
        End If
        sql = "update gp_suministrv2ent set situacion = 1 where situacion = 0"
        Conn.Execute sql
        enviaDeclaracionesPendientes = True
    Conn.Execute "COMMIT"
    arrancaUtilidad
    Exit Function
err_enviaDeclaracionesPendientes:
    Conn.Execute "ROLLBACK"
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, "enviaDeclaracionesPendientes"
End Function

Private Function enviaAEAT() As Boolean
    Dim ms As String
    Dim fichero As String
On Error GoTo err_enviaAEAT
    '-- Envia los fichero XML a la Agencia tributaria
    '-- 1* Comprueba que el directorio con la utilidad de la agencia existe y la utilidad se encuentra en el
    If Dir("C:\aeat\aduanetxml_3_0.jar") = "" Then
        ms = "La utilidad de la agencia tributaria (aduanetxml_3_0.jar) no se encuentra en el el directorio C:\aeat del equipo" & vbCrLf & _
                "Descárgela de la web de la agencia (http://www.agenciatributaria.es) y cópiela a " & App.path & "\aeat"
        MsgBox ms, vbExclamation
        Exit Function
    End If
    '-- 2* Comprobamos que la bandeja de salida está disponible
    If Dir("C:\aeat\aduanetxml\BandejaSalida", vbDirectory) = "" Then ' [4.0.13] Ahora Aeat está en un directorio que cuelga de C:
        ms = "No se encuentra la bandeja de salida en la que depositar los mensajes." & vbCrLf & _
                "Ejecute C:\aeat\aduanetxml_3_0.jar y los directorios se crearán automáticamente."
        MsgBox ms, vbExclamation
        Exit Function
    End If
    '-- 3* Copiar los ficheros a la bandeja de salida y borrarlos de origen
    fichero = Dir(App.path & "\*.xml")
    Do While fichero <> ""
        FileCopy App.path & "\" & fichero, "C:\aeat\aduanetxml\BandejaSalida\" & fichero ' [4.0.13] Ahora Aeat está en un directorio que cuelga de C:
        fichero = Dir
    Loop
    If Dir(App.path & "\*.xml") <> "" Then Kill App.path & "\*.xml"
    '-- 4 * Ya está, ahora si quiere puede arrancar la utilidad
    enviaAEAT = True
    Exit Function
err_enviaAEAT:
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, "enviaDeclaracionesPendientes"
End Function

Private Sub arrancaUtilidad()
    Dim ms As String
    On Error GoTo err_arrancaUtilidad
    '-- Arranca la utilidad de la agencia tributaria
    ms = "El proceso de traspaso de declaraciones ha finalizado corretamente." & _
            "¿Desea arrancar la utilidad de traspaso de la Agencia Tributaria?"
    If MsgBox(ms, vbYesNo + vbQuestion) = vbYes Then
        Shell "javaw -jar C:\aeat\aduanetxml_3_0.jar" ' [4.0.13] Ahora Aeat está en un directorio que cuelga de C:
    End If
    Exit Sub
err_arrancaUtilidad:
    MsgBox "No es posible arrancar automáticamente la utilidad. Hágalo manualmente", vbExclamation
End Sub

Public Function quedaPorDeclarar(Conn As ADODB.Connection, Fecha As Date) As Boolean
    '-- Comprueba si por debajo de una fecha determinada quedan operaciones por declarar.
On Error GoTo err_quedaPorDeclarar
    sql = "select a.numtarje, a.fecalbar, a.horalbar, a.cantidad, b.matricul, c.nifsocio, d.porcbd" & _
            " from scaalb as a, starje as b, ssocio as c, sartic as d" & _
            " where a.numtarje in (select numtarje from starje where tiptarje = 2)" & _
            " and a.codartic in (select codartic from sartic where gp = 1)" & _
            " and b.numtarje = a.numtarje" & _
            " and c.codsocio = a.codsocio" & _
            " and d.codartic = a.codartic" & _
            " and a.declaradogp = 0" & _
            " and a.fecalbar <= '" & Format(Fecha, "yyyy-mm-dd") & "'"
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not RS.EOF Then quedaPorDeclarar = True
    Exit Function
err_quedaPorDeclarar:
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, "enviaDeclaracionesPendientes"
End Function

Public Function recuperacionErrones(Conn As ADODB.Connection, directorio As String) As Boolean
    '-- Esta función se encarga de recuperar las respuestas del directorio que se indica y
    '   prepara unas declaraciones complementarias con el fin de recuperar las erróneas.
    Dim lector As String
    Dim fichero As String
    Dim respuesta As String
    Dim c As String
    Dim nf As Integer
    Dim pos1 As Long
    Dim pos2 As Long
    Dim i1 As Long
    Dim i2 As Long
    Dim mDeclaracion As Declaracion
    Dim mSuministro As Suministro
    '--
    Dim CodRet As String
    Dim IdMov As String
    Dim ID As String
    '--
    Dim sql As String
    Dim RS As ADODB.Recordset
    '-- 1 Leer los ficheros del directorio indicado
    fichero = Dir(directorio & "\*.xml")
    Do While fichero <> ""
        nf = FreeFile()
        lector = ""
        c = ""
        ID = Left(fichero, 14)
        Open directorio & "\" & fichero For Input As nf
        '-- Leemos en un solo paso todo el contenido del fichero
        Do While Not EOF(nf)
            Input #nf, c
            lector = lector & c
        Loop
        '--- 2 Buscar si es erróneo, en las respuestas
        pos1 = InStr(1, lector, "<Respuestas>")
        While pos1 <> 0
            pos2 = InStr(pos1, lector, "</Respuestas")
            '-- Tomamos la resouesta completa y sacamos sus valores
            IdMov = "": CodRet = ""
            i1 = pos1 + 12
            i2 = pos2 - i1
            respuesta = Mid(lector, i1, i2)
            i1 = InStr(1, respuesta, "<IdMovCont>")
            If i1 <> 0 Then
                i2 = InStr(i1 + 1, respuesta, "</IdMovCont>")
                i1 = i1 + 11
                i2 = i2 - i1
                IdMov = Mid(respuesta, i1, i2)
            End If
            i1 = InStr(1, respuesta, "<CodRet>")
            If i1 <> 0 Then
                i2 = InStr(i1 + 1, respuesta, "</CodRet>")
                i1 = i1 + 8
                i2 = i2 - i1
                CodRet = Mid(respuesta, i1, i2)
            End If
            '---- 3 Por cada erróneo generar el correcto.
            If CodRet = "7005" Then
                '-- Es erróneo y hay que volver a declararlo
                If mDeclaracion Is Nothing Then
                    Set mDeclaracion = New Declaracion
                    Set mDeclaracion.suministros = New ColSuministro
                    mDeclaracion.ID = ID & "000002"
                    mDeclaracion.test = "N"
                    mDeclaracion.codee = obtenerCEE(Conn)
                    mDeclaracion.Situacion = 0
                End If
                sql = "select * from gp_suministro where id = '" & ID & "000001'" & _
                        " and idmovcont = " & IdMov
                Set RS = New ADODB.Recordset
                RS.Open sql, Conn, adOpenForwardOnly
                If Not RS.EOF Then
                    Set mSuministro = New Suministro
                    With mSuministro
                        .ID = ID & "000002"
                        .idmovcont = Val(IdMov)
                        .Cim = RS!Cim
                        .Fecha = Format(RS!FechaHora, "dd/mm/yyyy")
                        .Hora = Format(RS!FechaHora, "hh:mm:ss")
                        .codpro = RS!codprod
                        .lit = RS!lit
                        .NIF = RS!NIF
                        .Matricula = RS!Matricula
                    End With
                    mDeclaracion.suministros.Add mSuministro
                End If
                
            End If
            pos1 = InStr(pos1 + 1, lector, "<Respuestas>")
        Wend
        If Not (mDeclaracion Is Nothing) Then
            '-- Se ha encontrado errores y hay algo que declarar.
            mDeclaracion.exportarXML
            mDeclaracion.guardar2 Conn ' graba en otro sitio
            Set mDeclaracion = Nothing
        End If
        fichero = Dir
    Loop
    enviaAEATRecuperada
End Function

Private Function obtenerCEE(Conn As ADODB.Connection) As String
On Error GoTo err_obtenerCEE
    sql = "select cee from sparam"
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not RS.EOF Then
        obtenerCEE = RS!Cee
    Else
        obtenerCEE = ""
    End If
    Exit Function
err_obtenerCEE:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.obtenerCEE"
End Function

Private Function enviaAEATRecuperada() As Boolean
    Dim ms As String
    Dim fichero As String
On Error GoTo err_enviaAEATRecuperada
    '-- 3* Copiar los ficheros a la bandeja de recuperadas y borrarlos de origen
    fichero = Dir(App.path & "\*.xml")
    Do While fichero <> ""
        FileCopy App.path & "\" & fichero, "C:\aeat\RespuestasRecuperadas\" & fichero ' [4.0.13] Ahora Aeat está en un directorio que cuelga de C:
        fichero = Dir
    Loop
    If Dir(App.path & "\*.xml") <> "" Then Kill App.path & "\*.xml"
    '-- 4 * Ya está, ahora si quiere puede arrancar la utilidad
    enviaAEATRecuperada = True
    Exit Function
err_enviaAEATRecuperada:
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, "enviaDeclaracionesPendientes"
End Function

Public Function detectaDuplicidades(Fecha As Date) As Boolean
'-- RV65 (Rafa en alzicoop)
On Error GoTo errDuplicidad
    sql = "select nif, matricula, fechahora, lit, count(*) " & _
            "from gp_suministro " & _
            "where fechahora >= '" & Format(Fecha, "yyyy-mm-dd") & "' " & _
            "and (nif, matricula, fechahora) not in (select nif, matricula, fechahora from gp_suministro2) " & _
            "group by 1,2,3,4 " & _
            "having count(*) > 1"
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not RS.EOF Then detectaDuplicidades = True
    Exit Function
errDuplicidad:
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, "detectaDuplicidades"
End Function

Public Function corrigeDuplicidades(Fecha As Date) As Boolean
'-- RV65 (rafa en alzicoop)
    Dim rs2 As ADODB.Recordset
    Dim cont As Integer
    Dim mDeclaracion As Declaracion
    Dim mSuministro As Suministro
    Dim mColSuministro As ColSuministro
    Dim i As Integer
'On Error GoTo errCorreccion (ojo)
    Conn.Execute "START TRANSACTION"
    sql = "select nif, matricula, fechahora, lit, count(*) " & _
            "from gp_suministro " & _
            "where fechahora >= '" & Format(Fecha, "yyyy-mm-dd") & "' " & _
            "and (nif, matricula, fechahora) not in (select nif, matricula, fechahora from gp_suministro2) " & _
            "group by 1,2,3,4 " & _
            "having count(*) > 1"
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not RS.EOF Then
        i = 0
        '-- Vamos a montar una declaración única con todas las duplicidades
        Set mDeclaracion = New Declaracion
        '-- el ID y el codEE se montan automáricamente al guardar
        mDeclaracion.test = "N" ' no es de test vamos a declarar de verdad
        Set mColSuministro = New ColSuministro
        '-- Miramos una por una las transacciones duplicadas.
        RS.MoveFirst
        While Not RS.EOF
            i = i + 1
            RaiseEvent procesando(i)
            sql = "select * from gp_suministro" & _
                    " where nif = '" & RS!NIF & "'" & _
                    " and matricula = '" & RS!Matricula & "'" & _
                    " and fechahora = '" & Format(CDate(RS!FechaHora), "yyyy-mm-dd hh:nn:ss") & "'" & _
                    " and lit = " & TransformaComasPuntos(CStr(RS!lit))
            cont = 0
            Set rs2 = New ADODB.Recordset
            rs2.Open sql, Conn, adOpenForwardOnly
            If Not rs2.EOF Then
                rs2.MoveFirst
                While Not rs2.EOF
                    cont = cont + 1
                    If cont > 1 Then ' solo duplicadas
                        Set mSuministro = New Suministro
                        With mSuministro
                            '-- el id se montará al guardar la Declaracion
                            '-- el idmovcont se genera automáticamente
                            .Fecha = CDate(Format(rs2!FechaHora, "dd/mm/yyyy"))
                            .Hora = CDate(Format(rs2!FechaHora, "hh:mm:ss"))
                            .codpro = rs2!codprod
                            .lit = -1 * rs2!lit ' este es el misterio, declaramos en negativo
                            .NIF = rs2!NIF
                            .Matricula = rs2!Matricula
                        End With
                        mColSuministro.Add mSuministro
                    End If
                    rs2.MoveNext
                Wend
            End If
            RS.MoveNext
        Wend
        Set mDeclaracion.suministros = mColSuministro
        If mDeclaracion.guardar3(Conn) Then
            Conn.Execute "COMMIT"
            mDeclaracion.exportarXML
            corrigeDuplicidades = True
        Else
            Conn.Execute "ROLLBACK"
        End If
    End If
    Exit Function
errCorreccion:
    Conn.Execute "ROLLBACK"
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, "detectaDuplicidades"
End Function

