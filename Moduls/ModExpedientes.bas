Attribute VB_Name = "ModExpedientes"
Option Explicit


Public Function SolicitaCrearExpte() As Boolean
Dim cad As String

    cad = "Se va a crear un Expediente para la venta." & vbCrLf
    cad = cad & "¿Desea continuar?"
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then SolicitaCrearExpte = True
    
End Function




'Private Function InsertarExpte(ByRef cVen As CVenta, cPre As CPresupuesto) As Boolean
''Inserta un expediente de Individuales o de Grupos en la tabla "expgrupo" o "expincab"
''El expediente se crea a partir del presupuesto aceptado
'Dim b As Boolean
'Dim numCont As String 'nuevo contador que obtenemos
'Dim sql As String
'
'    On Error GoTo EInsExp
'
'    'Obtener el numero de expediente (contador)
'    '----------------------------------------------
'    'comprobar si se va a insertar en expedientes de individuales o grupos
'    If cVen.ClaseVenta = 1 Then 'individuales
'        b = ObtenerContadorExpInd(cPre.FechaAcept, numCont)
'    ElseIf cVen.ClaseVenta = 2 Then 'grupos
'        'fecha expediente la del campo aceptado
'        b = ObtenerContadorExpGrp(cPre.FechaAcept, numCont)
'    End If
'
'
'
'    'Insertar en la tabla de expediente
'    '----------------------------------------------
'    If b Then
'        If cVen.ClaseVenta = 1 Then 'individuales
'            'insertamos en la tabla "expincab" los datos del presup. aceptado
'            sql = "INSERT INTO expincab (numexped,codempre,codagenc,codemple,numventa,fechaexp,codclien,sitexped,localiza,reserpor,okvended,fecieanu) "
'            sql = sql & " VALUES (" & numCont & "," & cVen.CodEmpresa & "," & cVen.CodAgencia & ","
'            sql = sql & DBSet(vSesion.Empleado, "N", "S") & "," 'empleado conectado
'            sql = sql & cVen.NumVenta & "," & DBSet(cPre.FechaAcept, "F", "N") & "," & cVen.CodCliente & ","
'            'situacion expediente 0=Abierto
'            sql = sql & "0," & DBSet(cVen.Localizador, "T", "S") & "," & DBSet(cVen.Solicita, "T", "S") & ",0,"
'            If cVen.ClaseAgencia = 1 Then 'Minorista
'                sql = sql & DBSet(cPre.FechaReg, "F", "S") & ")"
'            Else 'Mayorista
'                sql = sql & DBSet(cVen.FechaSal, "F", "S") & ")"
'            End If
'            conn.Execute sql
'
'        ElseIf cVen.ClaseVenta = 2 Then 'grupos
'            'insertamos en la tabla "expgrupo" los datos del presup. aceptado
'            sql = "INSERT INTO expgrupo (numexped,numventa,codempre,codagenc,codemple,empleges,codclien,coddesti,codprodu,codfovia,coditine,numerpax,ventapax,fechasal,fechareg,numedias,numnoche,sitadmon,trasantc,trasantp,okventas,okcompra) "
'            sql = sql & "VALUES (" & numCont & "," & cVen.NumVenta & "," & cVen.CodEmpresa & "," & cVen.CodAgencia & "," & cVen.CodEmple & ","
'            'empleado gestion el del empleado conectado cuando se genera el expte
'            sql = sql & DBSet(vSesion.Empleado, "N", "S") & "," & cVen.CodCliente & ","
'            'le ponemos al expediente los datos del presupuesto:
'            sql = sql & cPre.CodDesti & "," & DBSet(cPre.CodProdu, "N", "S") & ", " & DBSet(cPre.codFovia, "N", "S") & ", " & DBSet(cPre.CodItine, "N", "S") & ", "
'            'le ponemos al expediente los datos del presupuesto
'            sql = sql & cVen.NumPlazas
'            'si MINORISTA en ventapax=1 (venta por grupo)
'            'si MAYORISTA en ventapax=2 (plazas sueltas)
'            If cVen.ClaseAgencia = 1 Then
'                sql = sql & ",1,"
'            Else
'                sql = sql & ",2,"
'            End If
'            sql = sql & DBSet(cPre.FechaSal, "F") & "," & DBSet(cPre.FechaReg, "F") & "," & DBSet(cPre.NumeDias, "N") & "," & DBSet(cPre.NumeNoches, "N") & ","
'            'datos admon: expediente abierto=0
'            sql = sql & "0,0,0,0,0)"
'            conn.Execute sql
'        End If
'    End If
'
'
'    'Actualizar en la tabla de ventas(PREVENTA) el expediente asociado a la venta
'    '-----------------------------------------------------------------------------
'    If b Then
'        sql = "UPDATE preventa SET numexped=" & numCont
'        'poner el empleado gestion de la venta el q este conectado cuando se abre expte
'        If cVen.CodEmpleGes = 0 Then sql = sql & ", codemges=" & vSesion.Empleado
'        'Además actualizar el estado si esta en 1=Pendiente y pasarlo a 6=En gestion
'        If cVen.EstadoVenta = 1 Then sql = sql & ", estventa= 6 "
'        sql = sql & " WHERE numventa=" & cVen.NumVenta & " AND codempre=" & cVen.CodEmpresa
'        conn.Execute sql
'    End If
'
'EInsExp:
'    If Err.Number <> 0 Then
'        b = False
'        MuestraError Err.Number, "Insertar Expediente.", Err.Description
'    End If
'    InsertarExpte = b
'End Function
'

'## ANTES
'Public Function CrearExpedienteVenta(ByRef cVen As CVenta, cPre As CPresupuesto) As Boolean
''Crear un expediente asociado a una venta
'Dim strSelVenta As String 'where de seleccion de la venta
'Dim SQL As String
'Dim b As Boolean
'
'    'para ello comprobamos que la venta no tenga ya un expediente asociado
'    'si estamos en grupos
'    strSelVenta = " numventa=" & cVen.NumVenta & " AND codempre=" & cVen.CodEmpresa
'    If cVen.ClaseVenta = 1 Then 'individuales
'        SQL = "SELECT count(*) FROM expincab WHERE " & strSelVenta
'    ElseIf cVen.ClaseVenta = 2 Then 'grupos
'        SQL = "SELECT count(*) FROM expgrupo WHERE " & strSelVenta
'    End If
'
'    If RegistrosAListar(SQL) > 0 Then
'        MsgBox "Ya existe un Expediente para la venta. ", vbExclamation
'        b = True
'
'    Else
'        'abrimos el expediente
'        'Solo se puede abrir un expediente si venta esta: Pendiente/En gestion
'        If cVen.EstadoVenta = 1 Or cVen.EstadoVenta = 6 Then
'            'y el cliente no podra ser el 0
'            If cVen.CodCliente <> 0 Then
''                If Not cPre Is Nothing Then
'                    b = InsertarExpte(cVen, cPre)
''                Else
''                    b = InsertarExpte(cVen)
''                End If
'            Else
'                SQL = "No se puede abrir un expediente al cliente 000000."
'                MsgBox SQL, vbExclamation
'            End If
'        Else
'            SQL = "No se puede abrir un expediente de la venta." & vbCrLf
'            SQL = SQL & "La venta no esta en PENDIENTE ni EN GESTIÓN."
'            MsgBox SQL, vbExclamation
'        End If
'    End If
'    CrearExpedienteVenta = b
'
'End Function

'Public Function CrearExpedienteVenta(ByRef cVen As CVenta, cPre As CPresupuesto) As Boolean
'''Crear un expediente asociado a una venta
'Dim b As Boolean
'
'    If PuedeCrearExpediente(cVen) Then b = InsertarExpte(cVen, cPre)
'
'
'
''Dim strSelVenta As String 'where de seleccion de la venta
''Dim SQL As String
'
''
''    'para ello comprobamos que la venta no tenga ya un expediente asociado
''    'si estamos en grupos
''    strSelVenta = " numventa=" & cVen.NumVenta & " AND codempre=" & cVen.CodEmpresa
''    If cVen.ClaseVenta = 1 Then 'individuales
''        SQL = "SELECT count(*) FROM expincab WHERE " & strSelVenta
''    ElseIf cVen.ClaseVenta = 2 Then 'grupos
''        SQL = "SELECT count(*) FROM expgrupo WHERE " & strSelVenta
''    End If
''
''    If RegistrosAListar(SQL) > 0 Then
''        MsgBox "Ya existe un Expediente para la venta. ", vbExclamation
''        b = True
''
''    Else
''        'abrimos el expediente
''        'Solo se puede abrir un expediente si venta esta: Pendiente/En gestion
''        If cVen.EstadoVenta = 1 Or cVen.EstadoVenta = 6 Then
''            'y el cliente no podra ser el 0
''            If cVen.CodCliente <> 0 Then
'''                If Not cPre Is Nothing Then
''                    b = InsertarExpte(cVen, cPre)
'''                Else
'''                    b = InsertarExpte(cVen)
'''                End If
''            Else
''                SQL = "No se puede abrir un expediente al cliente 000000."
''                MsgBox SQL, vbExclamation
''            End If
''        Else
''            SQL = "No se puede abrir un expediente de la venta." & vbCrLf
''            SQL = SQL & "La venta no esta en PENDIENTE ni EN GESTIÓN."
''            MsgBox SQL, vbExclamation
''        End If
''    End If
'
'    CrearExpedienteVenta = b
'
'End Function



Public Function PuedeCrearExpediente(ByRef cVen As CVenta) As Boolean
'Crear un expediente asociado a una venta
Dim strSelVenta As String 'where de seleccion de la venta
Dim sql As String
Dim b As Boolean

    'para ello comprobamos que la venta no tenga ya un expediente asociado
    'si estamos en grupos
    strSelVenta = " numventa=" & cVen.NumVenta & " AND codempre=" & cVen.CodEmpresa
    If cVen.ClaseVenta = 1 Then 'individuales
        sql = "SELECT count(*) FROM expincab WHERE " & strSelVenta
    ElseIf cVen.ClaseVenta = 2 Then 'grupos
        sql = "SELECT count(*) FROM expgrupo WHERE " & strSelVenta
    End If
    
    If RegistrosAListar(sql) > 0 Then
        MsgBox "Ya existe un Expediente para la venta. ", vbExclamation
        b = False
        
    Else
        'abrimos el expediente
        'Solo se puede abrir un expediente si venta esta: 1=Pendiente/ 6=En gestion
        If cVen.EstadoVenta = 1 Or cVen.EstadoVenta = 6 Then
            'y el cliente no podra ser el 0
            If cVen.CodCliente <> 0 Then
                b = True
'                If Not cPre Is Nothing Then
'                    b = InsertarExpte(cVen, cPre)
'                Else
'                    b = InsertarExpte(cVen)
'                End If
            Else
                b = False
                sql = "No se puede abrir un expediente al cliente 000000."
                MsgBox sql, vbExclamation
            End If
        Else
            b = False
            sql = "No se puede abrir un expediente de la venta." & vbCrLf
            sql = sql & "La venta no esta en PENDIENTE ni EN GESTIÓN."
            MsgBox sql, vbExclamation
        End If
    End If
    PuedeCrearExpediente = b
    
End Function




Public Function ComprobarFechaFolleto(codFoll As String, codEmp As String, fecSal As String) As Boolean
'Comprueba que la fecha de salida del presupuesto/Expediente esta dentro
'del período de validez del folleto.
'(IN): codFoll,codEmp (cod. folleto,cod. empresa son la clave primaria del folleto)
'(IN): fecSal = fecha de Salida
Dim cFoll As CFolleto
    
    On Error GoTo ErrComp
    
    If codFoll <> "" Then 'Hay un folleto seleccionado
        Set cFoll = New CFolleto
        If cFoll.LeerDatos(codFoll, codEmp) Then
            If cFoll.ValidezFolleto(fecSal) Then ComprobarFechaFolleto = True
        End If
        Set cFoll = Nothing
    
    Else
        ComprobarFechaFolleto = True
    End If
    
    Exit Function
    
ErrComp:
    MuestraError Err.Number, "Comprobar Fechas validez del folleto.", Err.Description
End Function

