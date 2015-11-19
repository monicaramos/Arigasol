VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Suministro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Identifica cada uno de los suministros dentro de una declaración"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarid As String 'copia local
Private mvaridmovcont As Integer 'copia local
Private mvarfecha As Date 'copia local
Private mvarhora As Date 'copia local
Private mvarcodpro As Byte 'copia local
Private mvarlit As Single 'copia local
Private mvarnif As String 'copia local
Private mvarmatricula As String 'copia local
Private mvarCim As String 'copia local

Private sql As String
Private RS As ADODB.Recordset

Public Property Let Cim(ByVal vData As String)
Attribute Cim.VB_Description = "código de la gasolinea"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.cim = 5
    mvarCim = vData
End Property

Public Property Get Cim() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.cim
    Cim = mvarCim
End Property

Public Property Let Matricula(ByVal vData As String)
Attribute Matricula.VB_Description = "Matricula del vehículo al que se le proporcionó el suministro"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.matricula = 5
    mvarmatricula = vData
End Property

Public Property Get Matricula() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.matricula
    Matricula = mvarmatricula
End Property

Public Property Let NIF(ByVal vData As String)
Attribute NIF.VB_Description = "Nif del cliente al que se le ha hecho el suministro"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.nif = 5
    mvarnif = vData
End Property

Public Property Get NIF() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.nif
    NIF = mvarnif
End Property

Public Property Let lit(ByVal vData As Single)
Attribute lit.VB_Description = "Listros de combustible del movimiento 5E2D (-99999.99 / 99999.99)"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.lit = 5
    mvarlit = vData
End Property

Public Property Get lit() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.lit
    lit = mvarlit
End Property

Public Property Let codpro(ByVal vData As Byte)
Attribute codpro.VB_Description = "Porcentaje de biodiesel (0-100)"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codpro = 5
    mvarcodpro = vData
End Property

Public Property Get codpro() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codpro
    codpro = mvarcodpro
End Property

Public Property Let Hora(ByVal vData As Date)
Attribute Hora.VB_Description = "Hora del movimiento"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.hora = 5
    mvarhora = vData
End Property

Public Property Get Hora() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.hora
    Hora = mvarhora
End Property

Public Property Let Fecha(ByVal vData As Date)
Attribute Fecha.VB_Description = "Fecha del movimiento"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecha = 5
    mvarfecha = vData
End Property

Public Property Get Fecha() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecha
    Fecha = mvarfecha
End Property

Public Property Let idmovcont(ByVal vData As Integer)
Attribute idmovcont.VB_Description = "Identificador del movimiento dentro de la transacción"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.idmovcont = 5
    mvaridmovcont = vData
End Property

Public Property Get idmovcont() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.idmovcont
    idmovcont = mvaridmovcont
End Property

Public Property Let ID(ByVal vData As String)
Attribute ID.VB_Description = "Identificador de la transacción"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.id = 5
    mvarid = vData
End Property

Public Property Get ID() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.id
    ID = mvarid
End Property

Public Function Leer(ID As String, idmovcont As Integer, Conn As ADODB.Connection) As Boolean
On Error GoTo err_leer
    sql = "select * from gp_suministro where id = '" & ID & "'" & _
            " and idmovcont = " & CStr(idmovcont)
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not RS.EOF Then
        mvarid = RS!ID
        mvaridmovcont = RS!idmovcont
        mvarCim = RS!Cim
        mvarfecha = CDate(Format(RS!FechaHora, "dd/mm/yyyy"))
        mvarhora = CDate(Format(RS!FechaHora, "00/00/0000 hh:mm:ss"))
        mvarcodpro = RS!codpro
        mvarlit = RS!lit
        mvarnif = RS!NIF
        mvarmatricula = RS!Matricula
        Leer = True
    Else
        Leer = False
    End If
    RS.Close
    Exit Function
err_leer:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Suministro.leer"
End Function

Public Function guardar(Conn As ADODB.Connection, Optional deColeccion As Boolean) As Boolean
On Error GoTo err_guardar
    '-- Comprobación de que disponemos de los valores de clave primaria
    If mvarid = "" Then Exit Function
    If Not deColeccion And mvaridmovcont = 0 Then Exit Function
    '-- obtener el CIM
    If Not obtenerCim(Conn) Then
        MsgBox "No puedo obtener el CIM. Revise parámetros", vbExclamation
        Exit Function
    End If
    '-- Buscamos si el objeto ya existe en la bvase de datos
    sql = "select * from gp_suministro where id = '" & ID & "'" & _
            " and idmovcont = " & CStr(idmovcont)
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not RS.EOF Then
        '-- Ya existe y es una modificación
        sql = "update gp_suministro set" & _
                " cim = '" & mvarCim & "'," & _
                " fecha = '" & Format(mvarfecha, "yyyy-mm-dd") & " " & Format(mvarhora, "hh:mm:ss") & "'," & _
                " codprod = " & CStr(mvarcodpro) & "," & _
                " lit = " & TransformaComasPuntos(Format(mvarlit, "####0.00")) & "," & _
                " nif = '" & mvarnif & "'," & _
                " matricula = '" & mvarmatricula & "'" & _
                " where id = '" & ID & "'" & _
                " and idmovcont = " & CStr(idmovcont)
    Else
        '-- Debemos obtener un contador
        mvaridmovcont = siguienteIdmovcont(Conn)
        '-- No existe y es un alta
        sql = "insert into gp_suministro (id,idmovcont,cim,fechahora,codprod,lit,nif,matricula)" & _
                " values (" & _
                "'" & mvarid & "'," & _
                CStr(mvaridmovcont) & "," & _
                "'" & mvarCim & "'," & _
                "'" & Format(mvarfecha, "yyyy-mm-dd") & " " & Format(mvarhora, "hh:mm:ss") & "'," & _
                CStr(mvarcodpro) & "," & _
                TransformaComasPuntos(Format(mvarlit, "####0.00")) & "," & _
                "'" & mvarnif & "'," & _
                "'" & mvarmatricula & "'" & _
                ")"
    End If
    Conn.Execute sql
    RS.Close
    guardar = True
    Exit Function
err_guardar:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Suministro.guardar"
End Function

Public Function eliminar(Conn As ADODB.Connection) As Boolean
On Error GoTo err_eliminar
    sql = "delete from gp_suministro where id = '" & ID & "'" & _
            " and idmovcont = " & CStr(idmovcont)
    Conn.Execute sql
    eliminar = True
    Exit Function
err_eliminar:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Suministro.eliminar"
End Function


Private Function obtenerCim(Conn As ADODB.Connection) As Boolean
On Error GoTo err_obtenerCIM
    sql = "select cim from sparam"
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not RS.EOF Then
        mvarCim = RS!Cim
        obtenerCim = True
    Else
        obtenerCim = False
    End If
    Exit Function
err_obtenerCIM:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.obtenerCIM"
End Function

Private Function siguienteIdmovcont(Conn As ADODB.Connection) As Integer
    sql = "select max(idmovcont) from gp_suministro where id = '" & mvarid & "'"
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not IsNull(RS.Fields(0)) Then
        siguienteIdmovcont = RS.Fields(0) + 1
    Else
        siguienteIdmovcont = 1
    End If
End Function

Public Function guardar2(Conn As ADODB.Connection, Optional deColeccion As Boolean) As Boolean
On Error GoTo err_guardar2
    '-- Comprobación de que disponemos de los valores de clave primaria
    If mvarid = "" Then Exit Function
    If Not deColeccion And mvaridmovcont = 0 Then Exit Function
    '-- obtener el CIM
    If Not obtenerCim(Conn) Then
        MsgBox "No puedo obtener el CIM. Revise parámetros", vbExclamation
        Exit Function
    End If
    '-- Buscamos si el objeto ya existe en la bvase de datos
    sql = "select * from gp_suministro where id = '" & ID & "'" & _
            " and idmovcont = " & CStr(idmovcont)
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not RS.EOF Then
        '-- Ya existe y es una modificación
        sql = "update gp_suministro2 set" & _
                " cim = '" & mvarCim & "'," & _
                " fecha = '" & Format(mvarfecha, "yyyy-mm-dd") & " " & Format(mvarhora, "hh:mm:ss") & "'," & _
                " codprod = " & CStr(mvarcodpro) & "," & _
                " lit = " & TransformaComasPuntos(Format(mvarlit, "####0.00")) & "," & _
                " nif = '" & mvarnif & "'," & _
                " matricula = '" & mvarmatricula & "'" & _
                " where id = '" & ID & "'" & _
                " and idmovcont = " & CStr(idmovcont)
    Else
        '-- No existe y es un alta
        sql = "insert into gp_suministro2 (id,idmovcont,cim,fechahora,codprod,lit,nif,matricula)" & _
                " values (" & _
                "'" & mvarid & "'," & _
                CStr(mvaridmovcont) & "," & _
                "'" & mvarCim & "'," & _
                "'" & Format(mvarfecha, "yyyy-mm-dd") & " " & Format(mvarhora, "hh:mm:ss") & "'," & _
                CStr(mvarcodpro) & "," & _
                TransformaComasPuntos(Format(mvarlit, "####0.00")) & "," & _
                "'" & mvarnif & "'," & _
                "'" & mvarmatricula & "'" & _
                ")"
    End If
    Conn.Execute sql
    RS.Close
    guardar2 = True
    Exit Function
err_guardar2:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Suministro.guardar2"
End Function

Public Function guardar3(Conn As ADODB.Connection, Optional deColeccion As Boolean) As Boolean
On Error GoTo err_guardar3
    '-- Comprobación de que disponemos de los valores de clave primaria
    If mvarid = "" Then Exit Function
    If Not deColeccion And mvaridmovcont = 0 Then Exit Function
    '-- obtener el CIM
    If Not obtenerCim(Conn) Then
        MsgBox "No puedo obtener el CIM. Revise parámetros", vbExclamation
        Exit Function
    End If
    '-- Buscamos si el objeto ya existe en la bvase de datos
    sql = "select * from gp_suministro2 where id = '" & ID & "'" & _
            " and idmovcont = " & CStr(idmovcont)
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not RS.EOF Then
        '-- Ya existe y es una modificación
        sql = "update gp_suministro2 set" & _
                " cim = '" & mvarCim & "'," & _
                " fecha = '" & Format(mvarfecha, "yyyy-mm-dd") & " " & Format(mvarhora, "hh:mm:ss") & "'," & _
                " codprod = " & CStr(mvarcodpro) & "," & _
                " lit = " & TransformaComasPuntos(Format(mvarlit, "####0.00")) & "," & _
                " nif = '" & mvarnif & "'," & _
                " matricula = '" & mvarmatricula & "'" & _
                " where id = '" & ID & "'" & _
                " and idmovcont = " & CStr(idmovcont)
    Else
        '-- No existe y es un alta
        mvaridmovcont = siguienteIdmovcont3(Conn)
        sql = "insert into gp_suministro2 (id,idmovcont,cim,fechahora,codprod,lit,nif,matricula)" & _
                " values (" & _
                "'" & mvarid & "'," & _
                CStr(mvaridmovcont) & "," & _
                "'" & mvarCim & "'," & _
                "'" & Format(mvarfecha, "yyyy-mm-dd") & " " & Format(mvarhora, "hh:mm:ss") & "'," & _
                CStr(mvarcodpro) & "," & _
                TransformaComasPuntos(Format(mvarlit, "####0.00")) & "," & _
                "'" & mvarnif & "'," & _
                "'" & mvarmatricula & "'" & _
                ")"
    End If
    Conn.Execute sql
    RS.Close
    guardar3 = True
    Exit Function
err_guardar3:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Suministro.guardar3"
End Function

Private Function siguienteIdmovcont3(Conn As ADODB.Connection) As Integer
    sql = "select max(idmovcont) from gp_suministro2 where id = '" & mvarid & "'"
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not IsNull(RS.Fields(0)) Then
        siguienteIdmovcont3 = RS.Fields(0) + 1
    Else
        siguienteIdmovcont3 = 1
    End If
End Function

