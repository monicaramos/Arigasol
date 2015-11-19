VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Declaracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Incluye una declaración completa a la agencia tributaria de diferentes suministros"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private mvarid As String 'copia local
Private mvarcodee As String 'copia local
Private mvartest As Variant 'copia local
Private mvarsituacion As Byte 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarsuministros As ColSuministro 'copia local
Private sql As String
Private RS As ADODB.Recordset


Public Property Set suministros(ByVal vData As ColSuministro)
Attribute suministros.VB_Description = "Suministros que componen una declaración"
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.suministros = Form1
    Set mvarsuministros = vData
End Property


Public Property Get suministros() As ColSuministro
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.suministros
    Set suministros = mvarsuministros
End Property

Public Property Let Situacion(ByVal vData As Byte)
Attribute Situacion.VB_Description = "Posibles valores: 0 = Transacción creada pero no enviada, 1 = Transacción enviada a la agencia tributaria."
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.situacion = 5
    mvarsituacion = vData
End Property

Public Property Get Situacion() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.situacion
    Situacion = mvarsituacion
End Property

Public Property Let test(ByVal vData As Variant)
Attribute test.VB_Description = "Si 'S' indica que la transacción es tipo test, es decir, para probar simplemente"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.test = 5
    mvartest = vData
End Property

Public Property Set test(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.test = Form1
    Set mvartest = vData
End Property

Public Property Get test() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.test
    If IsObject(mvartest) Then
        Set test = mvartest
    Else
        test = mvartest
    End If
End Property

Public Property Let codee(ByVal vData As String)
Attribute codee.VB_Description = "Identificador de la entidad emisora de la tarjeta"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codee = 5
    mvarcodee = vData
End Property

Public Property Get codee() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codee
    codee = mvarcodee
End Property

Public Property Let ID(ByVal vData As String)
Attribute ID.VB_Description = "Identificador de transaccion en la forma AAAAMMDDHHMMSSNNNNNN"
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.id = 5
    mvarid = vData
End Property

Public Property Get ID() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.id
    ID = mvarid
End Property

Public Function Leer(ID As String, Conn As ADODB.Connection) As Boolean
    Dim vSuministro As Suministro
On Error GoTo err_leer
    sql = "select a.id, a.codee, a.test," & _
            " b.idmovcont, b.cim, b.fechahora, b.codprod, b.lit, b.nif, b.matricula" & _
            " from gp_suministrv2ent as a, gp_suministro as b" & _
            " where a.id = '" & ID & "'" & _
            " and b.id = a.id"
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not RS.EOF Then
        RS.MoveFirst
        '-- cargamos los atributos principales
        mvarid = RS!ID
        mvarcodee = RS!codee
        mvartest = RS!test
        Set mvarsuministros = New ColSuministro
        While Not RS.EOF
            '-- Cargamos la colección de suministros
            Set vSuministro = New Suministro
            With vSuministro
                .ID = RS!ID
                .idmovcont = RS!idmovcont
                .Cim = RS!Cim
                .Fecha = CDate(Format(RS!FechaHora, "dd/mm/yyyy"))
                .Hora = CDate(Format(RS!FechaHora, "hh:mm:ss"))
                .codpro = RS!codprod
                .lit = RS!lit
                .NIF = RS!NIF
                .Matricula = RS!Matricula
            End With
            mvarsuministros.Add vSuministro, CStr(vSuministro.idmovcont)
            RS.MoveNext
        Wend
        Leer = True
    Else
        Leer = False
    End If
    Exit Function
err_leer:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.leer"
End Function

Public Function guardar(Conn As ADODB.Connection) As Boolean
    Dim vSuministro As Suministro
On Error GoTo err_guardar
    '-- Generamos el valor de ID
    mvarid = Format(Now, "yyyymmddhhnnss000001")
    '-- Obtenemos el CEE
    If Not obtenerCEE(Conn) Then
        MsgBox "No se puede obtener el CEE. Revise parámetros", vbExclamation
        Exit Function
    End If
    sql = "delete from gp_suministrv2ent where id = '" & mvarid & "'"
    Conn.Execute sql
    sql = "insert into gp_suministrv2ent (id,codee,test,situacion)" & _
            " values (" & _
            "'" & mvarid & "'," & _
            "'" & mvarcodee & "'," & _
            "'" & mvartest & "'," & _
            "0" & _
            ")"
    Conn.Execute sql ' ya tenemos la cabecera
    For Each vSuministro In mvarsuministros
        '-- Confirmamos la clave principal
        vSuministro.ID = mvarid
        If Not vSuministro.guardar(Conn, True) Then
            Exit Function
        End If
    Next
    guardar = True
    Exit Function
err_guardar:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.guardar"
End Function

Public Function eliminar(Conn As ADODB.Connection) As Boolean
On Error GoTo err_eliminar
    sql = "delete from gp_suministro where id = '" & mvarid & "'"
    Conn.Execute sql
    sql = "delete from gp_suministrv2ent where id = '" & mvarid & "'"
    Conn.Execute sql
    eliminar = True
    Exit Function
err_eliminar:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.eliminar"
End Function

Private Function obtenerCEE(Conn As ADODB.Connection) As Boolean
On Error GoTo err_obtenerCEE
    sql = "select cee from sparam"
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly
    If Not RS.EOF Then
        mvarcodee = RS!Cee
        obtenerCEE = True
    Else
        obtenerCEE = False
    End If
    Exit Function
err_obtenerCEE:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.obtenerCEE"
End Function

'Public Function exportarXML() As Boolean
'    Dim nf As Integer
'    Dim lin As String
'    Dim mSuministro As Suministro
'On Error GoTo err_exportarXML
'    '-- comprobamos que hay algo para exportar
'    If mvarsuministros Is Nothing Then Exit Function
'    If mvarsuministros.Count = 0 Then Exit Function
'    '-- obtenemos el fichero en el que grabaremos el XML
'    nf = FreeFile()
'    Open App.path & "\D" & mvarid & ".XML" For Output As nf
'    '-- comenzamos a grabar los datos de cabecera
'        Print #nf, "<?xml version=""1.0"" encoding=""UTF-8""?>"
'        Print #nf, "<!--Generated by Arigasol (c) Ariadna Software S.L. (http://www.ariadnasoftware.com)-->"
'        lin = "<SuministrV2Ent CodEE=""" & mvarcodee & """ Id= """ & mvarid & """"
'        If mvartest = "S" Then lin = lin & " Test=""S"""
'        lin = lin & " xsi:schemaLocation=""https://www3.aeat.es/ADUA/internet/es/aeat/dit/adu/iegasolprof/SuministrV2Ent.xsd https://www3.aeat.es/ADUA/internet/es/aeat/dit/adu/iegasolprof/SuministrV2Ent.xsd"""
'        lin = lin & " xmlns=""https://www3.aeat.es/ADUA/internet/es/aeat/dit/adu/iegasolprof/SuministrV2Ent.xsd"""
'        lin = lin & " xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
'        Print #nf, lin
'        Set mSuministro = New Suministro
'        For Each mSuministro In mvarsuministros
'            Print #nf, "<Suministro>"
'                With mSuministro
'                    '-- hemos modificado este valor para que sea único (19/12/2007)[RAFA]
'                    Print #nf, "<IdMovCont>" & Left(mvarid, 14) & Format(.idmovcont, "000") & "</IdMovCont>"
'                    Print #nf, "<CIM>" & .Cim & "</CIM>"
'                    Print #nf, "<Fecha>" & Format(.Fecha, "yyyymmdd") & "</Fecha>"
'                    Print #nf, "<Hora>" & Format(.Hora, "hhmm") & "</Hora>"
'                    Print #nf, "<CodPro>" & Format(.codpro, "000") & "</CodPro>"
'                    Print #nf, "<Lit>" & TransformaComasPuntos(Format(.lit, "0.00")) & "</Lit>"
'                    Print #nf, "<NIF>" & .NIF & "</NIF>"
'                    Print #nf, "<Matricula>" & .Matricula & "</Matricula>"
'                End With
'            Print #nf, "</Suministro>"
'        Next
'        Print #nf, "</SuministrV2Ent>"
'    Close #nf
'    exportarXML = True
'    Exit Function
'err_exportarXML:
'    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.exportarXML"
'End Function

Public Function exportarXML() As Boolean
    Dim nf As Integer
    Dim lin As String
    Dim mSuministro As Suministro
On Error GoTo err_exportarXML
    '-- comprobamos que hay algo para exportar
    If mvarsuministros Is Nothing Then Exit Function
    If mvarsuministros.Count = 0 Then Exit Function
    '-- obtenemos el fichero en el que grabaremos el XML
    nf = FreeFile()
    Open App.path & "\D" & mvarid & ".XML" For Output As nf
    '-- comenzamos a grabar los datos de cabecera
        Print #nf, "<?xml version=""1.0"" encoding=""UTF-8""?>"
        Print #nf, "<!--Generated by Arigasol (c) Ariadna Software S.L. (http://www.ariadnasoftware.com)-->"
        lin = "<SuministrV3Ent" ' estamos en la versión 3
        lin = lin & " xmlns=""https://www3.aeat.es/ADUA/internet/es/aeat/dit/adu/adg0/gprof/sw/SuministrV3Ent.xsd"""
        lin = lin & " xmlns:td=""https://www3.aeat.es/ADUA/internet/es/aeat/dit/adu/adg0/gprof/sw/SuministrV3Dat.xsd"""
        lin = lin & " CodEE=""" & mvarcodee & """ Id= """ & mvarid & """"
        If mvartest = "S" Then
            lin = lin & " Test=""S"">"
        Else
            lin = lin & ">" ' Para cerrar el ID
        End If
        Print #nf, lin
        Set mSuministro = New Suministro
        For Each mSuministro In mvarsuministros
            Print #nf, "<Suministro>"
                With mSuministro
                    '-- hemos modificado este valor para que sea único (19/12/2007)[RAFA]
                    Print #nf, "<IdMovCont>" & Left(mvarid, 14) & Format(.idmovcont, "000") & "</IdMovCont>"
                    Print #nf, "<CIM>" & .Cim & "</CIM>"
                    Print #nf, "<Fecha>" & Format(.Fecha, "yyyymmdd") & "</Fecha>"
                    Print #nf, "<Hora>" & Format(.Hora, "hhmm") & "</Hora>"
                    Print #nf, "<CodPro>" & Format(.codpro, "000") & "</CodPro>"
                    Print #nf, "<Lit>" & TransformaComasPuntos(Format(.lit, "0.00")) & "</Lit>"
                    Print #nf, "<NIF>" & .NIF & "</NIF>"
                    Print #nf, "<Matricula>" & .Matricula & "</Matricula>"
                End With
            Print #nf, "</Suministro>"
        Next
        Print #nf, "</SuministrV3Ent>"
    Close #nf
    exportarXML = True
    Exit Function
err_exportarXML:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.exportarXML"
End Function


Public Function guardar2(Conn As ADODB.Connection) As Boolean
    Dim vSuministro As Suministro
On Error GoTo err_guardar2
    '-- Obtenemos el CEE
    If Not obtenerCEE(Conn) Then
        MsgBox "No se puede obtener el CEE. Revise parámetros", vbExclamation
        Exit Function
    End If
    sql = "delete from gp_suministrv2ent2 where id = '" & mvarid & "'"
    Conn.Execute sql
    sql = "insert into gp_suministrv2ent2 (id,codee,test,situacion)" & _
            " values (" & _
            "'" & mvarid & "'," & _
            "'" & mvarcodee & "'," & _
            "'" & mvartest & "'," & _
            "0" & _
            ")"
    Conn.Execute sql ' ya tenemos la cabecera
    For Each vSuministro In mvarsuministros
        If Not vSuministro.guardar2(Conn, True) Then
            Exit Function
        End If
    Next
    guardar2 = True
    Exit Function
err_guardar2:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.guardar2"
End Function

Public Function exportarXML_V2() As Boolean
    '-----------------------------------------------------
    '   Esta es la versión V2 de la exportación XML
    '   Como ya estamos en la V3, esta quedará obsoleta
    '-----------------------------------------------------
    Dim nf As Integer
    Dim lin As String
    Dim mSuministro As Suministro
On Error GoTo err_exportarXML_V2
    '-- comprobamos que hay algo para exportar
    If mvarsuministros Is Nothing Then Exit Function
    If mvarsuministros.Count = 0 Then Exit Function
    '-- obtenemos el fichero en el que grabaremos el XML
    nf = FreeFile()
    Open App.path & "\D" & mvarid & ".XML" For Output As nf
    '-- comenzamos a grabar los datos de cabecera
        Print #nf, "<?xml version=""1.0"" encoding=""UTF-8""?>"
        Print #nf, "<!--Generated by Arigasol (c) Ariadna Software S.L. (http://www.ariadnasoftware.com)-->"
        lin = "<SuministrV2Ent CodEE=""" & mvarcodee & """ Id= """ & mvarid & """"
        If mvartest = "S" Then lin = lin & " Test=""S"""
        lin = lin & " xsi:schemaLocation=""https://www3.aeat.es/ADUA/internet/es/aeat/dit/adu/iegasolprof/SuministrV2Ent.xsd https://www3.aeat.es/ADUA/internet/es/aeat/dit/adu/iegasolprof/SuministrV2Ent.xsd"""
        lin = lin & " xmlns=""https://www3.aeat.es/ADUA/internet/es/aeat/dit/adu/iegasolprof/SuministrV2Ent.xsd"""
        lin = lin & " xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
        Print #nf, lin
        Set mSuministro = New Suministro
        For Each mSuministro In mvarsuministros
            Print #nf, "<Suministro>"
                With mSuministro
                    '-- hemos modificado este valor para que sea único (19/12/2007)[RAFA]
                    Print #nf, "<IdMovCont>" & Left(mvarid, 14) & Format(.idmovcont, "000") & "</IdMovCont>"
                    Print #nf, "<CIM>" & .Cim & "</CIM>"
                    Print #nf, "<Fecha>" & Format(.Fecha, "yyyymmdd") & "</Fecha>"
                    Print #nf, "<Hora>" & Format(.Hora, "hhmm") & "</Hora>"
                    Print #nf, "<CodPro>" & Format(.codpro, "000") & "</CodPro>"
                    Print #nf, "<Lit>" & TransformaComasPuntos(Format(.lit, "0.00")) & "</Lit>"
                    Print #nf, "<NIF>" & .NIF & "</NIF>"
                    Print #nf, "<Matricula>" & .Matricula & "</Matricula>"
                End With
            Print #nf, "</Suministro>"
        Next
        Print #nf, "</SuministrV2Ent>"
    Close #nf
    exportarXML_V2 = True
    Exit Function
err_exportarXML_V2:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.exportarXML_V2"
End Function


Public Function guardar3(Conn As ADODB.Connection) As Boolean
    Dim vSuministro As Suministro
On Error GoTo err_guardar3
    '-- Generamos el valor de ID
    mvarid = Format(Now, "yyyymmddhhnnss000003")
    '-- Obtenemos el CEE
    If Not obtenerCEE(Conn) Then
        MsgBox "No se puede obtener el CEE. Revise parámetros", vbExclamation
        Exit Function
    End If
    sql = "delete from gp_suministrv2ent2 where id = '" & mvarid & "'"
    Conn.Execute sql
    sql = "insert into gp_suministrv2ent2 (id,codee,test,situacion)" & _
            " values (" & _
            "'" & mvarid & "'," & _
            "'" & mvarcodee & "'," & _
            "'" & mvartest & "'," & _
            "0" & _
            ")"
    Conn.Execute sql ' ya tenemos la cabecera
    For Each vSuministro In mvarsuministros
        vSuministro.ID = mvarid
        If Not vSuministro.guardar3(Conn, True) Then
            Exit Function
        End If
    Next
    guardar3 = True
    Exit Function
err_guardar3:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Declaracion.guardar3"
End Function

