Attribute VB_Name = "bus"
'NOTA: en este m�dul, adem�s, n'hi han funcions generals que no siguen de formularis (molt b�)
Option Explicit

'Definicion Conexi�n a BASE DE DATOS
'---------------------------------------------------
'Conexi�n a la BD AriGasol de la empresa
Public Conn As ADODB.Connection

'Conexi�n a la BD de Usuarios
'Public ConnUsuarios As ADODB.Connection

'Conexi�n a la BD de Contabilidad de la empresa conectada
Public ConnConta As ADODB.Connection

'Conexi�n a la BD de Contabilidad de otra empresa distinta a la conectada
Public ConnAuxCon As ADODB.Connection


'Que conexion a base de datos se va a utilizar
Public Const cPTours As Byte = 1 'trabajaremos con conn (conexion a BD AriGasol)
Public Const cConta As Byte = 2 'trabajaremos con connConta (cxion a BD Contabilidad)

Public Const vbFPTransferencia = 1
Public Const ValorNulo = "Null"

Public miRsAux As ADODB.Recordset

'Definicion de clases de la aplicaci�n
'-----------------------------------------------------
Public vParamAplic As CParamAplic   'parametros de la aplicacion



'Definicion de FORMATOS
'---------------------------------------------------
Public FormatoFecha As String
Public FormatoHora As String
Public FormatoImporte As String 'Decimal(12,2)
Public FormatoPrecio As String 'Decimal(8,3)
'Public FormatoCantidad As String 'Decimal(10,2)
Public FormatoPorcen As String 'Decimal(5,2) 'Porcentajes
Public FormatoExp As String  'Expedientes

Public FormatoDec10d2 As String 'Decimal(10,2)
Public FormatoDec10d3 As String 'Decimal(10,3)
Public FormatoDec5d4 As String 'Decimal(5,4)
Public FormatoDec10d5 As String 'Decimal(10,5)

Public FIni As String
Public FFin As String

'Public FormatoKms As String 'Decimal(8,4)


Public teclaBuscar As Integer 'llamada desde prismaticos

Public CadenaDesdeOtroForm As String

'Global para n� de registro eliminado
Public NumRegElim  As Long

'publica para almacenar control cambios en registros de formularios
'se utiliza en InsertarCambios
Public CadenaCambio As String
Public ValorAnterior As String

Public MensError As String

'Para algunos campos de texto suletos controlarlos
'Public miTag As CTag

'Variable para saber si se ha actualizado algun asiento
'Public AlgunAsientoActualizado As Boolean
'Public TieneIntegracionesPendientes As Boolean

'Public miRsAux As ADODB.Recordset

Public AnchoLogin As String  'Para fijar los anchos de columna
Public ImpresoraDefecto As String


' **** DATOS DEL LOGIN ****
'Public CodEmple As Integer
'Public codAgenc As Integer
'Public codEmpre As Integer
'Public codGrupo As Integer
'Public claEmpre As Integer
'Public TipEmple As Integer
'Public areEmple As Integer
' *************************


'Inicio Aplicaci�n
Public Sub Main()

     
    FormatoFecha = "yyyy-mm-dd"
     
     
     
     'obric la conexio
    If AbrirConexionAriGasol("root", "aritel") = False Then
        MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
        End
    End If


    Set vParamAplic = New CParamAplic
    If vParamAplic.Leer = 1 Then
        MsgBox "No se han podido cargar los par�metros de contabilidad. Debe configurar la aplicaci�n.", vbExclamation
        
        Set vParamAplic = Nothing
        Exit Sub
    Else
    If vParamAplic.NumeroConta <> 0 Then
    
        'Abrir conexi�n a la BDatos de Contabilidad para acceder a
        'Tablas: Cuentas, Tipos IVA,...
        If AbrirConexionConta(vParamAplic.UsuarioConta, vParamAplic.PasswordConta) = False Then
            MsgBox "La aplicaci�n no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
            AccionesCerrar
            End
        End If
'    ' ### [Monica] 06/09/2006
'    ' comento los niveles de contabilidad pq solo tengo las cuentas
'        If vEmpresa.LeerNiveles() = False Then  'De Contabilidad
'            MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicaci�n.", vbExclamation
'            AccionesCerrar
'            End
'        End If
'
'        FechasEjercicioConta FIni, FFin
'
    End If
    End If
'    Set vParam = New Cparametros
'    If vParam.Leer = False Then   'De AriGasol
'        MsgBox "No se han podido cargar los par�metros de la empresa. Debe configurar la aplicaci�n.", vbExclamation
'        Set vEmpresa = Nothing
'        Set vSesion = Nothing
'        Set Conn = Nothing
'        End
'    End If



'    If AbrirConexionConta("root", "aritel") = False Then
'        MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
'        End
'    End If

    frmFacturas.Show vbModal


End Sub

Public Function ComprovaVersio() As Boolean
  
'    Dim RS2 As Recordset
'    Dim cad2 As String
'    Dim major_ul As Integer
'    Dim minor_ul As Integer
'    Dim revis_ul As Integer
'
'    ComprovaVersio = False
'
'    cad2 = "SELECT * FROM ulversio"
'
'    Set RS2 = New ADODB.Recordset
'    RS2.Open cad2, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    If Not RS2.EOF Then
'        major_ul = RS2.Fields!major_ul
'        minor_ul = RS2.Fields!minor_ul
'        revis_ul = RS2.Fields!revis_ul
'    Else
'        MsgBox "Error al consultar la �ltima versi�n disponible", vbCritical
'        'ulVersio = False
'        Exit Function
'    End If
'
'    RS2.Close
'    Set RS2 = Nothing
'
'    If (App.Major <> major_ul) Or (App.Minor <> minor_ul) Or (App.Revision <> revis_ul) Then
'        ComprovaVersio = True
'    End If
'
'    Exit Function
    
End Function

'espera els segon que li digam
Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function


Public Function AbrirConexionAriGasol(Usuario As String, Pass As String) As Boolean
Dim cad As String
On Error GoTo EAbrirConexion
    
    AbrirConexionAriGasol = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient
    Conn.CursorLocation = adUseServer
'    cad = "DSN=plannertours;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=plannertours;SERVER=srvcentral;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"

'    cad = "DSN=arigasol;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=arigasol;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    cad = "DSN=arigasol;DESC=MySQL ODBC 3.51 Driver DSN;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    cad = cad & ";Persist Security Info=true"
    
    Conn.ConnectionString = cad
    Conn.Open
    AbrirConexionAriGasol = True
    Exit Function
    
EAbrirConexion:
    MsgBox Err.Number, "Abrir conexi�n AriGasol.", Err.Description
End Function


Public Function AbrirConexionConta(Usuario As String, Pass As String) As Boolean
Dim cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad
On Error GoTo EAbrirConexion
    
    AbrirConexionConta = False
    
    Set ConnConta = Nothing
    Set ConnConta = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

' ### [Monica] 06/09/2006
    nomConta = "conta" & vParamAplic.NumeroConta
'    vEmpresa.BDConta = nomConta
    
    If vParamAplic.NumeroConta <> 0 Then
'        cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamAplic.NumeroConta & ";SERVER=" & vParamAplic.ServidorConta & ";"
'        cad = cad & ";UID=" & vParamAplic.UsuarioConta
'        cad = cad & ";PWD=" & vParamAplic.PasswordConta
'        cad = cad & ";PORT=3306;OPTION=3;STMT=;"
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    
        If vParamAplic.ServidorConta <> "" Then  'especificamos servidor
            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & vParamAplic.ServidorConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        Else 'por defecto cogera la BD del servidor que haya en el ODBC
            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        End If
        cad = cad & ";Persist Security Info=true"

        ConnConta.ConnectionString = cad
        ConnConta.Open
        ConnConta.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionConta = True
    Else
        AbrirConexionConta = False
    End If
    Exit Function
EAbrirConexion:
    MsgBox Err.Number, "Abrir conexi�n Contabilidad.", Err.Description
End Function

Public Function AbrirConexionConta2(Usuario As String, Pass As String) As Boolean
'Abre

Dim cad As String
On Error GoTo EAbrirConexion

    AbrirConexionConta2 = False
    Set ConnConta = Nothing
    Set ConnConta = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
                       
'    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamAplic.NumeroConta & ";SERVER=" & vParamAplic.ServidorConta & ";"
'    cad = cad & ";UID=" & vParamAplic.UsuarioConta
'    cad = cad & ";PWD=" & vParamAplic.PasswordConta
'    '---- Laura: 29/09/2006
'    cad = cad & ";PORT=3306;OPTION=3;STMT=;"
'    '----
'    '++monica: tema de vista
'    cad = cad & "Persist Security Info=true"
    
    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta" & vParamAplic.NumeroConta & ";SERVER=" & vParamAplic.ServidorConta & ";"
    cad = cad & ";UID=" & vParamAplic.UsuarioConta
    cad = cad & ";PASSWORD=" & vParamAplic.PasswordConta
    cad = cad & ";PORT=3306;OPTION=3;STMT=;"
    cad = cad & ";Persist Security Info=true"
    
    ConnConta.ConnectionString = cad
    ConnConta.Open
    ConnConta.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionConta2 = True
    Exit Function
EAbrirConexion:
    MsgBox Err.Number, "Abrir conexi�n contabilidad.", Err.Description
End Function

Public Function AbrirConexionAuxCon(Empresa As String, Usuario As String, Pass As String) As Boolean
Dim cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad
On Error GoTo EAbrirConexion

    AbrirConexionAuxCon = False

    Set ConnAuxCon = Nothing
    Set ConnAuxCon = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnAuxCon.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

    'Obtener la BD de contabilidad
'    SQL = "select bdaconta FROM paramcon WHERE codempre=" & codEmpre
    serConta = "serconta"
    nomConta = DevuelveDesdeBDNew(2, "sparam", "bdaconta", "codempre", Empresa, "N", serConta)
'    vEmpresa.BDConta = nomConta
    If nomConta <> "" Then
    '    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamConta.NumeroConta & ";SERVER=" & vParamConta.ServidorConta & ";"
    '    cad = cad & ";UID=" & vParamConta.UsuarioConta
    '    cad = cad & ";PWD=" & vParamConta.PasswordConta
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        If serConta <> "" Then 'especificamos servidor
            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & serConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        Else 'por defecto cogera la BD del servidor que haya en el ODBC
            cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        End If
        ConnAuxCon.ConnectionString = cad
        ConnAuxCon.Open
        ConnAuxCon.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionAuxCon = True
    Else
        AbrirConexionAuxCon = False
    End If
    Exit Function
EAbrirConexion:
    MsgBox Err.Number, "Abrir conexi�n Contabilidad.", Err.Description
End Function

Public Function CerrarConexionConta()
  'Cerramos la conexion con BD: Contabilidad
  On Error Resume Next
   ConnConta.Close
   If Err.Number <> 0 Then Err.Clear
End Function

Public Sub LeerDatosEmpresa()
'Crea instancia de la clase Cempresa con los valores en
'Tabla: Empresas
'BDatos: PTours y Conta
 
    ' ### [Monica] 06/09/2006
    ' a�adido
    Set vParamAplic = New CParamAplic
    If vParamAplic.Leer = 1 Then
        MsgBox "No se han podido cargar los par�metros de contabilidad. Debe configurar la aplicaci�n.", vbExclamation
        
        Set vParamAplic = Nothing
        Exit Sub
    Else
    If vParamAplic.NumeroConta <> 0 Then
    
        'Abrir conexi�n a la BDatos de Contabilidad para acceder a
        'Tablas: Cuentas, Tipos IVA,...
        If AbrirConexionConta(vParamAplic.UsuarioConta, vParamAplic.PasswordConta) = False Then
            MsgBox "La aplicaci�n no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
            AccionesCerrar
            End
        End If
    ' ### [Monica] 06/09/2006
    ' comento los niveles de contabilidad pq solo tengo las cuentas
'        If vEmpresa.LeerNiveles() = False Then  'De Contabilidad
'            MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicaci�n.", vbExclamation
'            AccionesCerrar
'            End
'        End If
        
'        FechasEjercicioConta FIni, FFin
    
    End If
    End If
'    Set vParam = New Cparametros
'    If vParam.Leer = False Then   'De AriGasol
'        MsgBox "No se han podido cargar los par�metros de la empresa. Debe configurar la aplicaci�n.", vbExclamation
'        Set vEmpresa = Nothing
'        Set vSesion = Nothing
'        Set Conn = Nothing
'        End
'    End If
End Sub


Public Function PonerDatosPpal()
'    If Not vEmpresa Is Nothing Then
'        MDIppal.Caption = "AriGasol" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre
'    End If
    If Err.Number <> 0 Then MsgBox Err.Description, "Poniendo datos de la pantalla principal", Err.Description
End Function

    


Public Function DBSet(vData As Variant, tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim cad As String

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If

        If tipo <> "" Then
            Select Case tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        cad = (CStr(vData))
                        NombreSQL cad
                        DBSet = "'" & cad & "'"
                    End If
                    
                Case "N"    'Numero
                    If vData = "" Or vData = 0 Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        cad = CStr(ImporteFormateado(CStr(vData)))
                        DBSet = TransformaComasPuntos(cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If
                    
                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If
                    
                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                    
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
End Function

Public Function DBLetMemo(vData As Variant) As Variant
    On Error Resume Next
    
    DBLetMemo = vData
    
    
    
    If Err.Number <> 0 Then
        Err.Clear
        DBLetMemo = ""
    End If
End Function



Public Function DBLet(vData As Variant, Optional tipo As String) As Variant
'Para cuando recupera Datos de la BD
    If IsNull(vData) Then
        DBLet = ""
        If tipo <> "" Then
            Select Case tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = 0
                Case "F"    'Fecha
                     '==David
'                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    Conn.Execute "commit"
    Conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    Conn.Execute "commit"
    Conn.Execute "set autocommit=1"
End Sub

'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256,98
'   Tiene que venir num�rico
Public Function ImporteFormateado(Importe As String) As Double
Dim i As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            i = InStr(1, Importe, ".")
            If i > 0 Then Importe = Mid(Importe, 1, i - 1) & Mid(Importe, i + 1)
        Loop Until i = 0
        ImporteFormateado = Importe
    End If
End Function

' ### [Monica] 11/09/2006
Public Function ImporteSinFormato(cadena As String) As String
Dim i As Integer
'Quitamos puntos
Do
    i = InStr(1, cadena, ".")
    If i > 0 Then cadena = Mid(cadena, 1, i - 1) & Mid(cadena, i + 1)
Loop Until i = 0
ImporteSinFormato = TransformaPuntosComas(cadena)
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(cadena As String) As String
Dim i As Integer
    Do
        i = InStr(1, cadena, ",")
        If i > 0 Then
            cadena = Mid(cadena, 1, i - 1) & "." & Mid(cadena, i + 1)
        End If
    Loop Until i = 0
    TransformaComasPuntos = cadena
End Function

'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef cadena As String)
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, cadena, "'")
        If i > 0 Then
            Aux = Mid(cadena, 1, i - 1) & "\"
            cadena = Aux & Mid(cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
End Sub

Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim cad As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    End If
    If IsDate(cad) Then
        EsFechaOKString = True
        T = Format(cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function

Public Function DevNombreSQL(cadena As String) As String
Dim J As Integer
Dim i As Integer
Dim Aux As String
    J = 1
    Do
        i = InStr(J, cadena, "'")
        If i > 0 Then
            Aux = Mid(cadena, 1, i - 1) & "\"
            cadena = Aux & Mid(cadena, i)
            J = i + 2
        End If
    Loop Until i = 0
    DevNombreSQL = cadena
End Function


Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional tipo As String, Optional ByRef otroCampo As String) As String
    Dim Rs As Recordset
    Dim cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    cad = "Select " & kCampo
    If otroCampo <> "" Then cad = cad & ", " & otroCampo
    cad = cad & " FROM " & Ktabla
    cad = cad & " WHERE " & Kcodigo & " = "
    If tipo = "" Then tipo = "N"
    Select Case tipo
    Case "N"
        'No hacemos nada
        cad = cad & ValorCodigo
    Case "T", "F"
        cad = cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveDesdeBD = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MsgBox Err.Number, "Devuelve DesdeBD.", Err.Description
End Function



''Este metodo sustituye a DevuelveDesdeBD
''Funciona para claves primarias formadas por 2 campos
'Public Function DevuelveDesdeBDnew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String) As String
''IN: vBD --> Base de Datos a la que se accede
'Dim RS As Recordset
'Dim cad As String
'Dim Aux As String
'
'On Error GoTo EDevuelveDesdeBDnew
'    DevuelveDesdeBDnew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
'    cad = "Select " & kCampo
'    If otroCampo <> "" Then cad = cad & ", " & otroCampo
'    cad = cad & " FROM " & Ktabla
'    cad = cad & " WHERE " & Kcodigo1 & " = "
'    If tipo1 = "" Then tipo1 = "N"
'    Select Case tipo1
'        Case "N"
'            'No hacemos nada
'            If IsNumeric(valorCodigo1) Then
'                cad = cad & Val(valorCodigo1)
'            Else
'                MsgBox "El campo debe ser num�rico.", vbExclamation
'                DevuelveDesdeBDnew = "Error"
'                Exit Function
'            End If
'        Case "T", "F"
'            cad = cad & "'" & valorCodigo1 & "'"
'        Case Else
'            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
'            Exit Function
'    End Select
'
'    If KCodigo2 <> "" Then
'        cad = cad & " AND " & KCodigo2 & " = "
'        If tipo2 = "" Then tipo2 = "N"
'        Select Case tipo2
'        Case "N"
'            'No hacemos nada
'            If ValorCodigo2 = "" Then
'                cad = cad & "-1"
'            Else
'                cad = cad & Val(ValorCodigo2)
'            End If
'        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
'        Case "F"
'            cad = cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
'        Case Else
'            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
'            Exit Function
'        End Select
'    End If
'
'
'    'Creamos el sql
'    Set RS = New ADODB.Recordset
'
'    Select Case vBD
'        Case cPTours 'vBD=1: PlannerTours
'            RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        Case cConta 'BD 2: Contabilidad
'            RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
'        Case 3 'vBD=3: contabilidad distinta a la de la empresa conectada
'            RS.Open cad, ConnAuxCon, adOpenForwardOnly, adLockOptimistic, adCmdText
'    End Select
''    If vBD = cPTours Then 'vBD=1: PlannerTours
''        RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
''    ElseIf vBD = cConta Then  'BD 2: Contabilidad
''        RS.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
''    End If
'
'    If Not RS.EOF Then
'        DevuelveDesdeBDnew = DBLet(RS.Fields(0))
'        If otroCampo <> "" Then otroCampo = DBLet(RS.Fields(1))
'    End If
'    RS.Close
'    Set RS = Nothing
'    Exit Function
'
'EDevuelveDesdeBDnew:
'        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
'End Function


'LAURA
'Este metodo sustituye a DevuelveDesdeBD
'Funciona para claves primarias formadas por 3 campos
Public Function DevuelveDesdeBDNew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim Rs As Recordset
Dim cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    cad = "Select " & kCampo
    If otroCampo <> "" Then cad = cad & ", " & otroCampo
    cad = cad & " FROM " & Ktabla
    If Kcodigo1 <> "" Then
        cad = cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            cad = cad & Val(valorCodigo1)
        Case "T"
            cad = cad & DBSet(valorCodigo1, "T")
        Case "F"
            cad = cad & DBSet(valorCodigo1, "F")
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If
    
    If KCodigo2 <> "" Then
        cad = cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            cad = cad & DBSet(ValorCodigo2, "T")
        Case "F"
            cad = cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        cad = cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo3)
            End If
        Case "T"
            cad = cad & "'" & ValorCodigo3 & "'"
        Case "F"
            cad = cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    If vBD = cPTours Then 'BD 1: Ariges
        Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else    'BD 2: Conta
        Rs.Open cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    
    If Not Rs.EOF Then
        DevuelveDesdeBDNew = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
EDevuelveDesdeBDnew:
        MsgBox Err.Number, "Devuelve DesdeBD.", Err.Description
End Function





Public Function EsEntero(Texto As String) As Boolean
Dim i As Integer
Dim c As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEntero = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        c = 0
        L = 1
        Do
            i = InStr(L, Texto, ".")
            If i > 0 Then
                L = i + 1
                c = c + 1
            End If
        Loop Until i = 0
        If c > 1 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If c = 0 Then
            L = 1
            Do
                i = InStr(L, Texto, ",")
                If i > 0 Then
                    L = i + 1
                    c = c + 1
                End If
            Loop Until i = 0
            If c > 1 Then res = False
        End If
        
    End If
        EsEntero = res
End Function

Public Function TransformaPuntosComas(cadena As String) As String
    Dim i As Integer
    Do
        i = InStr(1, cadena, ".")
        If i > 0 Then
            cadena = Mid(cadena, 1, i - 1) & "," & Mid(cadena, i + 1)
        End If
        Loop Until i = 0
    TransformaPuntosComas = cadena
End Function

Public Sub InicializarFormatos()
    FormatoFecha = "yyyy-mm-dd"
    FormatoHora = "hh:mm:ss"
'    FormatoFechaHora = "yyyy-mm-dd hh:mm:ss"
    FormatoImporte = "#,###,###,##0.00"  'Decimal(12,2)
    FormatoPrecio = "##,##0.000"  'Decimal(8,3) antes decimal(10,4)
'    FormatoCantidad = "##,###,##0.00"   'Decimal(10,2)
    FormatoPorcen = "##0.00" 'Decima(5,2) para porcentajes
    
    FormatoDec10d2 = "##,###,##0.00"   'Decimal(10,2)
    FormatoDec10d3 = "##,###,##0.000"   'Decimal(10,3)
    FormatoDec5d4 = "0.0000"   'Decimal(5,4)
    FormatoDec10d5 = "##,##0.00000"   'Decimal(10,5)
    FormatoExp = "0000000000"
'    FormatoKms = "#,##0.00##" 'Decimal(8,4)
End Sub


Public Sub AccionesCerrar()
'cosas que se deben hacen cuando finaliza la aplicacion
    On Error Resume Next
    
    'cerrar clases q estan abiertas durante la ejecucion
    
'    Set vParam = Nothing
'    Set vParamAplic = Nothing
'    Set vParamConta = Nothing
    
    
    'Cerrar Conexiones a bases de datos
    Conn.Close
    ConnConta.Close
    Set Conn = Nothing
    Set ConnConta = Nothing
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub CommitConexion()
On Error Resume Next
    Conn.Execute "Commit"
    If Err.Number <> 0 Then Err.Clear
End Sub

'------------------------------------------------------------------
'   Comprobara si una daterminada fecha esta o no en los ejercicios
'   contables (actual y siguiente)
'   Dando un O: SI. Correcto. Ok
'            1: Inferior
'            2: Superior

Public Function EsFechaOKConta(Fecha As Date) As Byte
Dim F2 As Date
Dim Orden1 As String
Dim Orden2 As String

    Orden1 = ""
    Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")

    Orden2 = ""
    Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")

    If CDate(Orden1) > Fecha Then
        EsFechaOKConta = 1
    Else
        F2 = DateAdd("yyyy", 1, CDate(Orden2))
        If Fecha > F2 Then
            EsFechaOKConta = 2
        Else
            'OK. Dentro de los ejercicios contables
            EsFechaOKConta = 0
        End If
    End If

End Function


Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim ent As Integer
Dim cad As String

  ' Comprobaciones

  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un n�mero."
    Exit Function
  End If

  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If

  ' Redondeo.

  cad = "0"
  If NumDigitsAfterDecimals <> 0 Then cad = cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Val(TransformaComasPuntos(Format(Number, cad)))

End Function


