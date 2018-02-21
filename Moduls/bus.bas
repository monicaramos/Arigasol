Attribute VB_Name = "bus"
'NOTA: en este mòdul, ademés, n'hi han funcions generals que no siguen de formularis (molt bé)
Option Explicit

'Definicion Conexión a BASE DE DATOS
'---------------------------------------------------
'Conexión a la BD AriGasol de la empresa
Public Conn As ADODB.Connection

'Conexión a la BD de Usuarios
'Public ConnUsuarios As ADODB.Connection

'Conexión a la BD de Contabilidad de la empresa conectada
Public ConnConta As ADODB.Connection

'Conexión a la BD de Contabilidad de otra empresa distinta a la conectada
Public ConnAuxCon As ADODB.Connection


'Que conexion a base de datos se va a utilizar
Public Const cPTours As Byte = 1 'trabajaremos con conn (conexion a BD AriGasol)
Public Const cConta As Byte = 2 'trabajaremos con connConta (cxion a BD Contabilidad)

Public Const vbFPTransferencia = 1

Public miRsAux As ADODB.Recordset

'Definicion de clases de la aplicación
'-----------------------------------------------------
Public vEmpresa As Cempresa  'Los datos de la empresa
Public vParamAplic As CParamAplic   'parametros de la aplicacion
Public vSesion As CSesion   'Los datos del usuario que hizo login

'LOG de acciones relevantes
Public LOG As cLOG   'Se instancia , se ejecuta LOG.insertar y se elimina :LOG=nothing   Ver ejemplo borre facturas
Public pPdfRpt As String

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

'Global para nº de registro eliminado
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

Public Const SerieFraPro = "1"

Public ResultadoFechaContaOK As Byte
Public MensajeFechaOkConta As String

Public DesdeCierreTurno As Boolean

' **** DATOS DEL LOGIN ****
'Public CodEmple As Integer
'Public codAgenc As Integer
'Public codEmpre As Integer
'Public codGrupo As Integer
'Public claEmpre As Integer
'Public TipEmple As Integer
'Public areEmple As Integer
' *************************


'Inicio Aplicación
Public Sub Main()
Dim SQL As String

'[Monica]19/01/2018: quito esto para que puedan ejecutarse mas de una sesion
'    If App.PrevInstance Then
'        MsgBox "AriGasol ya se esta ejecutando", vbExclamation
'        End
'     End If
     
     'obric la conexio
    If AbrirConexionAriGasol("root", "aritel", "arigasol") = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
        End
    End If
    

    'Necesitaremos el archivo login.dat
    CadenaDesdeOtroForm = ""
    frmIdentifica.Show vbModal
   
    If CadenaDesdeOtroForm = "" Then
        End
    Else
        If CadenaDesdeOtroForm = "OK" Then
            frmLogin.Show vbModal
        Else
        '[Monica]22/12/2015: añadida esta instruccion con el else
            vSesion.CadenaConexion = "arigasol"
        End If
    End If
   
    CerrarConexionArigasol
    CerrarConexionConta
      
      
    If AbrirConexionAriGasol("root", "aritel", vSesion.CadenaConexion) = False Then
          MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
          End
    End If

    'Carga Datos de la Empresa y los Niveles de cuentas de Contabilidad de la empresa
    'Crea la Conexion a la BD de la Contabilidad
    LeerDatosEmpresa

    InicializarFormatos
    teclaBuscar = 43

    Load frmPpal
            
    MDIppal.Show
    
    If vParamAplic.ContabilidadNueva And (vSesion.Nivel = 0 Or vSesion.Nivel = 1) Then FrasPendientesContabilizar False
    
End Sub


Public Function UltimaFechaCorrectaSII(DiasAVisoSII As Integer, FechaPresentacion As Date) As Date
Dim DiaSemanaPresen As Integer
Dim DiaSemanaUltimoDiaPresentar As Integer
Dim F As Date
Dim Resta As Integer
    
    If DiasAVisoSII > 5 Then
        UltimaFechaCorrectaSII = DateAdd("d", -DiasAVisoSII, FechaPresentacion)
    Else
        DiaSemanaPresen = Weekday(FechaPresentacion, vbMonday)
        If DiaSemanaPresen >= 6 Then
            'Si presento el sabado o el domingo tengo mas dias
            If DiaSemanaPresen = 6 Then
                Resta = DiasAVisoSII
            Else
                Resta = DiasAVisoSII + 1
            End If
        Else
            F = DateAdd("d", -DiasAVisoSII, FechaPresentacion)
            DiaSemanaUltimoDiaPresentar = Weekday(F, vbMonday)
            If DiaSemanaUltimoDiaPresentar > DiaSemanaPresen Then
                Resta = DiasAVisoSII + 2
            Else
                'Directamente la resta son 4
                Resta = DiasAVisoSII
            End If
        End If
        UltimaFechaCorrectaSII = DateAdd("d", -Resta, FechaPresentacion)
    End If
    UltimaFechaCorrectaSII = Format(UltimaFechaCorrectaSII, "dd/mm/yyyy")

End Function




Public Sub FrasPendientesContabilizar(EsRecoleccion As Boolean)
Dim SQL As String
Dim sql2 As String
Dim SqlBd As String
Dim SQLinsert As String
Dim RsBd As ADODB.Recordset
Dim BBDD As String

Dim frmMens As frmMensajes

    On Error GoTo eFrasPendientesContabilizar

    SQL = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute SQL

    SQLinsert = "insert into tmpinformes (codusu,codigo1,nombre1,importe1,nombre2,fecha1,importe2) "
        
    SQL = " select " & vSesion.Codigo & ",0, concat(letraser,numfactu),schfac.codsocio, nomsocio, fecfactu, totalfac from schfac inner join ssocio on schfac.codsocio = ssocio.codsocio where intconta = 0 "
    If vEmpresa.TieneSII Then
        SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
    SQL = SQL & " union "
    SQL = SQL & " select " & vSesion.Codigo & ",0, concat(letraser,numfactu),schfacr.codsocio, nomsocio, fecfactu, totalfac from schfacr inner join ssocio on schfacr.codsocio = ssocio.codsocio where intconta = 0 "
    If vEmpresa.TieneSII Then
        SQL = SQL & " and fecfactu >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecfactu <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If
        
    SQL = SQL & " union "
    SQL = SQL & " select " & vSesion.Codigo & ",1, numfactu, scafpc.codprove, scafpc.nomprove, fecrecep, totalfac from scafpc where intconta = 0 "
    If vEmpresa.TieneSII Then
        SQL = SQL & " and fecrecep >= " & DBSet(vEmpresa.SIIFechaInicio, "F") & " and fecrecep <= " & DBSet(DateAdd("d", -1, Now), "F")
    End If

    Conn.Execute SQLinsert & SQL
    
    SQL = "select codusu,codigo1,nombre1,importe1,nombre2,fecha1,importe2 from tmpinformes where codusu = " & vSesion.Codigo
    
    If TotalRegistrosConsulta(SQL) > 0 Then
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 25
        frmMens.Cadena = SQL
        frmMens.Show vbModal
    
        Set frmMens = Nothing
    End If
    Exit Sub
    
eFrasPendientesContabilizar:
    MuestraError Err.Number, "Facturas Pendientes de Integrar a Contabilidad", Err.Description
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
'        MsgBox "Error al consultar la última versión disponible", vbCritical
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


Public Function AbrirConexionAriGasol(Usuario As String, Pass As String, BBDD As String) As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion
    
    AbrirConexionAriGasol = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient
    Conn.CursorLocation = adUseServer
    
'[Monica]12/03/2015: cambiamos la forma de entrar
    Cad = "DSN=arigasol;DESC=MySQL ODBC 3.51 Driver DSN;DESC=;DATABASE=" & BBDD & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    Cad = Cad & ";Persist Security Info=true"
    
'    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= Arigasol;DATABASE=" & BBDD
'    cad = cad & ";UID=" & Usuario
'    cad = cad & ";PWD=" & Pass
'    cad = cad & ";Persist Security Info=true"
      
    Conn.ConnectionString = Cad
    Conn.Open
    AbrirConexionAriGasol = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión AriGasol.", Err.Description
End Function


Public Function AbrirConexionConta(Usuario As String, Pass As String) As Boolean
Dim Cad As String
Dim nomConta As String 'nombre de la BD de la contabilidad
Dim serConta As String 'servidor donde esta la BD de la contabilidad
On Error GoTo EAbrirConexion
    
    AbrirConexionConta = False
    
    Set ConnConta = Nothing
    Set ConnConta = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnConta.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

' ### [Monica] 06/09/2006
    If vParamAplic.ContabilidadNueva Then
        nomConta = "ariconta" & vParamAplic.NumeroConta
    Else
        nomConta = "conta" & vParamAplic.NumeroConta
    End If
'    vEmpresa.BDConta = nomConta
    
    If vParamAplic.NumeroConta <> 0 Then
'        cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=conta" & vParamAplic.NumeroConta & ";SERVER=" & vParamAplic.ServidorConta & ";"
'        cad = cad & ";UID=" & vParamAplic.UsuarioConta
'        cad = cad & ";PWD=" & vParamAplic.PasswordConta
'        cad = cad & ";PORT=3306;OPTION=3;STMT=;"
    '    cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta2;SERVER=david;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    
        If vParamAplic.ServidorConta <> "" Then  'especificamos servidor
            Cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & vParamAplic.ServidorConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        Else 'por defecto cogera la BD del servidor que haya en el ODBC
            Cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        End If
        Cad = Cad & ";Persist Security Info=true"

        ConnConta.ConnectionString = Cad
        ConnConta.Open
        ConnConta.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionConta = True
    Else
        AbrirConexionConta = False
    End If
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión Contabilidad.", Err.Description
End Function

Public Function AbrirConexionConta2(Usuario As String, Pass As String) As Boolean
'Abre

Dim Cad As String
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
    
    Cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=conta" & vParamAplic.NumeroConta & ";SERVER=" & vParamAplic.ServidorConta & ";"
    Cad = Cad & ";UID=" & vParamAplic.UsuarioConta
    Cad = Cad & ";PASSWORD=" & vParamAplic.PasswordConta
    Cad = Cad & ";PORT=3306;OPTION=3;STMT=;"
    Cad = Cad & ";Persist Security Info=true"
    
    ConnConta.ConnectionString = Cad
    ConnConta.Open
    ConnConta.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionConta2 = True
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión contabilidad.", Err.Description
End Function

Public Function AbrirConexionAuxCon(Empresa As String, Usuario As String, Pass As String) As Boolean
Dim Cad As String
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
            Cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";SERVER=" & serConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        Else 'por defecto cogera la BD del servidor que haya en el ODBC
            Cad = "DSN=vConta;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & nomConta & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
        End If
        ConnAuxCon.ConnectionString = Cad
        ConnAuxCon.Open
        ConnAuxCon.Execute "Set AUTOCOMMIT = 1"
        AbrirConexionAuxCon = True
    Else
        AbrirConexionAuxCon = False
    End If
    Exit Function
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión Contabilidad.", Err.Description
End Function

Public Function CerrarConexionArigasol()
  'Cerramos la conexion con BD:
  On Error Resume Next
   Conn.Close
   If Err.Number <> 0 Then Err.Clear
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
 
    Set vEmpresa = New Cempresa
    If vEmpresa.LeerDatos(1) = False Then  'De AriGasol
        MsgBox "No se han podido cargar los datos de la empresa. Debe configurar la aplicación.", vbExclamation
        Set vEmpresa = Nothing
       ' Set vSesion = Nothing
       ' Set conn = Nothing
        Exit Sub
    End If
    
    ' ### [Monica] 06/09/2006
    ' añadido
    Set vParamAplic = New CParamAplic
    If vParamAplic.Leer = 1 Then
        MsgBox "No se han podido cargar los parámetros de gasolinera. Debe configurar la aplicación.", vbExclamation
        
        Set vParamAplic = Nothing
        Exit Sub
    Else
    If vParamAplic.NumeroConta <> 0 Then
    
        'Abrir conexión a la BDatos de Contabilidad para acceder a
        'Tablas: Cuentas, Tipos IVA,...
        If AbrirConexionConta(vParamAplic.UsuarioConta, vParamAplic.PasswordConta) = False Then
            MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
            AccionesCerrar
            End
        End If
    ' ### [Monica] 06/09/2006
    ' comento los niveles de contabilidad pq solo tengo las cuentas
        If vEmpresa.LeerNiveles() = False Then  'De Contabilidad
            MsgBox "No se han podido cargar los niveles de la contabilidad de la empresa. Debe configurar la aplicación.", vbExclamation
            AccionesCerrar
            End
        End If
        
        If vParamAplic.ContabilidadNueva Then
            vEmpresa.BDConta = "ariconta" & vParamAplic.NumeroConta
        Else
            vEmpresa.BDConta = "conta" & vParamAplic.NumeroConta
        End If
        
        
        FechasEjercicioConta FIni, FFin
    
    End If
    End If
End Sub


Public Function PonerDatosPpal()
    If Not vEmpresa Is Nothing Then
        MDIppal.Caption = "AriGasol" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre
    End If
    If Err.Number <> 0 Then MuestraError Err.Description, "Poniendo datos de la pantalla principal", Err.Description
End Function

    

Public Sub MuestraError(numero As Long, Optional Cadena As String, Optional Desc As String)
    Dim Cad As String
    Dim Aux As String
    
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    Cad = "Se ha producido un error: " & vbCrLf
    If Cadena <> "" Then
        Cad = Cad & vbCrLf & Cadena & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If Conn.Errors.Count > 0 Then
        ControlamosError Aux
        Conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then Cad = Cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then Cad = Cad & "Número: " & numero & vbCrLf & "Descripción: " & Error(numero)
    MsgBox Cad, vbExclamation
End Sub

Public Function DBSet(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim Cad As String

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If

        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        Cad = (CStr(vData))
                        NombreSQL Cad
                        DBSet = "'" & Cad & "'"
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
                        Cad = CStr(ImporteFormateado(CStr(vData)))
                        DBSet = TransformaComasPuntos(Cad)
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



Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
'Para cuando recupera Datos de la BD
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
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
'   Tiene que venir numérico
Public Function ImporteFormateado(IMPORTE As String) As Double
Dim I As Integer

    If IMPORTE = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, IMPORTE, ".")
            If I > 0 Then IMPORTE = Mid(IMPORTE, 1, I - 1) & Mid(IMPORTE, I + 1)
        Loop Until I = 0
        ImporteFormateado = IMPORTE
    End If
End Function

' ### [Monica] 11/09/2006
Public Function ImporteSinFormato(Cadena As String) As String
Dim I As Integer
'Quitamos puntos
Do
    I = InStr(1, Cadena, ".")
    If I > 0 Then Cadena = Mid(Cadena, 1, I - 1) & Mid(Cadena, I + 1)
Loop Until I = 0
ImporteSinFormato = TransformaPuntosComas(Cadena)
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(Cadena As String) As String
Dim I As Integer
    Do
        I = InStr(1, Cadena, ",")
        If I > 0 Then
            Cadena = Mid(Cadena, 1, I - 1) & "." & Mid(Cadena, I + 1)
        End If
    Loop Until I = 0
    TransformaComasPuntos = Cadena
End Function

'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef Cadena As String)
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, Cadena, "'")
        If I > 0 Then
            Aux = Mid(Cadena, 1, I - 1) & "\"
            Cadena = Aux & Mid(Cadena, I)
            J = I + 2
        End If
    Loop Until I = 0
End Sub

Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim Cad As String
    
    Cad = T
    If InStr(1, Cad, "/") = 0 Then
        If Len(T) = 8 Then
            Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        Else
            If Len(T) = 6 Then Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        End If
    End If
    If IsDate(Cad) Then
        EsFechaOKString = True
        T = Format(Cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function

Public Function DevNombreSQL(Cadena As String) As String
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, Cadena, "'")
        If I > 0 Then
            Aux = Mid(Cadena, 1, I - 1) & "\"
            Cadena = Aux & Mid(Cadena, I)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSQL = Cadena
End Function


Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef otroCampo As String) As String
    Dim Rs As Recordset
    Dim Cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    Cad = "Select " & kCampo
    If otroCampo <> "" Then Cad = Cad & ", " & otroCampo
    Cad = Cad & " FROM " & Ktabla
    Cad = Cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        Cad = Cad & ValorCodigo
    Case "T", "F"
        Cad = Cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveDesdeBD = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
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
'                MsgBox "El campo debe ser numérico.", vbExclamation
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
Dim Cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    Cad = "Select " & kCampo
    If otroCampo <> "" Then Cad = Cad & ", " & otroCampo
    Cad = Cad & " FROM " & Ktabla
    If Kcodigo1 <> "" Then
        Cad = Cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            Cad = Cad & Val(valorCodigo1)
        Case "T"
            Cad = Cad & DBSet(valorCodigo1, "T")
        Case "F"
            Cad = Cad & DBSet(valorCodigo1, "F")
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If
    
    If KCodigo2 <> "" Then
        Cad = Cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            Cad = Cad & DBSet(ValorCodigo2, "T")
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        Cad = Cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo3)
            End If
        Case "T"
            Cad = Cad & "'" & ValorCodigo3 & "'"
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    If vBD = cPTours Then 'BD 1: Ariges
        Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else    'BD 2: Conta
        Rs.Open Cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    
    If Not Rs.EOF Then
        DevuelveDesdeBDNew = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
EDevuelveDesdeBDnew:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function




'CESAR
Public Function DevuelveDesdeBDnew2(kBD As Integer, kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional num As Byte, Optional ByRef otroCampo As String) As String
Dim Rs As Recordset
Dim Cad As String
Dim Aux As String
Dim v_aux As Integer
Dim campo As String
Dim Valor As String
Dim tip As String

On Error GoTo EDevuelveDesdeBDnew2
DevuelveDesdeBDnew2 = ""

Cad = "Select " & kCampo
If otroCampo <> "" Then Cad = Cad & ", " & otroCampo
Cad = Cad & " FROM " & Ktabla

If Kcodigo <> "" Then Cad = Cad & " where "

For v_aux = 1 To num
    campo = RecuperaValor(Kcodigo, v_aux)
    Valor = RecuperaValor(ValorCodigo, v_aux)
    tip = RecuperaValor(Tipo, v_aux)
        
    Cad = Cad & campo & "="
    If tip = "" Then Tipo = "N"
    
    Select Case tip
            Case "N"
                'No hacemos nada
                Cad = Cad & Valor
            Case "T", "F"
                Cad = Cad & "'" & Valor & "'"
            Case Else
                MsgBox "Tipo : " & tip & " no definido", vbExclamation
            Exit Function
    End Select
    
    If v_aux < num Then Cad = Cad & " AND "
  Next v_aux

'Creamos el sql
Set Rs = New ADODB.Recordset
Select Case kBD
    Case 1
        Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
End Select

If Not Rs.EOF Then
    DevuelveDesdeBDnew2 = DBLet(Rs.Fields(0))
    If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
Else
     If otroCampo <> "" Then otroCampo = ""
End If
Rs.Close
Set Rs = Nothing
Exit Function
EDevuelveDesdeBDnew2:
    MuestraError Err.Number, "Devuelve DesdeBDnew2.", Err.Description
End Function


Public Function EsEntero(Texto As String) As Boolean
Dim I As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEntero = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            I = InStr(L, Texto, ".")
            If I > 0 Then
                L = I + 1
                C = C + 1
            End If
        Loop Until I = 0
        If C > 1 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                I = InStr(L, Texto, ",")
                If I > 0 Then
                    L = I + 1
                    C = C + 1
                End If
            Loop Until I = 0
            If C > 1 Then res = False
        End If
        
    End If
        EsEntero = res
End Function

Public Function TransformaPuntosComas(Cadena As String) As String
    Dim I As Integer
    Do
        I = InStr(1, Cadena, ".")
        If I > 0 Then
            Cadena = Mid(Cadena, 1, I - 1) & "," & Mid(Cadena, I + 1)
        End If
        Loop Until I = 0
    TransformaPuntosComas = Cadena
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
    Set vEmpresa = Nothing
    Set vSesion = Nothing
    
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


Public Function OtrosPCsContraAplicacion() As String
Dim MiRS As Recordset
Dim Cad As String
Dim Equipo As String

    Set MiRS = New ADODB.Recordset
    Cad = "show processlist"
    MiRS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not MiRS.EOF
        If MiRS.Fields(3) = vSesion.CadenaConexion Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vSesion.Codusu Then
                    If Equipo <> "LOCALHOST" Then
                        If InStr(1, Cad, Equipo & "|") = 0 Then Cad = Cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraAplicacion = Cad
End Function


Public Function UsuariosConectados() As Boolean
Dim I As Integer
Dim Cad As String
Dim metag As String
Dim SQL As String
Cad = OtrosPCsContraAplicacion
UsuariosConectados = False
If Cad <> "" Then
    UsuariosConectados = True
    I = 1
    metag = "Los siguientes PC's están conectados a: " & vEmpresa.nomEmpre & " (" & vSesion.CadenaConexion & ")" & vbCrLf & vbCrLf
    Do
        SQL = RecuperaValor(Cad, I)
        If SQL <> "" Then
            metag = metag & "    - " & SQL & vbCrLf
            I = I + 1
        End If
    Loop Until SQL = ""
    MsgBox metag, vbExclamation
End If
End Function

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
    
    '[Monica]20/06/2017: de david
    If EsFechaOKConta = 0 Then
        'Si tiene SII
        If vParamAplic.ContabilidadNueva Then
            If vEmpresa.TieneSII Then
                '[Monica]06/10/2017: añadida la segunda condicion: fecha > vEmpresa.SIIFechaInicio
                '                    fallaba cuando la fecha es anterior a la declaracion del SII.
                '                    Caso de Coopic con una factura interna
'                If DateDiff("d",Fecha, Now) > vEmpresa.SIIDiasAviso And Fecha > vEmpresa.SIIFechaInicio Then

                '[Monica]19/02/2018: comprobamos la fecha para ver si es correcta con la funcion de david con los fines de semana
                If Fecha < UltimaFechaCorrectaSII(vEmpresa.SIIDiasAviso, Now) Then
                    MensajeFechaOkConta = "Fecha fuera de periodo de comunicación SII."
                    'LLEVA SII y han trascurrido los dias
                    If vSesion.Nivel = 0 Then
                        If MsgBox(MensajeFechaOkConta & vbCrLf & "¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then
                            EsFechaOKConta = 4
                        End If
                    Else
                        'NO tienen nivel
                        EsFechaOKConta = 5
                    End If
                End If
            End If
        End If
    Else
        MensajeFechaOkConta = "Fuera de ejercicios contables"
    End If

End Function


Private Function DateDiffSinFinde(FecIni As Date, FecFin As Date)
Dim Finde As Integer
Dim F1 As Date
Dim F2 As Date
Dim difdias As Integer
Dim DiaIni As Integer
Dim I As Integer
    
    Finde = 0
    F1 = FecIni
    F2 = FecFin
    difdias = DateDiff("d", F1, F2)
    DiaIni = Day(F1)
    For I = DiaIni To difdias + 1
        F1 = DateAdd("d", 1, F1)
        If Weekday(F1, vbMonday) >= 6 Then Finde = Finde + 1
    Next I
    
    DateDiffSinFinde = DateDiff("d", F1, F2) + Finde
    
End Function


'--------------------------------------------------------------------
'-------------------------------------------------------------------
'Para el envio de los mails
Public Function PrepararCarpetasEnvioMail(Optional NoBorrar As Boolean) As Boolean
    On Error GoTo EPrepararCarpetasEnvioMail
    PrepararCarpetasEnvioMail = False

    If Dir(App.path & "\temp", vbDirectory) = "" Then
        MkDir App.path & "\temp"
    Else
        If Not NoBorrar Then
            If Dir(App.path & "\temp\*.*", vbArchive) <> "" Then Kill App.path & "\temp\*.*"
        End If
    End If


    PrepararCarpetasEnvioMail = True
    Exit Function
EPrepararCarpetasEnvioMail:
    MuestraError Err.Number, "", "Preparar Carpetas temporal para envio eMail. Borrando tmp "
End Function


Public Function ejecutar(ByRef SQL As String, OcultarMsg As Boolean) As Boolean
    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        If Not OcultarMsg Then MuestraError Err.Number, Err.Description, SQL
        ejecutar = False
    Else
        ejecutar = True
    End If
End Function

Public Function ComprobarEmpresaBloqueada(Codusu As Long, ByRef Empresa As String) As Boolean
Dim Cad As String

ComprobarEmpresaBloqueada = False

'Antes de nada, borramos las entradas de usuario, por si hubiera kedado algo
Conn.Execute "Delete from usuarios.vbloqbd where codusu=" & Codusu

'Ahora comprobamos k nadie bloquea la BD
Cad = DevuelveDesdeBD("codusu", "usuarios.vbloqbd", "conta", Empresa, "T")
If Cad <> "" Then
    'En teoria esta bloqueada. Puedo comprobar k no se haya kedado el bloqueo a medias
    
    Set miRsAux = New ADODB.Recordset
    Cad = "show processlist"
    miRsAux.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        If miRsAux.Fields(3) = Empresa Then
            Cad = miRsAux.Fields(2)
            miRsAux.MoveLast
        End If
    
        'Siguiente
        miRsAux.MoveNext
    Wend
    
    If Cad = "" Then
        'Nadie esta utilizando la aplicacion, luego se puede borrar la tabla
        Conn.Execute "Delete from usuarios.vbloqbd where conta ='" & Empresa & "'"
        
    Else
        MsgBox "BD bloqueada.", vbCritical
        ComprobarEmpresaBloqueada = True
    End If
End If

Conn.Execute "commit"
End Function

