Attribute VB_Name = "ModFunciones"
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
'   En este modulo estan las funciones que recorren el form
'   usando el each for
'   Estas son
'
'   CamposSiguiente -> Nos devuelve el el text siguiente en
'           el orden del tabindex
'
'   CompForm -> Compara los valores con su tag
'
'   InsertarDesdeForm - > Crea el sql de insert e inserta
'
'   Limpiar -> Pone a "" todos los objetos text de un form
'
'   ObtenerBusqueda -> A partir de los text crea el sql a
'       partir del WHERE ( sin el).
'
'   ModifcarDesdeFormulario -> Opcion modificar. Genera el SQL
'
'   PonerDatosForma -> Pone los datos del RECORDSET en el form
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
Option Explicit
Public NombreCheck As String
Public Const ValorNulo = "Null"

Public Function CompForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox And Control.Visible = True Then
            Carga = mTag.Cargar(Control)
            If Carga = True Then
                Correcto = mTag.Comprobar(Control)
                If Not Correcto Then Exit Function
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function

                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                            Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function CompForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm2 = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                Carga = mTag.Cargar(Control)
                If Carga = True Then
                    Correcto = mTag.Comprobar(Control)
                    If Not Correcto Then Exit Function
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Carga = mTag.Cargar(Control)
                    If Carga = False Then
                        MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                        Exit Function
    
                    Else
                        If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                                MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                                Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    CompForm2 = True
End Function




'Public Function CampoSiguiente(ByRef formulario As Form, valor As Integer) As Control
'Dim Fin As Boolean
'Dim Control As Object
'
'On Error GoTo ECampoSiguiente
'
'    'Debug.Print "Llamada:  " & Valor
'    'Vemos cual es el siguiente
'    Do
'        valor = valor + 1
'        For Each Control In formulario.Controls
'            'Debug.Print "-> " & Control.Name & " - " & Control.TabIndex
'            'Si es texto monta esta parte de sql
'            If Control.TabIndex = valor Then
'                    Set CampoSiguiente = Control
'                    Fin = True
'                    Exit For
'            End If
'        Next Control
'        If Not Fin Then
'            valor = -1
'        End If
'    Loop Until Fin
'    Exit Function
'ECampoSiguiente:
'    Set CampoSiguiente = Nothing
'    Err.Clear
'End Function



'-----------------------------------

Public Function InsertarDesdeForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    cad = "1"
                    Else
                    cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                Der = Der & cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        cad = ValorNulo
                    Else
                        cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.Tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
    
    Conn.Execute cad, , adCmdText
    
    InsertarDesdeForm = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function InsertarDesdeForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm2 = False
    Der = ""
    Izda = ""
    
    For Each Control In formulario.Controls
    
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & ""
                        
                            'Parte VALUES
                            cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Izda <> "" Then Izda = Izda & ","
                    'Access
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.Value = 1 Then
                        cad = "1"
                        Else
                        cad = "0"
                    End If
                    If Der <> "" Then Der = Der & ","
                    If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                    Der = Der & cad
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                        If Control.ListIndex = -1 Then
                            cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & ","
                            Izda = Izda & "" & mTag.columna & ""
                            cad = Control.Index
                            If Der <> "" Then Der = Der & ","
                            Der = Der & cad
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
'                        If Control.Value Then
'                            If Izda <> "" Then Izda = Izda & ","
'                            Izda = Izda & "" & mTag.columna & ""
'                            cad = Control.index
'                            If Der <> "" Then Der = Der & ","
'                            Der = Der & cad
'                        End If
                        If Izda <> "" Then Izda = Izda & ","
                        Izda = Izda & "" & mTag.columna & ""
                        
                        'Parte VALUES
                        If Control.Visible Then
                            cad = ValorParaSQL(Control.Value, mTag)
                        Else
                            cad = ValorNulo
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        End If
        
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.Tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
    Conn.Execute cad, , adCmdText
    
     ' ### [Monica] 18/12/2006
    CadenaCambio = cad
   
    InsertarDesdeForm2 = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function


Public Function CadenaInsertarDesdeForm(ByRef formulario As Form) As String
'Equivale a InsertarDesdeForm, excepto que devuelve la candena SQL y hace el execute fuera de la función.
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    CadenaInsertarDesdeForm = ""
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    cad = "1"
                    Else
                    cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                Der = Der & cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        cad = ValorNulo
                    Else
                        cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.Tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
'    Conn.Execute cad, , adCmdText
    
    CadenaInsertarDesdeForm = cad
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function



Private Function ObtenerMaximoMinimo(vSQL As String, Optional vBD As Byte) As String
Dim RS As Recordset
    ObtenerMaximoMinimo = ""
    Set RS = New ADODB.Recordset
    If vBD = cConta Then
        RS.Open vSQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else
        RS.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    If Not RS.EOF Then
        If Not IsNull(RS.EOF) Then
            ObtenerMaximoMinimo = CStr(RS.Fields(0))
        End If
    End If
    RS.Close
    Set RS = Nothing
End Function


'====DAVID
'Public Function ObtenerBusqueda(ByRef formulario As Form) As String
'    Dim Control As Object
'    Dim Carga As Boolean
'    Dim mTag As CTag
'    Dim Aux As String
'    Dim cad As String
'    Dim SQL As String
'    Dim tabla As String
'    Dim RC As Byte
'
'    On Error GoTo EObtenerBusqueda
'
'    'Exit Function
'    Set mTag = New CTag
'    ObtenerBusqueda = ""
'    SQL = ""
'
'    'Recorremos los text en busca de ">>" o "<<"
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If Aux = ">>" Or Aux = "<<" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'                    If Aux = ">>" Then
'                        cad = " MAX(" & mTag.Columna & ")"
'                    Else
'                        cad = " MIN(" & mTag.Columna & ")"
'                    End If
'                    SQL = "Select " & cad & " from " & mTag.tabla
'                    SQL = ObtenerMaximoMinimo(SQL)
'                    Select Case mTag.TipoDato
'                    Case "N"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(SQL)
'                    Case "F"
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
'                    Case Else
'                        SQL = mTag.tabla & "." & mTag.Columna & " = '" & SQL & "'"
'                    End Select
'                    SQL = "(" & SQL & ")"
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los text en busca del NULL
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            Aux = Trim(Control.Text)
'            If UCase(Aux) = "NULL" Then
'                Carga = mTag.Cargar(Control)
'                If Carga Then
'
'                    SQL = mTag.tabla & "." & mTag.Columna & " is NULL"
'                    SQL = "(" & SQL & ")"
'                    Control.Text = ""
'                End If
'            End If
'        End If
'    Next
'
'
'
'    'Recorremos los textbox
'    For Each Control In formulario.Controls
'        If TypeOf Control Is TextBox Then
'            'Cargamos el tag
'            Carga = mTag.Cargar(Control)
'            If Carga Then
'                If mTag.Cargado Then
'                    Aux = Trim(Control.Text)
'                    If Aux <> "" Then
'                        If mTag.tabla <> "" Then
'                            tabla = mTag.tabla & "."
'                        Else
'                            tabla = ""
'                        End If
'                    RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, cad)
'                    If RC = 0 Then
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'            Else
'                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
'                Exit Function
'            End If
'
'        'COMBO BOX
'        ElseIf TypeOf Control Is ComboBox Then
'            mTag.Cargar Control
'            If mTag.Cargado Then
'                If Control.ListIndex > -1 Then
'                    If mTag.TipoDato <> "T" Then
'                        cad = Control.ItemData(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = " & cad
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    Else
'                        cad = Control.List(Control.ListIndex)
'                        cad = mTag.tabla & "." & mTag.Columna & " = '" & cad & "'"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'
'
'        'CHECK
'        ElseIf TypeOf Control Is CheckBox Then
'            If Control.Tag <> "" Then
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    If Control.Value = 1 Then
'                        cad = mTag.tabla & "." & mTag.Columna & " = 1"
'                        If SQL <> "" Then SQL = SQL & " AND "
'                        SQL = SQL & "(" & cad & ")"
'                    End If
'                End If
'            End If
'        End If
'
'
'    Next Control
'    ObtenerBusqueda = SQL
'Exit Function
'EObtenerBusqueda:
'    ObtenerBusqueda = ""
'    MuestraError Err.Number, "Obtener búsqueda. "
'End Function

Public Function ObtenerBusqueda(ByRef formulario As Form, Optional CHECK As String, Optional vBD As Byte, Optional cadwhere As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim cad As String
    Dim Sql As String
    Dim Tabla As String
    Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                If Control.Tag <> "" Then
                    Carga = mTag.Cargar(Control)
                    If Carga Then
                        If Aux = ">>" Then
                            cad = " MAX("
                        Else
                            cad = " MIN("
                        End If
                        'monica
                        Select Case mTag.TipoDato
                            Case "FHF"
                                cad = cad & "date(" & mTag.columna & "))"
                            Case "FHH"
                                cad = cad & "time(" & mTag.columna & "))"
                            Case Else
                                cad = cad & mTag.columna & ")"
                        End Select
                        
                        Sql = "Select " & cad & " from " & mTag.Tabla
                        If cadwhere <> "" Then Sql = Sql & " WHERE " & cadwhere
                        Sql = ObtenerMaximoMinimo(Sql, vBD)
                        Select Case mTag.TipoDato
                        Case "N"
                            Sql = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Case "F"
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case "FHF"
                            Sql = "date(" & mTag.Tabla & "." & mTag.columna & ") = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case "FHH"
                            Sql = "time(" & mTag.Tabla & "." & mTag.columna & ") = '" & Format(Sql, "hh:mm:ss") & "'"
                        Case Else
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Sql & "'"
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next

    

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                'Cargamos el tag
                Carga = mTag.Cargar(Control)
                If Carga Then
'                    Debug.Print Control.Tag
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.Tabla <> "" Then
                            Tabla = mTag.Tabla & "."
                            Else
                            Tabla = ""
                        End If
                        Rc = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.columna, Aux, cad)
                        If Rc = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex > -1 Then
                        If mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        cad = mTag.Tabla & "." & mTag.columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 15/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Aux = ""
                    If CHECK <> "" Then
                        Tabla = DBLet(Control.Index, "T")
                        If Tabla <> "" Then Tabla = "(" & Tabla & ")"
                        Tabla = Control.Name & Tabla & "|"
                        If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        cad = Control.Value
                        cad = mTag.Tabla & "." & mTag.columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function

'Añade: CESAR
'Para utilizar los campos con TAG dentro de un Frame
Public Function ObtenerBusqueda2(ByRef formulario As Form, Optional CHECK As String, Optional opcio As Integer, Optional nom_frame As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim cad As String
    Dim Sql As String
    Dim Tabla As String
    Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda2 = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Aux = ">>" Then
                            cad = " MAX(" & mTag.columna & ")"
                        Else
                            cad = " MIN(" & mTag.columna & ")"
                        End If
                        Sql = "Select " & cad & " from " & mTag.Tabla
                        Sql = ObtenerMaximoMinimo(Sql)
                        Select Case mTag.TipoDato
                        Case "N"
                            Sql = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Case "F"
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case Else
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Sql & "'"
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
          If Control.Tag <> "" Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.Tabla <> "" Then
                            Tabla = mTag.Tabla & "."
                            Else
                            Tabla = ""
                        End If
                        Rc = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.columna, Aux, cad)
                        If Rc = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then ' +-+- 12/05/05: canvi de Cèsar, no te sentit passar-li un control que no té TAG +-+-
                mTag.Cargar Control
                If mTag.Cargado Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Control.ListIndex > -1 Then
                            cad = Control.ItemData(Control.ListIndex)
                            cad = mTag.Tabla & "." & mTag.columna & " = " & cad
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            End If
            
         ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 27/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    ' añadido 12022007
                    Aux = ""
                    If CHECK <> "" Then
                        Tabla = DBLet(Control.Index, "T")
                        If Tabla <> "" Then Tabla = "(" & Tabla & ")"
                        Tabla = Control.Name & Tabla & "|"
                        If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        cad = Control.Value
                        cad = mTag.Tabla & "." & mTag.columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda2 = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda2 = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function


Public Function ModificaDesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                             cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"

                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.Visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.Visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                         cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
                    Else
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
'
'
'                   If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                   'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                   cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadwhere
    Conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function


Public Function ModificaDesdeFormulario2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                                 cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Control.Value = 1 Then
                        Aux = "TRUE"
                    Else
                        Aux = "FALSE"
                    End If
                    If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'Esta es para access
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.ListIndex = -1 Then
                            Aux = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Aux = Control.ItemData(Control.ListIndex)
                        Else
                            Aux = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                             cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                        End If
'
'
'                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                        'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            Aux = Control.Index
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                              If mTag.EsClave Then
                                  'Lo pondremos para el WHERE
                                   If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                                   cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
                              Else
                                  If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                  cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                              End If
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
'                        If Control.Value Then
                         If mTag.columna <> "" Then
'                            Aux = Control.index
                            If Control.Visible Then
                                Aux = ValorParaSQL(Control.Value, mTag)
                            Else
                                Aux = ValorNulo
                            End If
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                                cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadwhere
    Conn.Execute Aux, , adCmdText

    ' ### [Monica] 18/12/2006
    CadenaCambio = cadUPDATE

    ModificaDesdeFormulario2 = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar 2. " & Err.Description
End Function


Public Sub FormateaCampo(vTex As TextBox)
    Dim mTag As CTag
    Dim cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                cad = TransformaPuntosComas(vTex.Text)
                cad = Format(cad, mTag.Formato)
                vTex.Text = cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub


Public Function FormatoCampo(ByRef vTex As TextBox) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim cad As String
    
    On Error GoTo EFormatoCampo

    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        FormatoCampo = mTag.Formato
    End If
    
EFormatoCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'Añade: CESAR
'Para utilizalo en el arreglaGrid
Public Function FormatoCampo2(ByRef objec As Object) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim cad As String

    On Error GoTo EFormatoCampo2

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        FormatoCampo2 = mTag.Formato
    End If
    
EFormatoCampo2:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


Public Function TipoCamp(ByRef objec As Object) As String
Dim mTag As CTag
Dim cad As String

    On Error GoTo ETipoCamp

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        TipoCamp = mTag.TipoDato
    End If

ETipoCamp:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef cadena As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim Cont As Integer
Dim cad As String

    I = 0
    Cont = 1
    cad = ""
    Do
        J = I + 1
        I = InStr(J, cadena, "|")
        If I > 0 Then
            If Cont = Orden Then
                cad = Mid(cadena, J, I - J)
                I = Len(cadena) 'Para salir del bucle
                Else
                    Cont = Cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = cad
End Function


'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim I As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneral
'bol = vSesion.Nivel < 2

'Añadir, modificar y borrar deshabilitados si no nivel
With formulario
    For I = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(I).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(I).Tag)
            If J < vSesion.Nivel Then
                .Toolbar1.Buttons(I).Enabled = False
            End If
        End If
    Next I
End With

Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub


Public Sub PonerModoMenuGral(ByRef formulario As Form, activo As Boolean)
Dim I As Integer
'Dim j As Integer

On Error GoTo PonerModoMenuGral

'Añadir, modificar y borrar deshabilitados si no Modo
    With formulario
        For I = 1 To .Toolbar1.Buttons.Count
            Select Case .Toolbar1.Buttons(I).ToolTipText
                Case "Nuevo"
                    .Toolbar1.Buttons(I).Visible = Not .DeConsulta
                Case "Modificar", "Eliminar", "Imprimir"
                    .Toolbar1.Buttons(I).Visible = Not .DeConsulta
                    .Toolbar1.Buttons(I).Enabled = activo
'                Case "Modificar"
'                Case "Eliminar"
'                Case "Imprimir"
            End Select
        Next I
        
        
        'El menu Visible
        .mnModificar.Visible = Not .DeConsulta
        .mnEliminar.Visible = Not .DeConsulta
        'El menu activo
        .mnModificar.Enabled = activo
        .mnEliminar.Enabled = activo
    End With
    
    
Exit Sub
PonerModoMenuGral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub

Public Sub PonerOpcionesMenuGeneralNew(formulario As Form)
Dim Control As Object
Dim I As Integer
Dim J As Integer
'Dim bol As Boolean

On Error GoTo EPonerOpcionesMenuGeneralNew
'bol = vSesion.Nivel < 2
'Añadir, modificar y borrar deshabilitados si no nivel
    For Each Control In formulario.Controls
'        Debug.Print Control.Name
        
        If Mid(Control.Name, 1, 2) = "mn" And Mid(Control.Name, 1, 7) <> "mnBarra" _
           And Control.Name <> "mnOpciones" Then
            J = Val(Control.HelpContextID)
            If J < vSesion.Nivel And J <> 0 Then
                Control.Enabled = False
            End If
        End If
    Next Control

Exit Sub
EPonerOpcionesMenuGeneralNew:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



'Este modifica las claves prinipales y todo
'la sentenca del WHERE cod=1 and .. viene en claves
Public Function ModificaDesdeFormularioClaves(ByRef formulario As Form, Claves As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim cadUPDATE As String
Dim I As Integer

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    cadwhere = Claves
    'Construimos el SQL
    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadwhere
    Conn.Execute Aux, , adCmdText

ModificaDesdeFormularioClaves = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Function BLOQUEADesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox And Control.Visible = True Then
            If Control.Tag <> "" Then

                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                         cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
                    End If
                End If
            End If
        End If
    Next Control

    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.Tabla
        Aux = Aux & " WHERE " & cadwhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        Conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
    
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function



Public Function BloqueaRegistro(cadTABLA As String, cadwhere As String) As Boolean
Dim Aux As String

    On Error GoTo EBloqueaRegistro
        
    BloqueaRegistro = False
    Aux = "select * FROM " & cadTABLA
    Aux = Aux & " WHERE " & cadwhere & " FOR UPDATE"

    'Intenteamos bloquear
    PreparaBloquear
    Conn.Execute Aux, , adCmdText
    BloqueaRegistro = True

EBloqueaRegistro:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
End Function


Public Function BloqueaRegistroForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    Next Control

    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
'        Aux = "Insert into zBloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.tabla
        Aux = Aux & "',""" & AuxDef & """)"
        Conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If Conn.Errors.Count > 0 Then
            If Conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim Sql As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
'        SQL = "DELETE from zBloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.tabla & "'"
        Conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function


'====================== LAURA

Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function

Public Sub InsertarCambios(Tabla As String, ValorAnterior As String, numalbar As String)
Dim Sql As String
Dim Sql2 As String

    Sql = CadenaCambio

    Sql2 = "insert into cambios (codusu, fechacambio, tabla, numalbar, cadena, valoranterior) values ("
    Sql2 = Sql2 & DBSet(vSesion.Codusu, "N") & "," & DBSet(Now, "FH") & "," & DBSet(Tabla, "T") & ","
    Sql2 = Sql2 & DBSet(numalbar, "T") & ","
    Sql2 = Sql2 & DBSet(Sql, "T") & ","
    If ValorAnterior = ValorNulo Then
        Sql2 = Sql2 & ValorNulo & ")"
    Else
        Sql2 = Sql2 & DBSet(ValorAnterior, "T") & ")"
    End If

    Conn.Execute Sql2

End Sub
    
Public Sub CargarValoresAnteriores(formulario As Form, Optional opcio As Integer, Optional nom_frame As String)
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim cad As String
    Set mTag = New CTag

    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & " , "
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & " = "
                            'Parte VALUES
                            cad = ValorParaSQL(Control.Text, mTag)
                            Izda = Izda & cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Izda <> "" Then Izda = Izda & " , "
                    'Access
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & " = "
                    If Control.Value = 1 Then
                        cad = "1"
                        Else
                        cad = "0"
                    End If
                    If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                    Izda = Izda & cad
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & " , "
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & " = "
                        If Control.ListIndex = -1 Then
                            cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        Izda = Izda & cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & " , "
                            Izda = Izda & "" & mTag.columna & " = "
                            cad = Control.Index
                            Izda = Izda & cad
                        End If
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is DTPicker Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & " , "
                        Izda = Izda & "" & mTag.columna & " = "
                        
                        'Parte VALUES
                        If Control.Visible Then
                            cad = ValorParaSQL(Control.Value, mTag)
                        Else
                            cad = ValorNulo
                        End If
                        Izda = Izda & cad
                    End If
                End If
            End If
        End If
        
    Next Control

    ValorAnterior = Izda

End Sub

Public Function CalcularImporte(cantidad As String, Precio As String, Importe As String, tipo As String) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe = Cantidad * precio  Tipo <>"1"
'Cantidad = Importe / precio Tipo ="1"
Dim vImp As Currency
Dim vCan As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    cantidad = ComprobarCero(cantidad)
    Precio = ComprobarCero(Precio)
    Importe = ComprobarCero(Importe)
    
    If tipo <> "1" Then
       vImp = CCur(ImporteFormateado(cantidad)) * CCur(ImporteFormateado(Precio))
       vImp = Round2(vImp, 2)
       CalcularImporte = Format(vImp, "###,##0.00")
    Else
       vCan = CCur(ImporteFormateado(Importe)) / CCur(ImporteFormateado(Precio))
       vCan = Round2(vCan, 3)
       CalcularImporte = Format(vCan, "##,##0.000")
    End If
End Function

Public Sub CalcularImporteNue(ByRef cantidad As TextBox, ByRef Precio As TextBox, ByRef Importe As TextBox, tipo As Integer)
'Calcula el Importe de una linea de hcode facturas
Dim vImp As Currency
Dim vCan As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    cantidad = ComprobarCero(cantidad.Text)
    Precio = ComprobarCero(Precio.Text)
    Importe = ComprobarCero(Importe.Text)
    
    Select Case tipo
        Case 0 ' me han introducido la cantidad
            vImp = CCur(ImporteFormateado(cantidad.Text)) * CCur(ImporteFormateado(Precio.Text))
            vImp = Round2(vImp, 2)
            Importe.Text = Format(vImp, "###,##0.00")
        Case 1 ' me han introducido el precio
            vImp = CCur(ImporteFormateado(cantidad.Text)) * CCur(ImporteFormateado(Precio.Text))
            vImp = Round2(vImp, 2)
            Importe.Text = Format(vImp, "###,##0.00")
        Case 2 ' me han introducido el importe
            vCan = CCur(ImporteFormateado(Importe.Text)) / CCur(ImporteFormateado(Precio.Text))
            vCan = Round2(vCan, 3)
            cantidad.Text = Format(vCan, "##,##0.000")
    End Select
    
End Sub


'Public Function PonerNomEmple(codEmp As String) As String
'Dim nomEmp As String
'Dim cad As String
'
'    'apellidos i nombre del empleado
'    If (codEmp <> "") Then
'        nomEmp = "nomemple"
'        cad = DevuelveDesdeBDNew(cPTours, "empleado", "apeemple", "codemple", codEmp, "N", nomEmp, "codempre", CStr(vSesion.Empresa), "N", "codagenc", CStr(vSesion.Agencia), "N")
'        If cad <> "" Then cad = cad & ", " & nomEmp
'    End If
'    PonerNomEmple = cad
'End Function



Public Function ExisteCP(T As TextBox) As Boolean
'comprueba para un campo de texto que sea clave primaria, si ya existe un
'registro con ese valor
Dim vtag As CTag
Dim devuelve As String

    On Error GoTo ErrExiste

    ExisteCP = False
    If T.Text <> "" Then
        If T.Tag <> "" Then
            Set vtag = New CTag
            If vtag.Cargar(T) Then
'                If vtag.EsClave Then
                    devuelve = DevuelveDesdeBDNew(cPTours, vtag.Tabla, vtag.columna, vtag.columna, T.Text, vtag.TipoDato)
                    If devuelve <> "" Then
    '                    MsgBox "Ya existe un registro para " & vtag.Nombre & ": " & T.Text, vbExclamation
                        MsgBox "Ya existe el " & vtag.Nombre & ": " & T.Text, vbExclamation
                        ExisteCP = True
                        PonerFoco T
                    End If
'                End If
            End If
            Set vtag = Nothing
        End If
    End If
    Exit Function
    
ErrExiste:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar código.", Err.Description
End Function




Public Function TotalRegistros(vSQL As String) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim RS As ADODB.Recordset

    On Error Resume Next

    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalRegistros = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then TotalRegistros = RS.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function

Public Function TotalRegistrosConsulta(cadSQL) As Long
Dim cad As String
Dim RS As ADODB.Recordset

    On Error GoTo ErrTotReg
    cad = "SELECT count(*) FROM (" & cadSQL & ") x"
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RS.EOF Then
        TotalRegistrosConsulta = DBLet(RS.Fields(0).Value, "N")
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
ErrTotReg:
    MuestraError Err.Number, "", Err.Description
End Function

' del ariges de Laura

Public Function BloqueoManual(cadTABLA As String, cadwhere As String)
Dim Aux As String

On Error GoTo EBLOQ

    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "INSERT INTO zbloqueos(codusu,tabla,clave) VALUES(" & vSesion.Codigo & ",'" & cadTABLA
        Aux = Aux & "',""" & cadwhere & """)"
        Conn.Execute Aux
        BloqueoManual = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If Conn.Errors.Count > 0 Then
            If Conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
'        Else
'            MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
'    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueoManual(cadTABLA As String) As Boolean
Dim Sql As String
'Solo me interesa la tabla
On Error Resume Next

        Sql = "DELETE FROM zbloqueos WHERE codusu=" & vSesion.Codigo & " and tabla='" & cadTABLA & "'"
        Conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function

Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim ent As Integer
Dim cad As String

  ' Comprobaciones

  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un número."
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

Private Function RellenaABlancos(cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Space(longitud)
    If PorLaDerecha Then
        cad = cadena & cad
        RellenaABlancos = Left(cad, longitud)
    Else
        cad = cad & cadena
        RellenaABlancos = Right(cad, longitud)
    End If
    
End Function

Public Function QuitarCero(Valor As String) As String
    On Error Resume Next
    
    If Valor <> "" Then
        If CSng(Valor) = 0 Then
            QuitarCero = ""
        Else
            QuitarCero = Valor
        End If
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function PonerAlmacen(codAlm As String) As String
'Comprueba si existe el Almacen y lo pone en el Text
Dim devuelve As String
    
    On Error Resume Next

'    If codAlm = "" Then
'        MsgBox "Debe introducir el Almacen.", vbInformation
'    Else
'        devuelve = DevuelveDesdeBDNew(cptours, "salmpr", "codalmac", "codalmac", codAlm, "N")
'        If devuelve = "" Then
'            MsgBox "No existe el Almacen: " & Format(codAlm, "000"), vbInformation
'            PonerAlmacen = ""
'        Else
'            PonerAlmacen = Format(codAlm, "000")
'        End If
'    End If
    PonerAlmacen = 1

    If Err.Number <> 0 Then Err.Clear
End Function

Public Function CalcularDto(Importe As String, Dto As String) As String
'devuelve el Dto% del Importe
'Ej el 16% de 120 = 19.2
Dim vImp As Currency
Dim vDto As Currency
On Error Resume Next

    Importe = ComprobarCero(Importe)
    Dto = ComprobarCero(Dto)
    
    vImp = CCur(Importe)
    vDto = CCur(Dto)
    
    vImp = ((vImp * vDto) / 100)
    'vImp = Round(vImp, 2)
    
    CalcularDto = CStr(vImp)
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function CalcularImporteProv(cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte, ImpDto As String, Optional Bruto As String) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    cantidad = ComprobarCero(cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    If Bruto <> "" Then
        vImp = CCur(Bruto) - CCur(ImpDto)
    Else
        vImp = (CCur(cantidad) * CCur(vPre)) - CCur(ImpDto)
    End If
        
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    
    vImp = Round(vImp, 2)
    CalcularImporteProv = CStr(vImp)
End Function

Public Function ModificaDesdeFormulario1(ByRef formulario As Form, Opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario1
    ModificaDesdeFormulario1 = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        ElseIf TypeOf Control Is TextBox And Control.Visible = True Then
            If (Opcion = 1 And Control.Name = "Text1") Or (Opcion = 3 And Control.Name = "txtAux") Then
            If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                                 cadwhere = cadwhere & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.Visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.Visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        ElseIf TypeOf Control Is OptionButton And Control.Visible Then
            If Control.Enabled Then
                If Control.Value = True And Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        Aux = Control.Index
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadwhere
    Conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario1 = True
    Exit Function
    
EModificaDesdeFormulario1:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function



Public Sub BACKUP_Tabla(ByRef RS As ADODB.Recordset, ByRef Derecha As String, Optional canvi_nom As String, Optional canvi_valor As String)
Dim I As Integer
Dim nexo As String
Dim Valor As String
Dim tipo As Integer

    Derecha = ""
    nexo = ""
    For I = 0 To RS.Fields.Count - 1
        tipo = RS.Fields(I).Type
        
        If (canvi_nom <> "" And RS.Fields(I).Name = canvi_nom) Then
            Valor = canvi_valor
            If tipo = 133 Then
                Valor = "'" & Format(Valor, "yyyy-mm-dd") & "'"
            End If
        Else
            If tipo = 201 Then 'MEMO
                Valor = DBLetMemo(RS.Fields(I).Value)
                If Valor <> "" Then
                    NombreSQL Valor
                    Valor = "'" & Valor & "'"
                Else
                    Valor = "NULL"
                End If
            
            Else
                If IsNull(RS.Fields(I)) Then
                    Valor = "NULL"
                Else
                    'pruebas
                    Select Case tipo
                    'TEXTO
                    Case 129, 200
                        Valor = RS.Fields(I)
                        NombreSQL Valor
                        Valor = "'" & Valor & "'"
                    'Fecha
                    Case 133
                        Valor = CStr(RS.Fields(I))
                        Valor = "'" & Format(Valor, "yyyy-mm-dd") & "'"
                        
                    Case 134 'HORA
                        Valor = DBSet(Valor, "H")
                        
                    Case 135 'Fecha/Hora
                        Valor = DBSet(RS.Fields(I), "FH", "S")
                    'Numero normal, sin decimales
                    Case 2, 3, 16 To 19
                        Valor = RS.Fields(I)
                    
                    'Numero con decimales
                    Case 131, 6
                        Valor = CStr(RS.Fields(I))
                        Valor = TransformaComasPuntos(Valor)
                    Case Else
                        Valor = "Error grave. Tipo de datos no tratado." & vbCrLf
                        Valor = Valor & vbCrLf & "SQL: " & RS.Source
                        Valor = Valor & vbCrLf & "Pos: " & I
                        Valor = Valor & vbCrLf & "Campo: " & RS.Fields(I).Name
                        Valor = Valor & vbCrLf & "Valor: " & RS.Fields(I)
                        MsgBox Valor, vbExclamation
                        MsgBox "El programa finalizara. Avise al soporte técnico.", vbCritical
                        End
                    End Select
                End If
            End If
        End If
        Derecha = Derecha & nexo & Valor
        nexo = ","
    Next I
    Derecha = "(" & Derecha & ")"
End Sub

Public Function EsProveedorVarios(CodProve As String) As Boolean
Dim devuelve As String

    EsProveedorVarios = False
'    devuelve = DevuelveDesdeBD("provario", "proveedor", "codprove", codProve, "N")
'    If devuelve <> "" Then EsProveedorVarios = CBool(devuelve)
    'Es proveedor de varios Y podemos recuperar de ????
End Function

Public Function CalcularPorcentaje(Importe As Currency, Porce As Currency, NumDecimales As Long) As Variant
'devuelve el valor del Porcentaje aplicado al Importe
'Ej el 16% de 120 = 19.2
'Dim vImp As Currency
'Dim vDto As Currency
    
    On Error Resume Next
'
'    Importe = ComprobarCero(Importe)
'    Dto = ComprobarCero(Dto)
'
'    vImp = CCur(Importe)
'    vDto = CCur(Dto)
    
    
    'vImp = Round(vImp, 2)
    
    CalcularPorcentaje = Round2((Importe * Porce) / 100, NumDecimales)
    
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function DevuelveValor(vSQL As String) As Variant
'Devuelve el valor de la SQL
Dim RS As ADODB.Recordset

    On Error Resume Next

    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    DevuelveValor = 0
    If Not RS.EOF Then
        'antes RS.Fields(0).Value > 0
        If Not IsNull(RS.Fields(0).Value) Then DevuelveValor = RS.Fields(0).Value   'Solo es para saber que hay registros que mostrar
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        DevuelveValor = 0
        Err.Clear
    End If
End Function

Public Function CalcularImporte2(cantidad As String, Precio As String, Dto1 As String, Dto2 As String, TipoDto As Byte, ImpDto As String, Optional Bruto As String) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
On Error Resume Next

    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    cantidad = ComprobarCero(cantidad)
    vPre = ComprobarCero(Precio)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    If Bruto <> "" Then
        vImp = CCur(Bruto) - CCur(ImpDto)
    Else
        vImp = (CCur(cantidad) * CCur(vPre)) - CCur(ImpDto)
    End If
        
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    
    vImp = Round(vImp, 2)
    CalcularImporte2 = CStr(vImp)
End Function

'Añado Optional CHECK As String. Para poder realizar las busquedas con los checks
'monica corresponde al ObtenerBusqueda de laura
Public Function ObtenerBusqueda3(ByRef formulario As Form, paraRPT As Boolean, Optional CHECK As String) As String
Dim Control As Object
Dim Carga As Boolean
Dim mTag As CTag
Dim Aux As String
Dim cad As String
Dim Sql As String
Dim Tabla As String, columna As String
Dim Rc As Byte

    On Error GoTo EObtenerBusqueda3

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda3 = ""
    Sql = ""
    cad = ""
    
    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.Visible Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        If Not paraRPT Then
                            cad = " MAX(" & mTag.columna & ")"
                        Else
                            cad = " MAX({" & mTag.Tabla & "." & mTag.columna & "})"
                        End If
                    Else
                        If Not paraRPT Then
                            cad = " MIN(" & mTag.columna & ")"
                        Else
                            cad = " MIN({" & mTag.Tabla & "." & mTag.columna & "})"
                        End If
                    End If
                    If Not paraRPT Then
                        Sql = "Select " & cad & " from " & mTag.Tabla
                    Else
                        Sql = "Select " & cad & " from {" & mTag.Tabla & "}"
                    End If
                    Sql = ObtenerMaximoMinimo(Sql)
                    
                    Select Case mTag.TipoDato
                    Case "N"
                        If Not paraRPT Then
                            Sql = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(Sql)
                        Else
                            Sql = "{" & mTag.Tabla & "." & mTag.columna & "} = " & TransformaComasPuntos(Sql)
                        End If
                    Case "F"
                        If Not paraRPT Then
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Else
                            Sql = "{" & mTag.Tabla & "." & mTag.columna & "} = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        End If
                    Case Else
                        If Not paraRPT Then
                            Sql = mTag.Tabla & "." & mTag.columna & " = '" & Sql & "'"
                        Else
                            Sql = "{" & mTag.Tabla & "." & mTag.columna & "} = '" & Sql & "'"
                        End If
                    End Select
                    Sql = "(" & Sql & ")"
                End If
            End If
        End If
    Next

    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.Visible Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Not paraRPT Then
                        Sql = mTag.Tabla & "." & mTag.columna & " is NULL"
                    Else
                        Sql = "{" & mTag.Tabla & "." & mTag.columna & "} is NULL"
                    End If
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.Visible And Control.Name = "Text1" Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If mTag.Cargado Then
                    Aux = Trim(Control.Text)
                    Aux = QuitarCaracterEnter(Aux) 'Si es multilinea quitar ENTER
                    If Aux <> "" Then
                        If mTag.Tabla <> "" Then
                            If Not paraRPT Then
                                Tabla = mTag.Tabla & "."
                            Else
                                Tabla = "{" & mTag.Tabla & "."
                            End If
                        Else
                            Tabla = ""
                        End If
                        If Not paraRPT Then
                            columna = mTag.columna
                        Else
                            columna = mTag.columna & "}"
                        End If
                    Rc = SeparaCampoBusqueda3(mTag.TipoDato, Tabla & columna, Aux, cad, paraRPT)
                    If Rc = 0 Then
                        If Sql <> "" Then Sql = Sql & " AND "
                        If Not paraRPT Then
                            Sql = Sql & "(" & cad & ")"
                        Else
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If

        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible Then
            mTag.Cargar Control
            If mTag.Cargado Then
                If Control.ListIndex > -1 Then
                    If mTag.TipoDato <> "T" Then
                        cad = Control.ItemData(Control.ListIndex)
                        If Not paraRPT Then
                            cad = mTag.Tabla & "." & mTag.columna & " = " & cad
                        Else
                            cad = "{" & mTag.Tabla & "." & mTag.columna & "} = " & cad
                        End If
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    Else
                        cad = Control.List(Control.ListIndex)
                        If Not paraRPT Then
                            cad = mTag.Tabla & "." & mTag.columna & " = '" & cad & "'"
                        Else
                            cad = "{" & mTag.Tabla & "." & mTag.columna & "} = '" & cad & "'"
                        End If
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If


        'CHECK
                'CHECK
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    
                    Aux = ""
                    If CHECK <> "" Then
                        CheckBusqueda Control
                        Tabla = NombreCheck & "|"
                        If InStr(1, CHECK, Tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    
                    If Aux <> "" Then
                        If Not paraRPT Then
                            cad = mTag.Tabla & "." & mTag.columna
                        Else
                            cad = "{" & mTag.Tabla & "." & mTag.columna & "} "
                        End If
                        
                        cad = cad & " = " & Aux
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If 'cargado
                End If '<>""
            End If
        End If
    
    Next Control
    ObtenerBusqueda3 = Sql
Exit Function
EObtenerBusqueda3:
    ObtenerBusqueda3 = ""
    MuestraError Err.Number, "Obtener búsqueda. "
End Function

Public Function SeparaCampoBusqueda3(tipo As String, campo As String, cadena As String, ByRef DevSQL As String, Optional paraRPT) As Byte
Dim cad As String
Dim Aux As String
Dim CH As String
Dim Fin As Boolean
Dim I, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda3 = 1
DevSQL = ""
cad = ""
Select Case tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    '==== Laura: 11/07/05
    If IsNumeric(cadena) Then
        cadena = CStr(ImporteFormateado(cadena))
        cadena = TransformaComasPuntos(cadena)
    End If
    '====================
    I = CararacteresCorrectos(cadena, "N")
    If I > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    I = InStr(1, cadena, ":")
    If I > 0 Then
        'Intervalo numerico
        cad = Mid(cadena, 1, I - 1)
        Aux = Mid(cadena, I + 1)
        If Not IsNumeric(cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = campo & " >= " & cad & " AND " & campo & " <= " & Aux
        '----
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If cadena = ">>" Or cadena = "<<" Then
                DevSQL = "1=1"
             Else
                    Fin = False
                    I = 1
                    cad = ""
                    Aux = "NO ES NUMERO"
                    While Not Fin
                        CH = Mid(cadena, I, 1)
                        If CH = ">" Or CH = "<" Or CH = "=" Then
                            cad = cad & CH
                            Else
                                Aux = Mid(cadena, I)
                                Fin = True
                        End If
                        I = I + 1
                        If I > Len(cadena) Then Fin = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If cad = "" Then cad = " = "
                    DevSQL = campo & " " & cad & " " & Aux
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(cadena, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, cadena, ":")
    If I > 0 Then
        'Intervalo de fechas
        cad = Mid(cadena, 1, I - 1)
        Aux = Mid(cadena, I + 1)
        If Not EsFechaOK(cad) Or Not EsFechaOK(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        
'        If Not Left(campo, 1) = "{" Then
'                    Aux = "'" & Format(Aux, FormatoFecha) & "'"
'                Else
'                    Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
'                End If
        
        If paraRPT Then
            cad = "Date(" & Year(cad) & "," & Month(cad) & "," & Day(cad) & ")"
            Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
            DevSQL = campo & " >=" & cad & " AND " & campo & " <= " & Aux
        Else
            cad = Format(cad, FormatoFecha)
            Aux = Format(Aux, FormatoFecha)
            'En my sql es la ' no el #
            'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
            DevSQL = campo & " >='" & cad & "' AND " & campo & " <= '" & Aux & "'"
        End If
        '----
        'ELSE
    Else
            'Comprobamos que no es el mayor
            If cadena = ">>" Or cadena = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                I = 1
                cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    CH = Mid(cadena, I, 1)
                    If CH = ">" Or CH = "<" Or CH = "=" Then
                        cad = cad & CH
                        Else
                            Aux = Mid(cadena, I)
                            Fin = True
                    End If
                    I = I + 1
                    If I > Len(cadena) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOK(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                If Not Left(campo, 1) = "{" Then
                    Aux = "'" & Format(Aux, FormatoFecha) & "'"
                Else
                    Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
                End If
                If cad = "" Then cad = " = "
                DevSQL = campo & " " & cad & " " & Aux
            End If
    End If
    
    
Case "T"
    '---------------- TEXTO ------------------
    I = CararacteresCorrectos(cadena, "T")
    If I = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
     If cadena = ">>" Or cadena = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    
    'Comprobamos si es LIKE o NOT LIKE
    cad = Mid(cadena, 1, 2)
    If cad = "<>" Then
        cadena = Mid(cadena, 3)
        If Left(campo, 1) <> "{" Then
            'No es consulta seleccion para Report.
            DevSQL = campo & " NOT LIKE '"
        Else
            'Consulta de seleccion para Crystal Report
            DevSQL = "NOT (" & campo & " LIKE """ & cadena & """)"
        End If
    Else
        If Left(campo, 1) <> "{" Then
        'NO es para report
            DevSQL = campo & " LIKE '"
        Else  'Es para report
            I = InStr(1, cadena, "*")
            'Poner Consulta de seleccion para Crystal Report
            If I > 0 Then
                DevSQL = campo & " LIKE """ & cadena & """"
            Else
                DevSQL = campo & " = """ & cadena & """"
            End If
        End If
    End If
    
    
    'Cambiamos el * por % puesto que en ADO es el caraacter para like
    I = 1
    Aux = cadena
    If Not Left(campo, 1) = "{" Then
      'No es para report
       While I <> 0
           I = InStr(1, Aux, "*")
           If I > 0 Then
                Aux = Mid(Aux, 1, I - 1) & "%" & Mid(Aux, I + 1)
            End If
        Wend
    End If
    
    'Cambiamos el ? por la _ pue es su omonimo
    I = 1
    While I <> 0
        I = InStr(1, Aux, "?")
        If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "_" & Mid(Aux, I + 1)
    Wend
    
    
    'Poner el valor de la expresion
    If Left(campo, 1) <> "{" Then
        'No es consulta seleccion para Report.
        DevSQL = DevSQL & Aux & "'"
    'Else
        'Consulta de seleccion para Crystal Report
        'DevSQL = DevSQL & CADENA & """)"
    End If
    
    '=========
    'ANTES
'    If cad = "<>" Then
'        '====David
'        'Aux = Mid(CADENA, 3)
'        'LAura
'        Aux = Mid(Aux, 3)
'        '====
'        If Left(Campo, 1) <> "{" Then
'            'Mo es consulta seleccion para Report.
'            DevSQL = Campo & " NOT LIKE '" & Aux & "'"
'        Else
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " <> " & Aux & ""
'        End If
'    Else
'        If Left(Campo, 1) <> "{" Then
'            DevSQL = Campo & " LIKE '" & Aux & "'"
'        ElseIf Left(Aux, 4) = "like" Then
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " " & Aux
'        Else
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " = """ & Aux & """"
'        End If
'    End If
    
    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    I = InStr(1, cadena, "<>")
    If I = 0 Then
        'IGUAL A valor
        cad = " = "
        Else
            'Distinto a valor
        cad = " <> "
    End If
    'Verdadero o falso
    I = InStr(1, cadena, "V")
    If I > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = campo & " " & cad & " " & Aux
    
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda3 = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function

'---------------------------------------------------------------------------------
'
'       Para buscar en los checks con las dos opciones de true y false
'
'A partir de un check cualquiera devolvera nombre e indice, si tiene. Si no sera ()
Public Sub CheckBusqueda(ByRef CH As CheckBox)
    NombreCheck = ""
    NombreCheck = CH.Name & "("
    On Error Resume Next
    NombreCheck = NombreCheck & CH.Index
    If Err.Number <> 0 Then Err.Clear
    NombreCheck = NombreCheck & ")"
End Sub


