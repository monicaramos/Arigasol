VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' +-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: LAURA     +-+
' +-+- Fecha: 20/02/06  +-+
' +-+-+-+-+-+-+-+-+-+-+-+-+

'TABLA: Presupue

'Variables locales que contienen valores de propiedad
Private mNumVenta As String 'nº de venta
Private mCodEmpre As Integer 'empresa
Private mNumPresu As Integer 'presupuesto

Private mAnyPresu As Integer 'año del presupuesto
Private mEstaPres As Byte 'Estado del presupuesto
                          '1=Pendiente, 2=Realizado, 3=Entregado, 4=Aceptado
                          
Private mFechReal As String 'fecha estado pasa a Realizado
Private mFechEntr As String 'fecha estado pasa a Entregado
Private mFechAcep As String 'fecha estado pasa a Aceptado


Private mReferenc As String 'referencia
Private mFechaSal As String 'fecha de salida
Private mFechaReg As String 'fecha de regreso
Private mNumeDias As Integer 'nº de dias
Private mNumeNoch As Integer 'nº de noches

Private mcodDesti As Integer 'cod. destino
Private mCodRegim As Integer 'cod. régimen/servicio
Private mcodProdu As Integer 'cod. producto
Private mCodFovia As Long 'cod. folleto viaje
Private mcodItine As Long 'cod. itinerario


Private mLugaSali As String 'lugar de salida
Private mHoraSali As String 'hora de salida
Private mPoblaSal As String 'poblacion de salida
Private mItinerar As String 'itinerario

                                         


'------------------------------------------------
'Propiedades del modulo CPresupuesto
'------------------------------------------------

'Nº de venta
Public Property Let NumVenta(ByVal vData As String)
     mNumVenta = vData
End Property

Public Property Get NumVenta() As String
     NumVenta = mNumVenta
End Property


'codigo empresa
Public Property Let CodEmpresa(ByVal vData As Integer)
     mCodEmpre = vData
End Property

Public Property Get CodEmpresa() As Integer
     CodEmpresa = mCodEmpre
End Property



'Nº de presupuesto
Public Property Let numPresu(ByVal vData As Integer)
     mNumPresu = vData
End Property

Public Property Get numPresu() As Integer
     numPresu = mNumPresu
End Property



'Anyo del presupuesto
Public Property Let AnyPresu(ByVal vData As Integer)
     mAnyPresu = vData
End Property

Public Property Get AnyPresu() As Integer
     AnyPresu = mAnyPresu
End Property



'Estado del presupuesto
Public Property Let EstadoPresu(ByVal vData As Byte)
     mEstaPres = vData
End Property

Public Property Get EstadoPresu() As Byte
     EstadoPresu = mEstaPres
End Property


'Fecha cuando presupuesto pasa a Realizado
Public Property Let FechaReal(ByVal vData As String)
     mFechReal = vData
End Property

Public Property Get FechaReal() As String
     FechaReal = mFechReal
End Property


'Fecha cuando  presupuesto pasa a Entregado
Public Property Let FechaEntre(ByVal vData As String)
     mFechEntr = vData
End Property

Public Property Get FechaEntre() As String
     FechaEntre = mFechEntr
End Property


'Fecha cuando  presupuesto pasa a Aceptado
Public Property Let FechaAcept(ByVal vData As String)
     mFechAcep = vData
End Property

Public Property Get FechaAcept() As String
     FechaAcept = mFechAcep
End Property




'REferencia del presupuesto
Public Property Let ReferPresu(ByVal vData As String)
     mReferenc = vData
End Property

Public Property Get ReferPresu() As String
     ReferPresu = mReferenc
End Property


'Fecha de salida
Public Property Let FechaSal(ByVal vData As String)
     mFechaSal = vData
End Property

Public Property Get FechaSal() As String
     FechaSal = mFechaSal
End Property


'Fecha de regreso
Public Property Let FechaReg(ByVal vData As String)
     mFechaReg = vData
End Property

Public Property Get FechaReg() As String
     FechaReg = mFechaReg
End Property


'Nº de dias
Public Property Let NumeDias(ByVal vData As Integer)
     mNumeDias = vData
End Property

Public Property Get NumeDias() As Integer
     NumeDias = mNumeDias
End Property



'Nº de noches
Public Property Let NumeNoches(ByVal vData As Integer)
     mNumeNoch = vData
End Property

Public Property Get NumeNoches() As Integer
     NumeNoches = mNumeNoch
End Property


'Cod. destino
Public Property Let CodDesti(ByVal vData As Integer)
     mcodDesti = vData
End Property

Public Property Get CodDesti() As Integer
     CodDesti = mcodDesti
End Property


'Cod. Régimen/Servicio
Public Property Let CodRegim(ByVal vData As Integer)
     mCodRegim = vData
End Property

Public Property Get CodRegim() As Integer
     CodRegim = mCodRegim
End Property



'Cod. Producto
Public Property Let CodProdu(ByVal vData As Integer)
     mcodProdu = vData
End Property

Public Property Get CodProdu() As Integer
     CodProdu = mcodProdu
End Property


'Cod. Folleto viaje
Public Property Let codFovia(ByVal vData As Long)
     mCodFovia = vData
End Property

Public Property Get codFovia() As Long
     codFovia = mCodFovia
End Property


'Cod. Itinerario
Public Property Let CodItine(ByVal vData As Long)
     mcodItine = vData
End Property

Public Property Get CodItine() As Long
     CodItine = mcodItine
End Property



'Lugar de salida
Public Property Let LugarSal(ByVal vData As String)
     mLugaSali = vData
End Property

Public Property Get LugarSal() As String
     LugarSal = mLugaSali
End Property


'Hora de salida
Public Property Let HoraSal(ByVal vData As String)
     mHoraSali = vData
End Property

Public Property Get HoraSal() As String
     HoraSal = mHoraSali
End Property


'Poblacion de salida
Public Property Let PoblaSal(ByVal vData As String)
     mPoblaSal = vData
End Property

Public Property Get PoblaSal() As String
     PoblaSal = mPoblaSal
End Property



'Itinerario
Public Property Let Itinerario(vData As String)
     mItinerar = vData
End Property

Public Property Get Itinerario() As String
     Itinerario = mItinerar
End Property





'------------------------------------------------
'Procedimientos del modulo CPresupuesto
'------------------------------------------------


'****   LEER  ***
Public Function LeerDatos(codigo1 As String, codigo2 As String, codigo3 As String) As Boolean
'Leer los datos del presupuesto
'Lee de la BD: PlannerTours, Tabla: Presupue

    Dim RS As ADODB.Recordset
    Dim sql As String, Mens As String

    On Error GoTo ELeer
    
    sql = "SELECT * FROM presupue WHERE numventa=" & codigo1 & " AND codempre=" & codigo2 & " AND numpresu=" & codigo3

    Set RS = New ADODB.Recordset
    RS.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If RS.EOF Then
        LeerDatos = False
    Else
        mNumVenta = RS!NumVenta
        mCodEmpre = RS!codEmpre 'codigo empresa
        mNumPresu = RS!numPresu 'nº presupuesto
        
        mAnyPresu = RS!AnyPresu 'anyo presupuesto
        mEstaPres = RS!estapres 'estado presupuesto
        
        mFechReal = DBLet(RS!FechReal, "F")
        mFechEntr = DBLet(RS!fechentr, "F")
        mFechAcep = DBLet(RS!FechAcep, "F")
        
        
        mReferenc = DBLet(RS!Referenc, "T")
        mFechaSal = DBLet(RS!FechaSal, "F")
        mFechaReg = DBLet(RS!FechaReg, "F")
        mNumeDias = DBLet(RS!NumeDias, "N")
        mNumeNoch = DBLet(RS!Numenoch, "N")
        
        mcodDesti = DBLet(RS!CodDesti, "N")
        mCodRegim = DBLet(RS!CodRegim, "N")
        mcodProdu = DBLet(RS!CodProdu, "N")
        mCodFovia = DBLet(RS!codFovia, "N")
        mcodItine = DBLet(RS!CodItine, "N")
        
        mLugaSali = DBLet(RS!LugaSali, "T")
        mHoraSali = DBLet(RS!HoraSali, "T")
        mPoblaSal = DBLet(RS!PoblaSal, "T")
        
        mItinerar = DBLetMemo(RS!Itinerar)
                
        Mens = "Datos del presupuesto"
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function



Public Function ExisteOtroPresuAceptado() As Boolean
'Solo puede haber un presupuesto en estado de aceptado para la venta
Dim sql As String

    sql = "SELECT count(*) FROM presupue "
    sql = sql & " WHERE numventa=" & NumVenta & " AND codempre= " & CodEmpresa
    sql = sql & " AND estapres = 4 AND numpresu <> " & numPresu
    
    If RegistrosAListar(sql) > 0 Then
        MsgBox "Ya existe un presupuesto aceptado para la venta. Seleccione otro estado.", vbExclamation
        ExisteOtroPresuAceptado = True
    Else
        ExisteOtroPresuAceptado = False
    End If
End Function





Public Function ActualizarEstado(nEstado As Byte) As Boolean
'(IN) nEstado: nuevo estado al que va a pasar el presupuesto
Dim b As Boolean
Dim sql As String
Dim DentroTrans As Boolean

    On Error GoTo EActEst
    
    'Comprobar que se pasa a un estado correcto
    'se pasa siempre a un estado posterior, e.d.:
    '- si estado es 1=Pendiente, podrá pasar a 2=Realizado, 3=Entregado, o 4=Aceptado.
    '- si estado es 2=Realizado, podrá pasar a 3=Entregado o 4=Aceptado.
    If mEstaPres < nEstado Then
    
        'si vamos a pasar el presupuesto a estado 4=Aceptado
        If nEstado = 4 Then
            'Comprobar que no existe ya otro presupuesto aceptado para la venta
           If ExisteOtroPresuAceptado Then Exit Function
        End If
    
    
'        If Not DentroTransac Then Conn.BeginTrans
        DentroTrans = True
        conn.BeginTrans
    
        'Actualizar estado del presupuesto
        sql = "UPDATE presupue SET estapres=" & nEstado
        
        'Actualizar tambien fechas de los estados segun el estado anterio y el nuevo
        'estado al que pasa el presupuesto
        Select Case mEstaPres
            'segun el estado anterior
            Case 1 'Pendiente
                'si estaba pendiente y pasa a realizado,entregado,acept actualizar también
                'la fecha de realizado, entreg., acept. ( y las anteriores)
                Select Case nEstado
                    Case 2 'Realizado
                        sql = sql & ", fechreal=" & DBSet(Now, "F")
                    Case 3 'Entregado
                        sql = sql & ", fechreal=" & DBSet(Now, "F") & ", fechentr=" & DBSet(Now, "F")
                    Case 4 'Aceptado
                        sql = sql & ", fechreal=" & DBSet(Now, "F") & ", fechentr=" & DBSet(Now, "F") & ", fechacep=" & DBSet(Now, "F")
                End Select
                
            Case 2 'Realizado
                Select Case nEstado
                    Case 3 'Entregado
                        sql = sql & ", fechentr=" & DBSet(Now, "F")
                    Case 4 'Aceptado
                        sql = sql & ", fechentr=" & DBSet(Now, "F") & ", fechacep=" & DBSet(Now, "F")
                End Select
                
            Case 3 'Entregado
                If nEstado = 4 Then
                    sql = sql & ", fechacep=" & DBSet(Now, "F")
                End If
'            Case 4 'Aceptado
        End Select
        
        sql = sql & " WHERE numventa=" & NumVenta & " AND codempre=" & CodEmpresa & " AND numpresu=" & numPresu
        conn.Execute sql
        b = True
        
        
        
        'Si pasa a ser un presupuesto aceptado, crear directamente un expediente
        'para dicha venta
        If nEstado = 4 Then
'            If SolicitaCrearExpte Then b = GenerarExpteDePresu
        End If
        

    Else
        MsgBox "No se puede pasar a un estado anterior.", vbExclamation
    End If
    
    
EActEst:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Actualizar estado presupuesto.", Err.Description
    End If
    If DentroTrans Then
        If b Then
            conn.CommitTrans
        Else
            conn.RollbackTrans
        End If
    End If
    ActualizarEstado = b
End Function


'Private Function GenerarExpteDePresu() As Boolean
''A partir de un presupuesto aceptado de una venta generamos el expediente
''correspondiente a dicha venta
'Dim cVen1 As CVenta
'
'    'volvemos a cargar los datos del presupuesto ya que al actualizar el esado se
'    'han cambiado campos de valor
'    If Me.LeerDatos(NumVenta, CodEmpresa, numPresu) Then
'
'        Set cVen1 = New CVenta
'        'leemos los datos de la venta del presupuesto aceptado
'        If cVen1.LeerDatos(NumVenta, CodEmpresa) Then
'
'            GenerarExpteDePresu = CrearExpedienteVenta(cVen1, Me)
'        End If
'        Set cVen1 = Nothing
'    End If
'End Function
