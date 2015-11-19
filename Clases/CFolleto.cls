VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFolleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' +-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: LAURA     +-+
' +-+- Fecha: 11/05/06  +-+
' +-+-+-+-+-+-+-+-+-+-+-+-+

'Folleto de viaje
'TABLA: follviaj

'Variables locales que contienen valores de propiedad
Private mCodEmpre As String 'Empresa
Private mCodFovia As String 'Cod. del folleto
Private mDesFovia As String 'Descripcion del folleto
Private mAnyFovia As Integer 'Año del folleto
Private mFecFovia As String 'Fecha creacion del folleto
Private mFecDesde As String 'Fecha desde de validez del folleto
Private mFecHasta As String 'Fecha hasta de validez del folleto




'------------------------------------------------
'Propiedades del modulo CFolleto
'------------------------------------------------

'Cod Empresa
Public Property Let codEmpre(ByVal vdata As String)
     mCodEmpre = vdata
End Property

Public Property Get codEmpre() As String
     codEmpre = mCodEmpre
End Property



'Cod. Folleto
Public Property Let codFovia(ByVal vdata As String)
     mCodFovia = vdata
End Property

Public Property Get codFovia() As String
     codFovia = mCodFovia
End Property


'Desc. Folleto
Public Property Let DesFovia(ByVal vdata As String)
     mDesFovia = vdata
End Property

Public Property Get DesFovia() As String
     DesFovia = mDesFovia
End Property



'Año del folleto
Public Property Let AnyFovia(ByVal vdata As Integer)
     mAnyFovia = vdata
End Property

Public Property Get AnyFovia() As Integer
     AnyFovia = mAnyFovia
End Property


'Fecha del folleto
Public Property Let FecFovia(ByVal vdata As String)
     mFecFovia = vdata
End Property

Public Property Get FecFovia() As String
     FecFovia = mFecFovia
End Property


'Fecha desde del folleto
Public Property Let FecDesde(ByVal vdata As String)
     mFecDesde = vdata
End Property

Public Property Get FecDesde() As String
     FecDesde = mFecDesde
End Property


'Fecha hasta del folleto
Public Property Let FecHasta(ByVal vdata As String)
     mFecHasta = vdata
End Property

Public Property Get FecHasta() As String
     FecHasta = mFecHasta
End Property


'------------------------------------------------
'Procedimientos del modulo CVenta
'------------------------------------------------


'****   LEER  ***
Public Function LeerDatos(codFovia As String, codEmpre As String) As Boolean
'Leer los datos del folleto de viajes
'Lee de la BD: PlannerTours, Tabla: follviaj

    Dim RS As ADODB.Recordset
    Dim SQL As String, mens As String

    On Error GoTo ELeer

    SQL = "SELECT * FROM follviaj WHERE codfovia=" & codFovia & " AND codempre=" & codEmpre

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    mens = ""

    If RS.EOF Then
        LeerDatos = False
    Else
        mCodEmpre = RS!codEmpre
        mCodFovia = RS!codFovia
        mDesFovia = DBLet(RS!DesFovia, "T")
        mAnyFovia = DBLet(RS!AnyFovia, "N") 'año folleto
        mFecFovia = DBLet(RS!FecFovia, "T")
        mFecDesde = DBLet(RS!FecDesde, "T")
        mFecHasta = DBLet(RS!FecHasta, "T")
        
        mens = "Datos del Folleto"
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    mens = "Se ha producido un error. " & mens & vbCrLf
    mens = mens & "Número: " & Err.Number & vbCrLf
    mens = mens & "Descripción: " & Err.Description
    MsgBox mens, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function



Public Function ValidezFolleto(cadFecha As String) As Boolean
'Comprueba si la fecha cadFecha esta dentro de la fecha de validez del folleto

    If cadFecha <> "" Then
        If EsFechaOK(cadFecha) Then
            If FecDesde <> "" And FecHasta <> "" Then
                If (CDate(FecDesde) <= CDate(cadFecha)) And (CDate(cadFecha) <= CDate(FecHasta)) Then
                    ValidezFolleto = True
                End If
            ElseIf FecDesde <> "" Then
                If (CDate(FecDesde) <= CDate(cadFecha)) Then
                    ValidezFolleto = True
                End If
            ElseIf FecHasta <> "" Then
                If (CDate(cadFecha) <= CDate(FecHasta)) Then
                    ValidezFolleto = True
                End If
            Else
                'Campos fecdesde y fechasta son nulos
                ValidezFolleto = True
            End If
        End If
    Else
        'si no se ha introducido Fecha comprobar que la fecha
        'del folleto ????
         ValidezFolleto = True '¿????
    End If

End Function



