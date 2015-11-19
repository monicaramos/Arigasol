VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Variables locales que contienen valores de propiedad

Private mCodGuiav As Integer 'codigo guia
Private mNomGuiav As String 'nombre guia
Private mApe1Guia As String 'apellido 1 guia
Private mApe2Guia As String 'apellido 2 guia


'------------------------------------------------
'Propiedades del modulo CGuia
'------------------------------------------------

'codigo de cliente
Public Property Let CodGuia(ByVal vdata As Integer)
     mCodGuiav = vdata
End Property

Public Property Get CodGuia() As Integer
     CodGuia = mCodGuiav
End Property


'Nombre
Public Property Let NombreGuia(ByVal vdata As String)
     mNomGuiav = vdata
End Property

Public Property Get NombreGuia() As String
     NombreGuia = mNomGuiav
End Property


'Apellido 1
Public Property Let Apellido1Guia(ByVal vdata As String)
     mApe1Guia = vdata
End Property

Public Property Get Apellido1Guia() As String
     Apellido1Guia = mApe1Guia
End Property


'Apellido 2
Public Property Let Apellido2Guia(ByVal vdata As String)
     mApe2Guia = vdata
End Property

Public Property Get Apellido2Guia() As String
     Apellido2Guia = mApe2Guia
End Property


'****   LEER  ***
Public Function LeerDatos(codigo As String) As Boolean
'Leer los datos del guia de viaje
'Lee de la BD: PlannerTours, Tabla: guaviaj

    Dim RS As ADODB.Recordset
    Dim SQL As String, mens As String

    On Error GoTo ELeer
    
    SQL = "SELECT * FROM guiaviaj WHERE codguiav=" & codigo

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    mens = ""
    
    If RS.EOF Then
        LeerDatos = False
    Else
        mCodGuiav = RS!Codguiav

        mNomGuiav = DBLet(RS!NomGuiav, "T")
        mApe1Guia = DBLet(RS!Ape1Guia, "T")
        mApe2Guia = DBLet(RS!Ape2Guia, "T")
        
        mens = "Datos del guia del viaje"
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


'Public Function NombreCompleto() As String
'    NombreCompleto = NombreEmple & " " & ApellidoEmple
'End Function

