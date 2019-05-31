VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIppal 
   BackColor       =   &H8000000C&
   Caption         =   "AriGasol"
   ClientHeight    =   7860
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   11160
   Icon            =   "MDIppal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clientes"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Artículos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Traspaso Postes"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albaranes"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cuadre Diario"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturación"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Contabilización"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Histórico Facturas"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Impresión Tarjetas"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnParametros 
      Caption         =   "&Datos Básicos"
      Index           =   1
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Datos de Empresa"
         Index           =   1
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Parámetros"
         Index           =   2
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Tipos de Movimiento"
         Index           =   3
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Tipos de Documentos"
         Index           =   4
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Tipo de Crédito "
         Index           =   5
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Usuarios"
         Index           =   6
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Colectivos"
         Index           =   8
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "Clientes"
         Index           =   9
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Departamentos"
         Index           =   10
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Empleados"
         Index           =   11
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Familias"
         Index           =   12
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Artículos"
         Index           =   13
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Situaciones"
         Index           =   14
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "F&ormas de pago"
         Index           =   15
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "Bancos &propios"
         Index           =   16
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Grupo de Empresas"
         Index           =   17
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Proveedores"
         Index           =   18
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "Entidades Domiciliarias"
         Enabled         =   0   'False
         Index           =   19
         Visible         =   0   'False
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "Cambio Empresa"
         Index           =   21
      End
      Begin VB.Menu mnP_Generales 
         Caption         =   "&Salir"
         Index           =   22
      End
   End
   Begin VB.Menu mnGeneral 
      Caption         =   "&Ventas Diarias"
      Begin VB.Menu mnG_Ventas 
         Caption         =   "Traspaso datos &Postes"
         Index           =   1
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "&Albaranes"
         Index           =   3
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "&Resumen Ventas Articulos"
         Index           =   4
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "&Informe Prefacturación"
         Index           =   5
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "&Buscar errores Albaranes"
         Index           =   6
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "Comprobación &descuadres"
         Index           =   7
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "Cambio de Cliente"
         Index           =   8
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "&Estadística por Forma de Pago"
         Index           =   9
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "&Cuadre diario"
         Index           =   11
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnG_Ventas 
         Caption         =   "Contabilizar Cierre &Turno"
         Index           =   13
      End
   End
   Begin VB.Menu mnFacturacion 
      Caption         =   "&Facturación"
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "&Traspaso Facturas Tpv"
         Index           =   1
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "&Prefacturación Bonificación "
         Index           =   2
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "&Facturación"
         Index           =   3
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "&Reimpresión de Facturas"
         Index           =   4
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "&Enviar Facturas por email"
         Index           =   5
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "Facturación Web/Electrónica"
         Index           =   6
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "&Contabilizar Facturación"
         Index           =   8
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "&Localizador Tickets"
         Index           =   10
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "Facturas de &Abonos a Clientes"
         Index           =   12
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "F&acturas Rectificativas"
         Index           =   13
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "&Grabación Modelo 544 Gasoleo B"
         Index           =   14
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "&Ayuda Modelo 569"
         Index           =   15
      End
      Begin VB.Menu mnF_Facturacion 
         Caption         =   "Céntimo Sanitario"
         Index           =   16
      End
   End
   Begin VB.Menu mnFacturacionAjena 
      Caption         =   "Facturación &Ajena"
      Index           =   1
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "&Facturación"
         Index           =   1
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "&Reimpresión de Facturas"
         Index           =   2
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "&Histórico Facturas"
         Index           =   4
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "&Contabilización en Tesorería"
         Index           =   6
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "&Ventas por Cliente"
         Index           =   8
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "Ventas &Artículos por Cliente"
         Index           =   9
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "Facturas de &Abonos a Socios"
         Index           =   10
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "&Traspaso Datos Cooperativas"
         Index           =   12
      End
      Begin VB.Menu mnF_FacturacionAjena 
         Caption         =   "&Céntimo Sanitario"
         Index           =   13
      End
   End
   Begin VB.Menu mnEstadisticas 
      Caption         =   "&Estadísticas"
      Begin VB.Menu mnE_Estadist 
         Caption         =   "&Histórico Facturas"
         Index           =   1
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "&Diario de Facturación"
         Index           =   2
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "&Ventas por Cliente"
         Index           =   3
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "Ventas &Artículos por Cliente"
         Index           =   4
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "&Resumen Ventas Artículos"
         Index           =   5
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "Resumen Ventas Diarias"
         Index           =   6
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "Ventas Artículos por &Tarjeta"
         Index           =   7
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "Resumen por Rangos Horarios"
         Index           =   8
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "Declaraciones Gasóleo Profesional"
         Index           =   9
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "&Evolución Mensual Clientes"
         Index           =   10
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "&Certificado de Gasóleo B"
         Index           =   11
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "&Declaración Gasóleo B Hacienda"
         Index           =   12
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "Consumo entre &Fechas"
         Index           =   13
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "Movimientos de Artículos por Familia"
         Index           =   14
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "Margen Ventas por Artículos "
         Index           =   15
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "Margen Ventas por Cliente"
         Index           =   16
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "Traspaso a Histórico de Facturas1"
         Index           =   18
      End
      Begin VB.Menu mnE_Estadist 
         Caption         =   "Histórico de Facturas1"
         Index           =   19
      End
   End
   Begin VB.Menu mnTanques 
      Caption         =   "&Tanques-Mangueras"
      Begin VB.Menu mnE_Tanques 
         Caption         =   "Datos &Tanques/Mangueras"
         Index           =   1
      End
      Begin VB.Menu mnE_Tanques 
         Caption         =   "Datos &Recaudación"
         Index           =   2
      End
      Begin VB.Menu mnE_Tanques 
         Caption         =   "&Impresión Cierre de Turno"
         Index           =   3
      End
      Begin VB.Menu mnE_Tanques 
         Caption         =   "&Estadística de Artículos"
         Index           =   4
      End
      Begin VB.Menu mnE_Tanques 
         Caption         =   "Estadística &Artículos de Tanques"
         Index           =   5
      End
   End
   Begin VB.Menu mnCompras 
      Caption         =   "&Compras"
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "&Mant. Albaranes Proveedor"
         Index           =   1
      End
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "&Histórico Albaranes Anulados"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "&List. Pendiente de facturar"
         Index           =   3
      End
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "&Recepción Facturas"
         Index           =   5
      End
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "&Histórico Albarán/Factura"
         Index           =   6
      End
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "&Contabilizar Facturas"
         Index           =   8
      End
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "&Estadísticas"
         Index           =   10
         Begin VB.Menu mnCom_Est 
            Caption         =   "&Compras por Proveedor"
            Index           =   1
         End
         Begin VB.Menu mnCom_Est 
            Caption         =   "Compras por &Familia/Artículo"
            Index           =   2
         End
         Begin VB.Menu mnCom_Est 
            Caption         =   "&Albaranes por Proveedor"
            Index           =   3
         End
      End
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnCom_AlbCom 
         Caption         =   "&Inventario"
         Index           =   12
         Begin VB.Menu mnCom_Inven 
            Caption         =   "&Toma de Inventario"
            Index           =   1
         End
         Begin VB.Menu mnCom_Inven 
            Caption         =   "&Entrada Existencia Real"
            Index           =   2
         End
         Begin VB.Menu mnCom_Inven 
            Caption         =   "&Listado de Diferencias"
            Index           =   3
         End
         Begin VB.Menu mnCom_Inven 
            Caption         =   "&Actualizar Diferencias"
            Index           =   4
         End
         Begin VB.Menu mnCom_Inven 
            Caption         =   "&Valoración Stocks Inventariados"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnUtil 
      Caption         =   "&Utilidades"
      WindowList      =   -1  'True
      Begin VB.Menu mnE_Util 
         Caption         =   "Rangos &Horarios"
         Index           =   1
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "&Impresión de Tarjetas"
         Index           =   2
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Cambios &Registros"
         Index           =   3
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Declaración Gasóleo Profesional"
         Index           =   5
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Comprobación &Impuesto en Facturas"
         Index           =   7
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Revisión de caracteres en Multibase"
         Index           =   9
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Deshacer &Facturación"
         Index           =   11
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "&Copia de Seguridad local"
         Index           =   12
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "&Exportación TPV"
         Index           =   13
      End
      Begin VB.Menu mnE_Util 
         Caption         =   "Asignacion de codigos cliente"
         Index           =   14
      End
   End
   Begin VB.Menu mnSoporte 
      Caption         =   "&Soporte"
      Begin VB.Menu mnE_Soporte1 
         Caption         =   "&Web Soporte"
      End
      Begin VB.Menu mnp_Barra2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnE_Soporte2 
         Caption         =   "&Acerca de"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "MDIppal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private PrimeraVez As Boolean
Dim TieneEditorDeMenus As Boolean

Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim I As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub


Private Sub MDIForm_Activate()
'Dim cad As String

    If PrimeraVez Then
        PrimeraVez = False
'        frmMensaje.pTitulo = "Últimas modificaciones.         14/03/06"
'
''        cad = cad & "-----------------------------------------------------------------------------------------------" & vbCrLf
''        cad = cad & "Para actualizar el estado de un presupuesto desde la pantalla de ventas sin "
''        cad = cad & "entrar en la pantalla de modificación de presupuesto, seleccionar la línea del "
''        cad = cad & "presupuesto a modificar, pulsar botón izquierdo del ratón, se despliega un menu "
''        cad = cad & "con los posibles estados y se selecciona el nuevo estado." & vbCrLf & vbCrLf
'
''        cad = cad & "- Imprimir informes de subcontratación." & vbCrLf
''        cad = cad & "- Ventas pendientes." & vbCrLf
''        cad = cad & "-------------------------------------------------------------------------" & vbCrLf & vbCrLf
'
'        cad = cad & "- Mantenimiento de No Conformidades y lineas de acciones y reclamaciones." & vbCrLf & vbCrLf
'        cad = cad & "- Informes:" & vbCrLf
'        cad = cad & "     Comunicación con cliente." & vbCrLf
'        cad = cad & "     Confirmación de servicios." & vbCrLf
'        cad = cad & "     No conformidad." & vbCrLf
'        cad = cad & "     Reclamación." & vbCrLf & vbCrLf
'
'
'        frmMensaje.pValor = cad
'        frmMensaje.Show vbModal
    End If
End Sub

Private Sub MDIForm_Load()
Dim Cad As String

    PrimeraVez = True
    CargarImagen
    PonerDatosFormulario

    If vEmpresa Is Nothing Then
        Caption = "AriGasol" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & " FALTA CONFIGURAR"
    Else
        Caption = "AriGasol" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre & Cad & _
                  "   -  Usuario: " & vSesion.Nombre
    End If

    ' *** per als iconos XP ***
    GetIconsFromLibrary App.path & "\iconos.dll", 1, 32
    
    GetIconsFromLibrary App.path & "\iconos.dll", 1, 24
    GetIconsFromLibrary App.path & "\iconos_BN.dll", 2, 24
    GetIconsFromLibrary App.path & "\iconos_OM.dll", 3, 24
    
    GetIconsFromLibrary App.path & "\iconosArigasol.dll", 4, 24
    
    
  
    'CARGAR LA TOOLBAR DEL FORM PRINCIPAL
    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListPpal

        .Buttons(1).Image = 2   'Clientes
        .Buttons(2).Image = 7   'Articulos
        'el 3 son separadors
        .Buttons(4).Image = 3   'Traspaso postes
        .Buttons(5).Image = 10    'Albaranes
        .Buttons(6).Image = 8    'Cuadre diario
        'el 7 son separadors
        .Buttons(8).Image = 9   'Facturacion
        .Buttons(9).Image = 4   'Contabilizacion
        'el 10 son separadors
        .Buttons(11).Image = 5   'Historico de Facturas
        .Buttons(13).Image = 11  'Impresion de Tarjetas
'        .Buttons(7).Image = 1   'Salir
    End With
    
    
    GetIconsFromLibrary App.path & "\iconos.dll", 1, 16
    GetIconsFromLibrary App.path & "\iconos_BN.dll", 2, 16
    GetIconsFromLibrary App.path & "\iconos_OM.dll", 3, 16

    LeerEditorMenus

    PonerDatosFormulario
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    AccionesCerrar
    End
End Sub

Private Sub mnCom_AlbCom_Click(Index As Integer)
    SubmnC_Compras_Click (Index)
End Sub

Private Sub mnCom_Est_Click(Index As Integer)
    SubmnE_EstComp_Click (Index)
End Sub

Private Sub mnCom_Inven_Click(Index As Integer)
    SubmnC_ComprasInven_Click (Index)
End Sub

Private Sub mnE_Estadist_Click(Index As Integer)
    SubmnE_Estadist_Click (Index)
End Sub

Private Sub mnE_Soporte1_Click()
    Screen.MousePointer = vbHourglass
    LanzaHome "websoporte"
    espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnE_Util_Click(Index As Integer)
    SubmnE_Util_Click (Index)
End Sub

Private Sub mnE_Tanques_Click(Index As Integer)
    SubmnE_Tanques_Click (Index)
End Sub

Private Sub mnE_Soporte2_Click()
    frmMensaje.OpcionMensaje = 6
    frmMensaje.Show vbModal
End Sub

Private Sub mnF_FacturacionAjena_Click(Index As Integer)
    SubmnF_FacturacionAjena_Click (Index)
End Sub

Private Sub mnP_Generales_Click(Index As Integer)
    SubmnP_Generales_Click (Index)
End Sub

Private Sub mnG_Ventas_Click(Index As Integer)
    SubmnG_Ventas_Click (Index)
End Sub

Private Sub mnF_Facturacion_Click(Index As Integer)
    SubmnF_Facturacion_Click (Index)
End Sub

Private Sub mnP_Salir1_Click()
'    Unload frmPpal
'    Unload Me
    BotonSalir
End Sub

Private Sub mnP_Salir2_Click()
'    Unload frmPpal
'    Unload Me
    BotonSalir
End Sub

Private Sub BotonSalir()
    Unload frmPpal
    Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Clientes
            SubmnP_Generales_Click (9)
        Case 2 'Articulos
            SubmnP_Generales_Click (13)
        Case 4 'Traspaso de postes
            SubmnG_Ventas_Click (1)
        Case 5 'Albaranes
            SubmnG_Ventas_Click (3)
        Case 6 'Cuadre Diario
            SubmnG_Ventas_Click (11)
        Case 8 'Facturación
            SubmnF_Facturacion_Click (3)
        Case 9 'Contabilización
            SubmnF_Facturacion_Click (8)
        Case 11 'Hco.Facturas
            SubmnE_Estadist_Click (1)
        Case 13
            SubmnE_Util_Click (2)
    End Select
End Sub

' ### [Monica] 05/09/2006
Private Sub PonerDatosFormulario()
Dim Config As Boolean

    Config = (vEmpresa Is Nothing) Or (vParamAplic Is Nothing)
    
    If Not Config Then HabilitarSoloPrametros_o_Empresas True

    'FijarConerrores
    CadenaDesdeOtroForm = ""

    'Poner datos visible del form
'    PonerDatosVisiblesForm
    
    'Habilitar/Deshabilitar entradas del menu segun el nivel de usuario
    PonerMenusNivelUsuario

    'Si no hay carpeta interaciones, no habra integraciones
'    Me.mnComprobarPendientes.Enabled = vConfig.Integraciones <> ""


    'Habilitar
    If Config Then HabilitarSoloPrametros_o_Empresas False
    'Panel con el nombre de la empresa
'    If Not vEmpresa Is Nothing Then
'        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
'    Else
'        Me.StatusBar1.Panels(2).Text = "Falta configurar"
'    End If


    'Si tiene editor de menus
    If TieneEditorDeMenus Then PoneMenusDelEditor
    
    
    BloqueoMenusSegunCooperativa
    
    
    'Comprobar que los iconos de la barra su correspondiente
    'entrada de menu esta habilitada sino desabilitar
'    PoneBarraMenus
    
End Sub

' ### [Monica] 05/09/2006
Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim Cad As String

    On Error Resume Next
    For Each T In Me
        Cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            'If LCase(Mid(T.Name, 1, 8)) <> "mn_b" Then
                T.Enabled = Habilitar
            'End If
        End If
    Next
    
    Me.Toolbar1.Enabled = Habilitar
    Me.Toolbar1.visible = Habilitar
    Me.mnParametros(1).Enabled = True
    Me.mnP_Generales(1).Enabled = True
    Me.mnP_Generales(2).Enabled = True
    Me.mnP_Generales(6).Enabled = True
    Me.mnP_Generales(17).Enabled = True
    
'    Me.mnCambioEmpresa.Enabled = True
End Sub


' ### [Monica] 07/11/2006
' añadida esta parte para la personalizacion de menus

Private Sub LeerEditorMenus()
Dim SQL As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    TieneEditorDeMenus = False
    SQL = "Select count(*) from appmenus where aplicacion='Arigasol'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneEditorDeMenus = True
        End If
    End If
    miRsAux.Close
        

ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub PoneMenusDelEditor()
Dim T As Control
Dim SQL As String
Dim C As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    
    SQL = "Select * from appmenususuario where aplicacion='Arigasol' and codusu = " & Val(Right(CStr(vSesion.Codusu), 3))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            SQL = SQL & miRsAux.Fields(3) & "·"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
   
    If SQL <> "" Then
        SQL = "·" & SQL
        For Each T In Me.Controls
            If TypeOf T Is menu Then
                C = DevuelveCadenaMenu(T)
                C = "·" & C & "·"
                If InStr(1, SQL, C) > 0 Then T.visible = False
           
            End If
        Next
    End If
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function DevuelveCadenaMenu(ByRef T As Control) As String

On Error GoTo EDevuelveCadenaMenu
    DevuelveCadenaMenu = T.Name & "|"
    DevuelveCadenaMenu = DevuelveCadenaMenu & T.Index '& "|"   Monica:con esto no funcionaba
    Exit Function
EDevuelveCadenaMenu:
    Err.Clear
    
End Function

Private Sub LanzaHome(Opcion As String)
    Dim I As Integer
    Dim Cad As String
    On Error GoTo ELanzaHome
    
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBD("websoporte", "sparam", "codparam", 1, "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en parametros.", vbExclamation
        Exit Sub
    End If
        
    I = FreeFile
    Cad = ""
    Open App.path & "\lanzaexp.dat" For Input As #I
    Line Input #I, Cad
    Close #I
    
    'Lanzamos
    If Cad <> "" Then Shell Cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
    
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, Cad & vbCrLf & Err.Description
    CadenaDesdeOtroForm = ""
End Sub


Private Sub BloqueoMenusSegunCooperativa()
Dim b As Boolean

    b = (vParamAplic.Cooperativa = 2)
    
    ' la facturacion ajena es unicamente de regaixo
    mnFacturacionAjena(1).visible = b
    mnFacturacionAjena(1).Enabled = b
    
    
'[Monica]13/02/2013: el mto de tarjetas debe estar abierto para todos
'    ' el mantenimiento de solicitud de tarjetas es solo de regaixo
'    mnE_Util(2).visible = b
'    mnE_Util(2).Enabled = b
'
'    Me.Toolbar1.Buttons(13).visible = b
'    Me.Toolbar1.Buttons(13).Enabled = b
    
'30/03/2007 dejamos que la factura de abono a clientes se use en Alzira
    ' factura de abono a clientes
'    mnF_Facturacion(9).visible = b
'    mnF_Facturacion(9).Enabled = b

    
    ' grabacion de modelo gasoleo B tb solo de regaixo
'    mnF_Facturacion(10).visible = b
'    mnF_Facturacion(10).Enabled = b

'06/03/2007 este listado lo pueden sacar tb en alzira
'    ' estadisticas por articulo y forma de pago para Regaixo
'    mnE_Estadist(9).visible = b
'    mnE_Estadist(9).Enabled = b
    
   ' impresion de cierre de turno
    mnE_Tanques(3).visible = b
    mnE_Tanques(3).Enabled = b

'30/05/2007 declaracion de gasoleo a B a Hda
    mnE_Estadist(12).visible = Not b
    mnE_Estadist(12).Enabled = Not b
    
    ' 21/09/2011: la estadistica diaria por forma de pago solo es para Pobla del Duc
    Me.mnG_Ventas(9).visible = (vParamAplic.Cooperativa = 4)  ' solo para Pobla del Duc
    Me.mnG_Ventas(9).Enabled = (vParamAplic.Cooperativa = 4)  ' solo para Pobla del Duc
    
    
    '[Monica]29/05/2014: el traspaso de tarjetas de tpv solo la hace Alzira
    Me.mnE_Util(13).visible = (vParamAplic.Cooperativa = 1)
    Me.mnE_Util(13).Enabled = (vParamAplic.Cooperativa = 1)
    
    
    '[Monica]02/01/2019: Asignacion de nuevo cliente (solo lo hace ribarroja)
    Me.mnE_Util(14).visible = (vParamAplic.Cooperativa = 5)
    Me.mnE_Util(14).Enabled = (vParamAplic.Cooperativa = 5)
    
    
    
    '[Monica]26/06/2014: la impresion de tarjetas no la ve Alzira pq tiene la propia en el mto.de socios
    Me.mnE_Util(2).visible = (vParamAplic.Cooperativa <> 1)
    Me.mnE_Util(2).Enabled = (vParamAplic.Cooperativa <> 1)
    Me.Toolbar1.Buttons(13).visible = (vParamAplic.Cooperativa <> 1)
    Me.Toolbar1.Buttons(13).Enabled = (vParamAplic.Cooperativa <> 1)
    
    '[Monica]01/07/2014: la contabilizacion de las facturas es el traspaso al unico
    If vParamAplic.Cooperativa = 4 Then
        Me.mnF_Facturacion(8).Caption = "Traspaso a Unicoo"
    End If
    
    '[Monica]12/03/2015: cambio de empresa solo para Alzira
    Dim NRegs As Integer
    NRegs = TotalRegistros("select count(*) from usuarios.empresasarigasol")
    Me.mnP_Generales(21).visible = (NRegs > 1)
    Me.mnP_Generales(21).Enabled = (NRegs > 1)
    
    
   ' impresion de estadisticas por turno solo para pobla del duc
    mnE_Tanques(5).visible = (vParamAplic.Cooperativa = 4)
    mnE_Tanques(5).Enabled = (vParamAplic.Cooperativa = 4)
    
    
    
End Sub




Private Sub CargarImagen()

On Error GoTo eCargarImagen
    Me.Picture = LoadPicture(App.path & "\fondo.dat")
    Exit Sub
eCargarImagen:
    MuestraError Err.Number, "Error cargando imagen. LLame a soporte"
    End
End Sub


Private Sub PonerMenusNivelUsuario()
Dim b As Boolean

    b = (vSesion.Nivel = 0)    'Administradores y root

    Me.mnE_Util(11).Enabled = b
    Me.mnE_Util(11).visible = b
    
    
    Me.mnE_Util(3).Enabled = b
    Me.mnE_Util(3).visible = b
    
    
End Sub

Public Sub mnCambioEmpresa()
Dim Cad As String

    'Borramos temporal
    Conn.Execute "Delete from zbloqueos where codusu = " & vSesion.Codigo


    CadenaDesdeOtroForm = vSesion.Login & "|" & vSesion.PasswdPROPIO & "|"
    
    frmLogin.Show vbModal

    Screen.MousePointer = vbHourglass
    'Cerramos la conexion
    Conn.Close
    ConnConta.Close


    If AbrirConexionAriGasol("root", "aritel", vSesion.CadenaConexion) = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
        Screen.MousePointer = vbDefault
        End
    End If

    'Carga Datos de la Empresa y los Niveles de cuentas de Contabilidad de la empresa
    'Crea la Conexion a la BD de la Contabilidad
    LeerDatosEmpresa


    InicializarFormatos
    teclaBuscar = 43

    Load frmPpal
    
    PonerDatosFormulario
    
    If vEmpresa Is Nothing Then
        Caption = "AriGasol" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & " FALTA CONFIGURAR"
    Else
        Caption = "AriGasol" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  Empresa: " & vEmpresa.nomEmpre & Cad & _
                  "   -  Usuario: " & vSesion.Nombre
    End If
    
    LeerEditorMenus

    If vParamAplic.ContabilidadNueva And (vSesion.Nivel = 0 Or vSesion.Nivel = 1) Then FrasPendientesContabilizar False




    Screen.MousePointer = vbDefault


End Sub



