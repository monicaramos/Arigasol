Attribute VB_Name = "modMenuClick"
Option Explicit


Private Sub Construc(Nom As String)
    MsgBox Nom & ": en construcció..."
End Sub

' ******* DATOS BASICOS *********

Public Sub SubmnP_Generales_Click(Index As Integer)

    Select Case Index
        Case 1: frmConfParamGral.Show vbModal
                PonerDatosPpal
        Case 2: frmConfParamAplic.Show vbModal
        Case 3: frmManMovim.Show vbModal
        Case 4: frmConfParamRpt.Show vbModal
        Case 5: frmManTipCre.Show vbModal
        Case 6: frmMantenusu.Show vbModal
        
        Case 8: frmManCoope.Show vbModal
        Case 9: frmManClien.Show vbModal
        Case 10: frmManTraba.Show vbModal
        Case 11: frmManFamia.Show vbModal
        Case 12: frmManArtic.Show vbModal
        Case 13: frmManSitua.Show vbModal
        Case 14: frmManFpago.Show vbModal
        Case 15: frmManBanco.Show vbModal
        Case 16: frmManGrupo.Show vbModal 'VRS:2.0.2(0)
        Case 17: frmManProve.Show vbModal 'VRS:2.0.2(0)
        Case 18: frmManEntidades.Show vbModal
        Case 20: MDIppal.mnCambioEmpresa        ' cambio de empresa
        
        Case 21: End
    End Select
End Sub


' *******  VENTAS DIARIAS *********

Public Sub SubmnG_Ventas_Click(Index As Integer)
    Select Case Index
        Case 1: DesBloqueoManual ("TRASPOST")
                If Not BloqueoManual("TRASPOST", "1") Then
                    MsgBox "No se puede realizar el Traspaso de Poste. Hay otro usuario realizándolo.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
'[Monica] 25/01/2010 añadido nuevo trapaso de poste para castelduc, diferenciamos cooperativas
                    Select Case vParamAplic.Cooperativa
                        '[Monica]09/01/2013: Nueva cooperativa Ribarroja
                        Case 1, 2, 5 ' Alzira y Regaixo y Ribarroja
                    
    '--monica : traen turno van a cambiar tb en alzira
    ''--MONICA: antes del nuevo traspaso de postes
    '                    If vParamAplic.Cooperativa = 1 Then
    '                        frmTrasPoste.Show vbModal
    '                    Else
    '' de momento solo cambiasn en regaixo.
    ''                        frmTrasPoste2.Show vbModal
                            frmTrasAlvic.Show vbModal
    '                    End If
    
                        Case 3 ' Castelduc
                            frmTrasPosteCast.Show vbModal
                    End Select
                        
    
                End If
        Case 3: ' mantenimiento de albaranes
                If vParamAplic.Cooperativa = 4 Then
                    frmAlbaranQuatre.Show vbModal
                Else
                    frmAlbaran.Show vbModal
                End If
        Case 4: frmEstArticAlb.Show vbModal
        Case 5: frmPrefactur.Show vbModal
        Case 6: frmErroresAlb.Show vbModal
        Case 7: frmCompDescuadre.Show vbModal
        Case 8: frmCambioCliente.Show vbModal
        Case 9: ' estadistica por forma de pago
                frmEstFPago.Show vbModal
        Case 11: frmCuadreDiario.Show vbModal
        Case 13: ' los cierres de turno son diferentes en regaixo que en alzira
                 '[Monica]09/01/2013: Nueva cooperativa Ribarroja
                 If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 5 Then
                    frmContCieTurno.Show vbModal
                 Else
                    frmContCieTurnoReg.Show vbModal
                 End If
    End Select
End Sub

' *******  FACTURACION *********

Public Sub SubmnF_Facturacion_Click(Index As Integer)
    Select Case Index
        Case 1:    'Bloquear para que nadie mas pueda realizar el traspaso
                DesBloqueoManual ("TRASTPV")
                If Not BloqueoManual("TRASTPV", "1") Then
                    MsgBox "No se puede realizar el traspaso TPV. Hay otro usuario realizándolo.", vbExclamation
                    Screen.MousePointer = vbDefault
                    
                Else
                    frmTrasTpv.Show vbModal 'modifico el programa original de Manolo
                End If
                
        Case 2: ' proceso de prefacturacion para Alzira se cambian importes y precios de albaranes
                ' aplicando un margen sobre los precios solo para clientes de bonificacion especial
                DesBloqueoManual ("PREFACT")
                If Not BloqueoManual("PREFACT", "1") Then
                    MsgBox "No se puede realizar el proceso de Prefacturación. Hay otro usuario realizándolo.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    frmPreFactBonif.Show vbModal 'modifico el programa original de Manolo
                End If
                
                
        Case 3: DesBloqueoManual ("FACTURAC")
                If Not BloqueoManual("FACTURAC", "1") Then
                    MsgBox "No se puede realizar el proceso de facturación. Hay otro usuario realizándolo.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    frmFacturas.Show vbModal
                End If
        
        Case 4: frmFactgas.NumCod = 0
                frmFactgas.Show vbModal
        Case 5: ' envio de facturas por email
                AbrirListadoOfer 315
                
        Case 6: ' envio de facturas FacturaE
                AbrirListadoOfer 316
                
        Case 8: '[Monica]01/07/2014: en el caso de Pobla del Duc pasan los datos al Unico
                If vParamAplic.Cooperativa = 4 Then
                    frmPaseUnico.Show vbModal
                Else
                    frmContabFact.Show vbModal
                End If
        Case 10: frmBusLinFactu.Show vbModal
        Case 12: 'MsgBox "En proceso de implementación", vbExclamation
                frmFacturaAbonoCli.Show vbModal
        Case 13: frmListado.OpcionListado = 12
                 frmListado.Show vbModal
        Case 14: frmGrabGasoleoB.Show vbModal
        Case 15: frmEst569.Show vbModal
        Case 16: ' Devolucion del Centimo Sanitario
                 frmCentimoSanitario.Ajenas = 0
                 frmCentimoSanitario.Show vbModal
        
    End Select
End Sub

Public Sub SubmnF_FacturacionAjena_Click(Index As Integer)
    Select Case Index
        Case 1: DesBloqueoManual ("FACTURAC")
                If Not BloqueoManual("FACTURAC", "1") Then
                    MsgBox "No se puede realizar el proceso de facturación ajena. Hay otro usuario realizándolo.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    frmFacturaAjena.Show vbModal
                End If
        Case 2: ' reimpresion de facturas ajenas
                frmFactgas.NumCod = 1
                frmFactgas.Show vbModal
        Case 4: ' hco de facturas ajenas
                frmHcoFact.Tipo = 1 ' seleccionamos la tabla schfac
                frmHcoFact.Show vbModal
        Case 6: ' contabilización de las facturas únicamente en tesoreria
                frmContabTesor.Show vbModal
        Case 8: ' ventas  por cliente en ajena
                frmEstCliimp.NumCod = 1
                frmEstCliimp.Show vbModal
        Case 9: ' ventas de articulos por cliente en ajena
                frmEstCliArtAje.Show vbModal
        Case 10: ' factura de bonificacion a socios
                frmFacturaAbonoSoc.Show vbModal
        Case 12: ' Contabilizar Factura de Catadau
                frmContabFactCoop.Show vbModal
        Case 13: ' Devolucion del Centimo Sanitario
                frmCentimoSanitario.Ajenas = 1
                frmCentimoSanitario.Show vbModal
                
    End Select
End Sub


' *******  ESTADÍSTICAS *********

Public Sub SubmnE_Estadist_Click(Index As Integer)
    Select Case Index
        Case 1: frmHcoFact.Tipo = 0 ' seleccionamos la tabla schfac
                frmHcoFact.numfactu = 0
                frmHcoFact.letraserie = ""
                frmHcoFact.Show vbModal
        Case 2: frmEstDiario.Show vbModal
        Case 3: frmEstCliimp.NumCod = 0
                frmEstCliimp.Show vbModal
        Case 4: frmEstCliArt.Show vbModal
        Case 5: frmEstArtic.Show vbModal
        Case 6: frmEstVtasdia.Show vbModal
        Case 7: frmEstTarArt.Show vbModal
        Case 8: frmEstRangos.Show vbModal
        Case 9: frmEstGasPro.Show vbModal
        Case 10: frmEvoMensCli.Show vbModal
        Case 11: frmCertGasB.Show vbModal
        Case 12: frmCertGasBHda.Show vbModal
        Case 13: frmEstConsFec.Show vbModal
    
        Case 14: frmListMovArtFam.Show vbModal
        Case 15: frmListMargenVtas.Show vbModal
        Case 16: frmListMargenVtasCli.Show vbModal
        Case 18:
                'Bloquear para que nadie mas pueda realizar el traspaso
                DesBloqueoManual ("TRASHCO")
                If Not BloqueoManual("TRASHCO", "1") Then
                    MsgBox "No se puede realizar el traspaso a Histórico 1. Hay otro usuario realizándolo.", vbExclamation
                    Screen.MousePointer = vbDefault
                Else
                    frmTrasHcoFras1.Show vbModal
                End If
        Case 19: '[Monica]24/07/2013
                frmHcoFact.Tipo = 2 ' seleccionamos la tabla schfac1
                frmHcoFact.numfactu = 0
                frmHcoFact.letraserie = ""
                frmHcoFact.Show vbModal
    End Select
End Sub

' *******  TANQUES - MANGUERAS *********

Public Sub SubmnE_Tanques_Click(Index As Integer)

    Select Case Index
        Case 1: frmTurnos.Show vbModal
        Case 2: frmRecauda.Show vbModal
        Case 3: frmCieTurnoReg.Show vbModal
        Case 4: frmEstArticRegaixo.Show vbModal
    End Select
End Sub


' *******  Compras  *********
' *******  ALBARANES PROVEEDOR  *********

Public Sub SubmnC_Compras_Click(Index As Integer)
    Select Case Index
        Case 1: 'albaranes proveedor
                frmComEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
                frmComEntAlbaranes.EsHistorico = False
                frmComEntAlbaranes.Show vbModal
        Case 2: 'Historico albaranes de compras a proveedores
                frmComEntAlbaranes.EsHistorico = True
                frmComEntAlbaranes.Show vbModal
        Case 3: 'Listado de Albaranes pendientes de Factura
                AbrirListadoOfer (308) '308: List. Albaranes pte facturar
        Case 5: 'Recepción de facturas de proveedor
                frmComFacturar.Show vbModal
        Case 6: 'historico de facturas
                frmComHcoFacturas.hcoCodMovim = ""
                frmComHcoFacturas.Show vbModal
        Case 8: 'Contabilizar Facturas
                AbrirListado2 (224) 'Para pedir datos
        
    End Select
End Sub

' *******  ESTADISTICAS DE COMPRAS *********

Public Sub SubmnE_EstComp_Click(Index As Integer)
    Select Case Index
        Case 1: AbrirListadoOfer (310)
        Case 2: AbrirListadoOfer (311)
        Case 3: AbrirListadoOfer (312)
    
    End Select
End Sub

' *******  Compras  *********
' *******  INVENTARIO  *********

Public Sub SubmnC_ComprasInven_Click(Index As Integer)
    Select Case Index
        Case 1: ' Toma de Inventario
                AbrirListado2 (12)
        Case 2:  'entrada de existencia real
                frmAlmInventario.Show vbModal
        Case 3:  'listado de diferencias
                AbrirListado2 (13)
        Case 4:  'actualizar diferencias
                AbrirListado2 (14)
        Case 5:  'valoracion de stocks inventariados
                AbrirListado2 (16)
    End Select
End Sub



' *******  UTILIDADES *********

Public Sub SubmnE_Util_Click(Index As Integer)
    Select Case Index
        Case 1:  frmManRangos.Show vbModal
        Case 2 'VRS:2.0.2(3) solicitud de tarjetas
               ' frmSolTarjetas.Show vbModal antes era este ahora van a imprimir ellos
                frmImpTarjetas.Show vbModal
        Case 3: frmCambios.Show vbModal
        Case 5: frmDecGasPro.Show vbModal
        Case 7: frmComprobImpuesto.Show vbModal
        Case 9: ' comprobacion de caracteres de multibase
                frmCaracteresMB.Show vbModal
        Case 11: frmDeshacerFact.Show vbModal ' deshacer facturacion
        Case 12: frmBackUP.Show vbModal
        Case 13: frmExpTPV2.Show vbModal
    End Select
End Sub



Public Sub AbrirListado2(numero As Integer)
    Screen.MousePointer = vbHourglass
    frmListado2.OpcionListado = numero
'    frmListado2.OptProve = (Not DeTransporte)
'    DeTransporte = False
    
    frmListado2.Show vbModal
    Screen.MousePointer = vbDefault
End Sub




