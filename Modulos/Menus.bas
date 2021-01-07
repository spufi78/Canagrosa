Attribute VB_Name = "Menus"
Public Sub pBuildMenus()
    With frmMenu.SmartMenuXP1.MenuItems
        .Clear
        .Add 0, "menuLaboratorio", , "Laboratorio"
          .Add "menuLaboratorio", "opLaboratorio_01", , "Recepci�n de Muestras"
          .Add "menuLaboratorio", "opLaboratorio_02", , "Recepci�n Control de Eficacia"
          .Add "menuLaboratorio", "opLaboratorio_03", , "Recepci�n de Sellante"
'          .Add "menuLaboratorio", "opLaboratorio_15", , "Recepci�n de Pintura"
          .Add "menuLaboratorio", "opLaboratorio_16", , "Recepci�n de Ensayos F�sicos Iberia"
          .Add "menuLaboratorio", "opLaboratorio_17", , "Recepci�n Probetas HENKEL"
          .Add "menuLaboratorio", , smiSeparator
          .Add "menuLaboratorio", "opLaboratorio_04", , "Plantilla de Toma de Muestras"
          .Add "menuLaboratorio", , smiSeparator
          .Add "menuLaboratorio", "opLaboratorio_05", , "Listado y Registro de Muestras", , vbCtrlMask, vbKeyF2
          .Add "menuLaboratorio", "opLaboratorio_06", , "Determinaciones Pendientes"
          .Add "menuLaboratorio", "opLaboratorio_07", , "Trabajo Pendiente"
          .Add "menuLaboratorio", "opLaboratorio_08", , "Muestras a entregar"
          .Add "menuLaboratorio", "opLaboratorio_09", , "Probetas Pendientes Analisis"
          .Add "menuLaboratorio", , smiSeparator
          .Add "menuLaboratorio", "opLaboratorio_10", , "Localizador", , , vbKeyF1
        .Add 0, "menuInformes", , "Informes"
          .Add "menuInformes", "opInformes_01", , "Muestras Pendientes de Cerrar"
          .Add "menuInformes", "opInformes_06", , "Muestras Pendientes de Revisi�n"
          .Add "menuInformes", "opInformes_07", , "Muestras Pendientes de Env�o"
          .Add "menuInformes", "opInformes_08", , "Muestras Fuera de Plazo"
          .Add "menuInformes", "opInformes_02", , "Muestras Analizadas por Cliente y Fecha"
          .Add "menuInformes", "opInformes_03", , "Informes de Registro"
          .Add "menuInformes", "opInformes_05", , "Informe de Duplicados"
          .Add "menuInformes", , smiSeparator
          .Add "menuInformes", "opInformes_10", , "Muestras Cerradas por Analista"
          .Add "menuInformes", "opInformes_11", , "Determinaciones Realizadas por Analista"
          .Add "menuInformes", , smiSeparator
          .Add "menuInformes", "opInformes_04", , "Informes de Facturaci�n"
          .Add "menuInformes", "opInformes_09", , "Informe de Pedidos de Clientes"
        .Add 0, "menuBanos", , "Ba�os"
          .Add "menuBanos", "opBanos_01", , "Hist�rico de Ba�os"
        .Add 0, "menuAlodine", , "Alodine y Suministros"
          .Add "menuAlodine", "opAlodine_01", , "Listado de Lotes Suministrados Alodine"
          .Add "menuAlodine", "opAlodine_05", , "Listado de Suministros"
          .Add "menuAlodine", "opAlodine_06", , "Listado de Productos Controlados"
          .Add "menuAlodine", , smiSeparator
          .Add "menuAlodine", "opAlodine_02", , "Tipos de Lotes de Alodine"
          .Add "menuAlodine", "opAlodine_04", , "Tipos de Suministros"
          .Add "menuAlodine", "opAlodine_03", , "Tipos de Capacidad"
          'M1290-I
          .Add "menuAlodine", , smiSeparator
          .Add "menuAlodine", "opAlodine_07", , "Informe de Alodine Suministrado"
          'M1290-F
        .Add 0, "menuReactivos", , "Reactivos"
          .Add "menuReactivos", "opReactivos_01", , "Externos"
          .Add "menuReactivos", "opReactivos_04", , "   Inventario"
          .Add "menuReactivos", "opReactivos_02", , "Propios"
          .Add "menuReactivos", , smiSeparator
          .Add "menuReactivos", "opReactivos_03", , "Pedidos"
          .Add "menuReactivos", , smiSeparator
          .Add "menuReactivos", "opReactivos_05", , "Informe de Pedidos realizados"
          .Add "menuReactivos", "opReactivos_06", , "Informe de Pedidos Por Fecha de Entrega"
        .Add 0, "menuFacturacion", , "Facturaci�n"
          .Add "menuFacturacion", "opFacturacion_01", , "Muestras Pendientes de Facturaci�n"
          .Add "menuFacturacion", "opFacturacion_02", , "Listado de Documentos de Pago"
          .Add "menuFacturacion", "opFacturacion_03", , "Albaranes Pendientes de Facturar"
          .Add "menuFacturacion", , smiSeparator
          .Add "menuFacturacion", "opFacturacion_04", , "Facturaci�n de Conceptos"
          .Add "menuFacturacion", , smiSeparator
          .Add "menuFacturacion", "opFacturacion_05", , "Facturaci�n de Alodine"
          .Add "menuFacturacion", "opFacturacion_09", , "Facturaci�n de Suministros"
          .Add "menuFacturacion", "opFacturacion_10", , "Facturaci�n de Productos Controlados"
          .Add "menuFacturacion", "opFacturacion_12", , "Facturaci�n Henkel"
          .Add "menuFacturacion", , smiSeparator
'          .Add "menuFacturacion", "opFacturacion_06", , "Contabilidad"
          .Add "menuFacturacion", "subContabilidad", , "Contabilidad"
            .Add "subContabilidad", "opFacturacion_06", smiPicture, "Exportar facturas a Contaplus"
            .Add "subContabilidad", "opFacturacion_07", smiPicture, "Importar asientos desde Contaplus"
          .Add "menuFacturacion", , smiSeparator
          .Add "menuFacturacion", "opFacturacion_08", , "Informe por Familias y Sectores"
          .Add "menuFacturacion", "opFacturacion_11", , "Informe de muestras y conceptos"
        .Add 0, "menuIndicadores", , "Indicadores"
          'M1039-I
          .Add "menuIndicadores", "opIndicadores_07", , "Gesti�n de Indicadores"
          .Add "menuIndicadores", "opIndicadores_06", , "Listado de Indicadores"
          .Add "menuIndicadores", , smiSeparator
          .Add "menuIndicadores", "opIndicadores_01", , "Generar Informe (Antiguo)"
          .Add "menuIndicadores", , smiSeparator
          .Add "menuIndicadores", "opIndicadores_02", , "Indicadores  (Antiguo)"
          .Add "menuIndicadores", "opIndicadores_03", , "Campos (Antiguo)"
          .Add "menuIndicadores", "opIndicadores_04", , "Frecuencias (Antiguo)"
          .Add "menuIndicadores", "opIndicadores_05", , "Funciones (Antiguo)"
          'M1039-F
        .Add 0, "menuCalidad", , "Calidad"
          .Add "menuCalidad", "subCalidadDocumentos", , "Documentos de Calidad"
            .Add "subCalidadDocumentos", "opCalidad_01", smiPicture, "Listado de Documentos"
            .Add "subCalidadDocumentos", "opCalidad_31", smiPicture, "Requerimientos para PNT"
            .Add "subCalidadDocumentos", , smiSeparator
            .Add "subCalidadDocumentos", "opCalidad_02", smiPicture, "Familias"
            .Add "subCalidadDocumentos", "opCalidad_03", smiPicture, "SubFamilias"
            .Add "subCalidadDocumentos", "opCalidad_04", smiPicture, "Responsables"
            .Add "subCalidadDocumentos", "opCalidad_05", smiPicture, "Estados"
          .Add "menuCalidad", "subCalidadNormas", , "Listado de Normas Controladas"
            .Add "subCalidadNormas", "opCalidad_06", smiPicture, "Listado de Normas"
            .Add "subCalidadNormas", , smiSeparator
            .Add "subCalidadNormas", "opCalidad_07", smiPicture, "Tipos"
            .Add "subCalidadNormas", "opCalidad_08", smiPicture, "Sectores"
            .Add "subCalidadNormas", "opCalidad_09", smiPicture, "Estados"
            .Add "subCalidadNormas", "opCalidad_10", smiPicture, "Subtipos"
          .Add "menuCalidad", , smiSeparator
          .Add "menuCalidad", "opCalidad_11", , "Ofertas"
          .Add "menuCalidad", , smiSeparator
          .Add "menuCalidad", "opCalidad40", , "Listado de Incidencias"
            .Add "opCalidad40", "opCalidad_12", smiPicture, "Listado de Incidencias"
            .Add "opCalidad40", "opCalidad_20", smiPicture, "Listado de Proc. No Conformidad"
            .Add "opCalidad40", "opCalidad_21", smiPicture, "Listado de Acciones de Proc. NC"
            .Add "opCalidad40", , smiSeparator
'            .Add "opCalidad40", "opCalidad_40", smiPicture, "Causas (Auditorias, Cliente, etc)"
            .Add "opCalidad40", "opCalidad_41", smiPicture, "Auditor�as Internas"
            .Add "opCalidad40", "opCalidad_42", smiPicture, "Auditor�as Externas"
            .Add "opCalidad40", "opCalidad_43", smiPicture, "Detecci�n Interna"
            .Add "opCalidad40", , smiSeparator
            .Add "opCalidad40", "opCalidad_13", smiPicture, "Tipos de Hechos"
            .Add "opCalidad40", "opCalidad_14", smiPicture, "Or�genes"
            .Add "opCalidad40", "opCalidad_15", smiPicture, "Estados"
            .Add "opCalidad40", "opCalidad_16", smiPicture, "Departamentos"
            .Add "opCalidad40", "opCalidad_17", smiPicture, "Afectados"
            .Add "opCalidad40", "opCalidad_22", smiPicture, "Estudio"
          .Add "menuCalidad", , smiSeparator
          .Add "menuCalidad", "opCalidad50", , "Programa de Auditorias"
            .Add "opCalidad50", "opCalidad_18", smiPicture, "Gesti�n del Programa"
            .Add "opCalidad50", "opCalidad_19", smiPicture, "Gesti�n de �reas"
          'MXXXX-I
          '.Add "menuCalidad", "opCalidad_29", , "Formadores externos"
          'MXXXX-F
          .Add "menuCalidad", , smiSeparator
          .Add "menuCalidad", "opCalidad_30", , "Matriz Cualificaciones"
        .Add 0, "menuEquipos", , "Equipos"
          .Add "menuEquipos", "opEquipos10", , "Gesti�n de Equipos"
            .Add "opEquipos10", "opEquipos_01", smiPicture, "Listado de Equipos"
            .Add "opEquipos10", , smiSeparator
            .Add "opEquipos10", "opEquipos_08", smiPicture, "Tipos de Equipo"
            .Add "opEquipos10", , smiSeparator
            .Add "opEquipos10", "opEquipos_02", smiPicture, "Areas de Metrolog�a"
            .Add "opEquipos10", "opEquipos_03", smiPicture, "Situaciones"
            .Add "opEquipos10", "opEquipos_04", smiPicture, "Periodicidad"
            .Add "opEquipos10", "opEquipos_09", smiPicture, "Par�metros T�cnicos"
          .Add "menuEquipos", "opEquipos20", , "Planes de Mantenimiento"
            .Add "opEquipos20", "opEquipos_05", smiPicture, "Planes de Mantenimiento"
            .Add "opEquipos20", , smiSeparator
            .Add "opEquipos20", "opEquipos_06", smiPicture, "Acciones"
            .Add "opEquipos20", "opEquipos_07", smiPicture, "Familias de Acciones"
        'M1241-I
        .Add 0, "menuPedidos", , "Subcontrataciones y Pedidos"
          .Add "menuPedidos", "opPedidos_01", , "Subcontratacion de Ensayos y Calibraciones"
          .Add "menuPedidos", "opPedidos_02", , "Pedidos a Proveedor"
        'M1241-F
        .Add 0, "menuEnvios", , "Env�os"
          .Add "menuEnvios", "opEnvios_02", , "Env�o de Paquetes"
          .Add "menuEnvios", "opEnvios_03", , "Empresas de Mensajer�a"
        If pc_es_tablet Then
            .Add 0, "menuTablets", , "Tablets"
              .Add "menuTablets", "opTablets_01", , "Listado y Registro de Muestras"
        End If
        .Add 0, "menuRRHH", , "R.R.H.H."
          .Add "menuRRHH", , smiSeparator
          .Add "menuRRHH", "opRRHH_01", , "Gesti�n de Personal"
          .Add "menuRRHH", "opRRHH_02", , "Departamentos y Categorias"
          
'          If USUARIO.getPER_EMPLEADOS = True And (USUARIO.getUSO = "MARIBEL-PC" Or USUARIO.getUSO = "RRHH-PC" Or USUARIO.getUSO = "DES-JGM") Then
'          .Add "menuRRHH", , smiSeparator
'          .Add "menuRRHH", "opRRHH_03", , "Control de N�minas"
'          End If
          
        'M0996-I
        .Add 0, "menuFormacion", , "Formaci�n"
          .Add "menuFormacion", "opFormacion_01", , "Registros de Formaci�n (RFI)"
        'M0996-F
        'M1110-I
          .Add "menuFormacion", "opFormacion_03", , "Plan de formaci�n Anual"
          .Add "menuFormacion", , smiSeparator
          .Add "menuFormacion", "opFormacion_02", , "Cursos Formaci�n / Documentaci�n"
          .Add "menuFormacion", "opFormacion_04", , "Certificaci�n de formadores"
        'M1110-F
        'JGM : TESORERIA
        .Add 0, "menuTesoreria", , "Tesorer�a"
          .Add "menuTesoreria", "opTesoreria_01", , "Listado de Facturas de Proveedor"
          .Add "menuTesoreria", "opTesoreria_02", , "Listado de Otros Gastos"
          .Add "menuTesoreria", "opTesoreria_10", , "Remesas de Pago"
          .Add "menuTesoreria", "opTesoreria_07", , "Bancos"
          .Add "menuTesoreria", "opTesoreria_08", , "Seguros"
          .Add "menuTesoreria", "opTesoreria_09", , "Prestamos"
          .Add "menuTesoreria", , smiSeparator
          .Add "menuTesoreria", "opTesoreria_11", , "Tesoreria"
          .Add "menuTesoreria", , smiSeparator
          .Add "menuTesoreria", "subContabilidadPro", , "Contabilidad Proveedores"
            .Add "subContabilidadPro", "opTesoreria_03", smiPicture, "Exportar facturas de proveedor a Contaplus"
            'M1362-I
            .Add "subContabilidadPro", "opTesoreria_06", smiPicture, "Exportar asientos de pago a Contaplus"
            'M1362-F
          .Add "menuTesoreria", , smiSeparator
          .Add "menuTesoreria", "subTesoreriaSubcuentas", , "Subcuentas"
            .Add "subTesoreriaSubcuentas", "opTesoreria_04", smiPicture, "Subcuentas de Gastos"
            .Add "subTesoreriaSubcuentas", "opTesoreria_05", smiPicture, "Subcuentas de Pagos"
        'TESORERIA-F
        .Add 0, "menuMantenimiento", , "Mantenimiento"
          .Add "menuMantenimiento", "opMantenimiento_01", , "Tipos de Muestras"
          .Add "menuMantenimiento", "opMantenimiento_02", , "Tipos de An�lisis"
          .Add "menuMantenimiento", "opMantenimiento_03", , "Tipos de Determinaciones"
          .Add "menuMantenimiento", "opMantenimiento_04", , "Dependencias Determinaciones"
          .Add "menuMantenimiento", "opMantenimiento_05", , "Datos Espec�ficos"
          .Add "menuMantenimiento", "opMantenimiento_61", , "Descripciones de Productos"
          .Add "menuMantenimiento", "opMantenimiento60", , "Ba�os"
            .Add "opMantenimiento60", "opMantenimiento_06", smiPicture, "Ba�os"
            .Add "opMantenimiento60", , smiSeparator
            .Add "opMantenimiento60", "opMantenimiento_07", smiPicture, "Lineas"
            .Add "opMantenimiento60", "opMantenimiento_60", smiPicture, "Instalaciones"
            .Add "opMantenimiento60", "opMantenimiento_08", smiPicture, "Procesos Base"
            .Add "opMantenimiento60", "opMantenimiento_09", smiPicture, "Soluciones"
            .Add "opMantenimiento60", "opMantenimiento_10", smiPicture, "Frecuencias de Muestreo"
          .Add "menuMantenimiento", "opMantenimiento11", , "Reactivos"
            .Add "opMantenimiento11", "opMantenimiento_11", smiPicture, "Tipos de Sustancias/Materiales" ' "Tipos de Reactivos Externos"
            .Add "opMantenimiento11", "opMantenimiento_12", smiPicture, "Botes de Reactivos Externos/Productos Controlados" ' "Tipos de Botes de Reactivos Externos"
            .Add "opMantenimiento11", , smiSeparator
            .Add "opMantenimiento11", "opMantenimiento_13", smiPicture, "Tipos de Reactivos Propios/Suministros"
          .Add "menuMantenimiento", "opMantenimiento_14", , "F�rmulas"
          .Add "menuMantenimiento", "opMantenimiento_15", , "Unidades"
          .Add "menuMantenimiento", "opMantenimiento_16", , "Envases"
          .Add "menuMantenimiento", "opMantenimiento_17", , "Tipos de Caducidad"
          .Add "menuMantenimiento", "opMantenimiento_38", , "Videos"
          .Add "menuMantenimiento", "opMantenimiento18", , "Ensayos de Eficacia"
            .Add "opMantenimiento18", "opMantenimiento_18", smiPicture, "Fichas de Control"
            .Add "opMantenimiento18", "opMantenimiento_19", smiPicture, "Tipos de Ensayos de Eficacia"
            .Add "opMantenimiento18", "opMantenimiento_20", smiPicture, "Lotes de Probetas"
            .Add "opMantenimiento18", "opMantenimiento_37", smiPicture, "Materiales/Pinturas"
            .Add "opMantenimiento18", "opMantenimiento_39", smiPicture, "Dimensiones"
          .Add "menuMantenimiento", "opMantenimiento_21", , "Sellantes"
          .Add "menuMantenimiento", "opMantenimiento_22", , "Fluidos"
'          .Add "menuMantenimiento", "opMantenimiento_40", , "Pinturas"
          .Add "menuMantenimiento", "opMantenimiento_41", , "Etiquetas para Soluciones"
          .Add "menuMantenimiento", "opMantenimiento50", , "Ensayos F�sicos Iberia"
            .Add "opMantenimiento50", "opMantenimiento_51", smiPicture, "Procesos"
            .Add "opMantenimiento50", "opMantenimiento_52", smiPicture, "Fichas"
            .Add "opMantenimiento50", "opMantenimiento_53", smiPicture, "Ensayos"
            .Add "opMantenimiento50", "opMantenimiento_54", smiPicture, "Recubrimientos"
            .Add "opMantenimiento50", "opMantenimiento_55", smiPicture, "Clientes Internos"
            .Add "opMantenimiento50", "opMantenimiento_56", smiPicture, "Fabricantes"
            .Add "opMantenimiento50", "opMantenimiento_57", smiPicture, "Product Type"
            .Add "opMantenimiento50", "opMantenimiento_58", smiPicture, "Number And Type"
          .Add "menuMantenimiento", "opMantenimiento70", , "Airbus"
            .Add "opMantenimiento70", "opMantenimiento_70", smiPicture, "Plants"
            .Add "opMantenimiento70", "opMantenimiento_71", smiPicture, "Plants -> Definitions"
          .Add "menuMantenimiento", , smiSeparator
          .Add "menuMantenimiento", "opMantenimiento_23", , "Clientes"
          .Add "menuMantenimiento", "opMantenimiento_24", , "Proveedores"
'          .Add "menuMantenimiento", "opMantenimiento_25", , "Usuarios", pGetPicture(1)
          .Add "menuMantenimiento", "opMantenimiento_25", , "Usuarios"
'          .Add "menuMantenimiento", "opMantenimiento_26", , "Empleados"
          .Add "menuMantenimiento", , smiSeparator
          .Add "menuMantenimiento", "opMantenimiento_27", , "Sectores"
          .Add "menuMantenimiento", "opMantenimiento_28", , "Familias"
          .Add "menuMantenimiento", "opMantenimiento_29", , "Formas de Pago"
          .Add "menuMantenimiento", "subTarifas", , "Tarifas"
            .Add "subTarifas", "opMantenimiento_30", smiPicture, "Gesti�n de Tarifas"
            .Add "subTarifas", , smiSeparator
            .Add "subTarifas", "opMantenimiento_31", smiPicture, "Alta, Baja y Modificaci�n de Tarifas"
            .Add "subTarifas", "opMantenimiento_32", smiPicture, "C�digos Tarifarios"
            .Add "subTarifas", "opMantenimiento_33", smiPicture, "Familias de c�digos Tarifarios"
          .Add "menuMantenimiento", , smiSeparator
          .Add "menuMantenimiento", "opMantenimiento_62", , "Inventario"
          .Add "menuMantenimiento", , smiSeparator
          .Add "menuMantenimiento", "opMantenimiento_34", , "Usuarios Conectados"
          .Add "menuMantenimiento", "opMantenimiento_35", , "Parametros"
          .Add "menuMantenimiento", "opMantenimiento_36", , "Acerca de iXitec..."
        .Add 0, "menuSalir", , "Salir"
          .Add "menuSalir", "opSalir_01", , "Cambiar de Usuario"
          .Add "menuSalir", "opSalir_02", , "Salir de la aplicaci�n"
    End With
    frmMenu.SmartMenuXP1.Font.Name = "Ms Sans Serif"
    frmMenu.SmartMenuXP1.Font.Size = 8
End Sub

Public Sub menuLaboratorio(ID As Integer)
    Select Case ID
        Case 1 ' Recepci�n de Muestra
            frmRecepcion.Show
        Case 2 ' Recepci�n CE
            frmCE_Recepcion.Show
        Case 3 ' Recepci�n Sellante
            frmSE_Recepcion.Show
        Case 4 ' Plantilla
            frmPlantilla.Show
        Case 5 ' Registro de muestras
            Dim oform As New frmListadoMuestras
            oform.Show
            Set oform = Nothing
        Case 6 ' Determinaciones Pendientes
            frmListadoDeterminaciones.Show
        Case 7 ' Trabajo Pendiente
            frmListadoDeterminacionesPendientes.Show
        Case 8 ' Muestras a entregar
            frmTrabajo_Pendiente.Show
        Case 9 ' Probetas CE pendientes
            frmCE_Listado_Probetas.Show
        Case 10 ' Localizador
            frmEtiquetasLocalizador.Show 1
        Case 11 ' Metrohm
            frmMetrohm.Show 1
'        Case 15 ' Recepcion de pintura
'            frmPinturasRecepcionAdministrativaListado.Show
        Case 16 ' Recepci�n de plasma
            frmPlasma_Recepcion.Show
        Case 17 ' Recepci�n de henkel
            frmHenkel_Recepcion.Show
    End Select
End Sub

Public Sub menuInformes(ID As Integer)
    Select Case ID
        Case 1 ' Informe muestras pendientes cierre
            frmInformeMuestrasPendientesCierre.Show
        Case 6 ' Muestras pendientes revision
            frmInformeMuestrasPendientesRevision.Show
        Case 2 '
            frmInformeMuestrasAnalizadasPorClienteFecha.Show
        Case 3 ' Informe de alimentos
            frmInformeRegistro.Show
        Case 4 ' Facturaci�n
            frmInformeFacturacion.Show
        Case 5 ' Duplicados
            frmDuplicados_Informe.Show
        'MANTIS-808-I
        Case 7 ' Muestras pendientes de env�o
            frmInformeMuestrasPendientesEnvio.Show
        'MANTIS-808-F
        Case 8
            frmInformeMuestrasFueraPlazo.Show
        Case 9
            frmInformePedidosClientes.Show
        Case 10
            frmInformeMuestrasAnalizadasPorAnalista.Show
        Case 11
            frmInformeMuestrasAnalizadasPorAnalistaDeterminacion.Show
    End Select
End Sub
Public Sub menuBanos(ID As Integer)
    Select Case ID
        Case 1 ' Hist�rico
            frmEads_Historico.Show
    End Select
End Sub
Public Sub menuFacturacion(ID As Integer)
    gdoc = 0
    Select Case ID
        Case 1 ' Muestras pendientes de facturaci�n
            frmMuestraPendientesFacturacion2.Show
        Case 2 ' Listado documentos de pago
            frmListadoDocPago.Show
        Case 3 ' Listado albaranes pendientes de facturar
            frmListadoAlbaranes.Show
        Case 4 ' Facturaci�n de Conceptos
            frmFacturaConceptos.Show 1
        Case 5 ' Facturaci�n de Alodine
            frmAlodine_Facturacion.Show
        Case 6 ' Contabilidad, exportar a contaplus
            frmContabilidad.Show
        Case 7 ' Contabilidad, importar asientos desde contaplus
            frmContabilidad_Asientos.Show
        Case 8 ' Facturaci�n por sectores
            frmFacturacion_Sectores.Show
        Case 9 ' Facturacion de Suministros
            frmSuministros_Facturacion.Show
        Case 10
            frmPC_Facturacion.Show
        Case 11 ' Informe de muestras y conceptos
            frmFacturacionInformeMuestrasConceptos.Show
        Case 12 ' Facturacion Henkel
            frmFacturacion_henkel.Show
    End Select
End Sub

Public Sub menuIndicadores(ID As Integer)
    Select Case ID
        Case 1 ' Generaci�n
            frmIndicadores_Gestion.Show
        Case 2 ' Indicadores
            frmIndicadores_Lista.Show
        Case 3 ' Campos
            frmIndicadores_Lista_Campos.Show
        Case 4 ' Frecuencias
            frmIndicadores_Frecuencias.Show
        Case 5 ' Funciones
            frmIndicadores_Funciones.Show
        'M1039-I
        Case 6 ' Indicadores
            frmIndicador_Listado.Show
        Case 7 ' Gesti�n
            frmIndicador_Gestion.Show 1
        'M1039-F
    End Select
End Sub
Public Sub menuReactivos(ID As Integer)
    Select Case ID
        Case 1
            frmREX_Gestion.Show
        Case 2
            frmRPR_Gestion.Show
        Case 3
            frmREX_Pedidos_Listado.Show
        Case 4 ' Inventario
            frmREX_Inventario_Listado.Show
        Case 5: 'Informe de pedidos realizados
            frmInformePedidos.Show
        Case 6: 'Informe de pedidos realizados por Fecha
            frmInformePedidosFecha.Show
    End Select
End Sub
Public Sub menuAlodine(ID As Integer)
    Select Case ID
        Case 1 ' Lote
            frmAlodine_Listado_Lotes.Show
        Case 2 ' Tipos de Alodine
            frmAlodine_Listado.Show
        Case 3 ' Capacidades
            frmAlodine_Capacidades.Show 1
        Case 4 ' Tipos de Suministros
            frmSuministros_Listado.Show
        Case 5 ' Lotes de suministros
            frmSuministros_Listado_Lotes.Show
        Case 6 ' Lotes de productos controlados
            frmPC_Listado.Show
        Case 7 ' Listado de Alodine Suministrado
            frmAlodine_Listado_Suministros.Show
    End Select
End Sub
'Public Function pGetPicture(sFileName As String) As StdPicture
'    Set pGetPicture = LoadPicture(App.Path + "\Images\" + sFileName + ".ico")
'End Function
'Public Function pGetPicture(Index As Integer) As StdPicture
'    Set pGetPicture = frmMenu.menus.ListImages(Index).Picture
'End Function

Public Sub menuCalidad(ID As Integer)
    Dim oform As New frmDecodificadora
    Select Case ID
      Case 1
'        frmCA_Listado_Documentos.VINCULAR = False
        Dim ofrmCA As New frmCA_Listado_Documentos2
        ofrmCA.Show
        Set ofrmCA = Nothing
'        frmCA_Listado_Documentos2.Show
      Case 2
        oform.CODIGO = DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS
      Case 3
        oform.CODIGO = DECODIFICADORA.CA_DOCUMENTOS_SUBFAMILIAS
      Case 4
        oform.CODIGO = DECODIFICADORA.CA_DOCUMENTOS_RESPONSABLES
      Case 5
        oform.CODIGO = DECODIFICADORA.CA_DOCUMENTOS_ESTADOS
      Case 6
        ' E0501-I
        'Dim Ret As Long
        'Ret = SetParent(frmCA_Listado_Normas.hWnd, frmMenu.hWnd)
        frmCA_Listado_Normas.VINCULAR = False
'        frmCA_Listado_Normas.Show 1
        frmCA_Listado_Normas.Show
        ' E0501-F
      Case 7
        oform.CODIGO = DECODIFICADORA.CA_NORMAS_TIPOS
      Case 8
        oform.CODIGO = DECODIFICADORA.CA_NORMAS_SECTORES
      Case 9
        oform.CODIGO = DECODIFICADORA.CA_NORMAS_ESTADOS
      Case 10
        oform.CODIGO = DECODIFICADORA.CA_NORMAS_SUBTIPOS
      Case 11
        frmOferta_Listado.Show
      Case 12 ' Listado NC
        frmNC_Listado.Show
'      Case 40 ' AUDITORIAS
'        oform.CODIGO = DECODIFICADORA.PROCNC_AUDITORIAS
      Case 41 ' AUDITORIAS INTERNA
        oform.CODIGO = DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_INTERNA
      Case 42 ' AUDITORIAS EXTERNA
        oform.CODIGO = DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_EXTERNA
      Case 43 ' DETECCION INTERNA
        oform.CODIGO = DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_DETECCION
      Case 13 ' TIPOS DE HECHOS
        oform.CODIGO = DECODIFICADORA.NC_TIPOS_HECHOS
      Case 14 ' ORIGENES
        oform.CODIGO = DECODIFICADORA.NC_ORIGENES
      Case 15 ' ESTADOS
        oform.CODIGO = DECODIFICADORA.NC_ESTADOS
      Case 16 ' Departamentos
        oform.CODIGO = DECODIFICADORA.NC_DEPARTAMENTOS
      Case 17 ' Afectado
        oform.CODIGO = DECODIFICADORA.NC_AFECTADO
      Case 18 ' Auditoria - Programas
        frmAU_Programa_Listado.Show
      Case 19 ' Auditoria - Areas
        frmAU_Areas_Listado.Show
      Case 20 ' Procedimientos No Conformidades
        Set objFrmProcNoConf = New frmProcNC_Listado
        objFrmProcNoConf.Show
      Case 21 ' Acciones Correctivas Pendientes
        frmProcNC_AvisosAccCorrectivas.Show
'        Dim objfrm As New frmProcNC_AvisosAccCorrectivas
'        Load objfrm
'        If objfrm.SinAvisos Then
'            MsgBox "No existen Acciones Correctivas Pendientes", vbInformation, "Acciones Correctivas Pendientes"
'            Unload objfrm
'            Set objfrm = Nothing
'        Else
'            objfrm.Show
'        End If
      Case 22 ' Desviaciones
        oform.CODIGO = DECODIFICADORA.NC_DESVIACIONES
      Case 26 ' Plan de formaci�n
        frmFormacion_PlanAnual_Detalle.Show
      Case 27 ' Plan de formaci�n
        frmFormacion_PlanAnual_Listado.Show
      Case 30
        frmEmpleados_Matriz.Show
      Case 31
        frmCA_Listado_Req.Show
    End Select
    Select Case ID
     Case 2, 3, 4, 5, 7, 8, 9, 10, 13, 14, 15, 16, 17, 22, 40, 41, 42, 43
        oform.Show
    End Select
    Set oform = Nothing
End Sub
Public Sub menuRRHH(ID As Integer)
    Select Case ID
      Case 1
        frmEmpleados_Listado.Show
      Case 2
        frmEmpleados_Categorias.Show
'      Case 3
'        frmEmpleados_Nominas_Gestion.Show
    End Select
End Sub


Public Sub menuEquipos(ID As Integer)
    Dim oform As New frmDecodificadora
    Dim objfrm As New frmEquipoListado
    
    
    Dim ret As Long
    Select Case ID
        Case 1 ' LISTADO
           'frmEquipos_Listado.Show
           objfrm.Show vbModeless
           Set objfrm = Nothing
           
        Case 2 ' MTO FAMILIAS
        
            'oform.codigo = decodificadora.EQ_FAMILIAS
            Dim objFamFrm As New frmListadoFamiliasEquipos
            objFamFrm.Show vbModeless
            Set objFamFrm = Nothing
            
        'E0309-I
        Case 3 ' MTO SITUACIONES
            oform.CODIGO = DECODIFICADORA.EQ_SITUACIONES
        Case 4 ' MTO PERIODICIDADES
            oform.CODIGO = DECODIFICADORA.EQ_periodicidad
        'E0309-F
        Case 5 ' MTO Planes de Mantenimiento
            ' Aunque no sea MDICHILD, la ventana padre es frmenu (as� se queda dentro de esa ventana)
            'frmEquipos_PlanesMantenimiento_listado.Show
            ret = SetParent(frmEquipos_PlanesMantenimiento_listado.Hwnd, frmMenu.Hwnd)
            frmEquipos_PlanesMantenimiento_listado.Show
        Case 6 ' MTO Acciones (de los planes de mantenimiento)
            ' Aunque no sea MDICHILD, la ventana padre es frmenu (as� se queda dentro de esa ventana)
            'frmEquipos_Planes_Acciones.Show
            ret = SetParent(frmEquipos_Planes_Acciones.Hwnd, frmMenu.Hwnd)
            frmEquipos_Planes_Acciones.Show
        Case 7 ' MTO FAMILIAS DE ACCIONES DE LOS PLANES DE MTO
            oform.CODIGO = DECODIFICADORA.EQ_FAMILIAS_ACCIONES_PLANES_MTO
        'E0500-I
        Case 8 ' Avisos calibraci�n, verificaci�n y mantenimiento
            'frmEquipos_Listado_Avisos.Show
            oform.CODIGO = DECODIFICADORA.EQ_TIPOS_EQUIPO
        'E0500-F
        Case 9 ' Parametros T�cnicos
            Dim frmPT As New frmEquipoParametrosTecnicos
            frmPT.Show
            Set frmPT = Nothing
    End Select
    Select Case ID
     Case 3, 4, 7, 8
        oform.Show
    End Select
End Sub
' COMPRAS
' Se quitan todos los frmSO del proyecto
'Public Sub menuCompras(ID As Integer)
'    Dim oform As New frmDecodificadora
'    Dim objfrm As Object
'    Dim Ret As Long
'    Select Case ID
'        Case 1 ' Equipos
'            Set objfrm = New frmSOEquipos
'        Case 2 ' Calibraciones
'            Set objfrm = New frmSOCalibracion
'        Case 3 ' Patrones
'            Set objfrm = New frmSOPatrones
'        Case 4 ' Reactivos
'            Set objfrm = New frmSOReactivos
'        Case 5 ' Productos Controlados
'            Set objfrm = New frmSOProdControlados
'        Case 6 ' Material Oficina
'            Set objfrm = New frmSOMatOficina
'        Case 7 ' Estructurales
'            Set objfrm = New frmSOEstructurales
'    End Select
'
'    objfrm.Show
'    Set objfrm = Nothing
'
'End Sub
' COMPRAS

Public Sub menuEnvios(ID As Integer)
    Dim oform As New frmDecodificadora
    Select Case ID
    'M1241-I
'        Case 1: ' Ensayos que se subcontratan
'            frmSC_Ensayos_subcontratan_listado.Show
'            frmSC_Listado.Show
    'M1241-F
        Case 2: ' Env�o de paquetes
            frmEP_Listado.Show
        Case 3: ' Empresas mensajer�a
            oform.CODIGO = DECODIFICADORA.EP_EMPRESAS_MENSAJERIA
            oform.Show
    End Select
End Sub
'M1241-I
Public Sub menuPedidos(ID As Integer)
    Select Case ID
        Case 1: ' Ensayos que se subcontratan
            frmSC_Listado.Show
        Case 2: ' Env�o de Pedidos
            frmPP_Listado.Show
    End Select
End Sub
'M1241-F

Public Sub menuTablets(ID As Integer)
    Select Case ID
        Case 1: ' Ensayos que se subcontratan
            frmListadoMuestrasTablet.Show 1
    End Select
End Sub
Public Sub menuMantenimiento(ID As Integer)
    Dim oform As New frmDecodificadora
    Select Case ID
        Case 1 ' Tipos de Muestras
            frmTM_Listado.Show
        Case 2 ' Tipos de Muestras
            frmTA_Listado.Show
        Case 3 ' Tipos de dETER
            frmTD_Listado.Show
        Case 4 ' Tipos de dependencias
            frmDD_Listado.Show
        Case 5 ' Datos especificos
            frmTDE_Listado.Show
        Case 6 ' Ba�os
            frmBANO_Listado.Show
        Case 7 ' Lineas
            frmLineas.Show
        Case 8 ' Procesos Base
            frmProcesosBase.Show
        Case 9 ' Soluciones
            frmSoluciones.Show
        Case 10 ' Frecuencias de muestreo
            frmTipos_Frecuencia.Show
        Case 11 ' tipos reactivos externos
            frmREX_Listado.Show
        Case 12 ' Tipos Botes reactivos externos
        frmREX_Botes_Listado.PK_TIPO_REACTIVO_ID = 0
            frmREX_Botes_Listado.Show
        Case 13 ' Tipos Botes reactivos propios
            frmRPR_Listado.Show
        Case 14 ' frmFormulas
            frmFORMULA_Listado.Show
        Case 15 ' Unidades
            frmUnidades.Show
        Case 16 ' Envases
            frmformatos.Show
        Case 17 ' Caducidad
            frmTipos_caducidad.Show 1
        Case 18
            frmCE_Listado_Fichas.Show
        Case 19
            frmCE_Listado_Tipos_ensayo.Show
        Case 20
            frmCE_Listado_Lotes_Probetas.Show
        Case 21 ' Sellantes
            frmSE_Listado.Show
        Case 22 ' Fluidos
            frmFluidos_Listado.Show
        Case 23 ' Listado Clientes
            frmListadoClientes.Show
        Case 24 ' Proveedores
            frmProveedores_Listado.Show
        Case 25 ' Usuarios
            frmUsuarios_Listado.Show
'        Case 26 ' Empleados
'            frmEmpleados_Listado.Show
        Case 27 ' Sectores
            frmSectores.Show
        Case 28 ' Familias
            frmFamilias.Show
        Case 29 ' FP
            frmFP.Show
        Case 30 ' Gesti�n de tarifas
            frmTarifas.Show
        Case 31 ' Listado de tarifas
            frmTarifas_Listado.Show
        Case 32 ' c�digos tarifarios
            frmTarifas_Codigos.Show
        Case 33 ' familias de c�digos
            frmTarifas_Familias.Show
        Case 34 ' Usuarios conectados
            frmUsuarios_Conectados.Show
        Case 35 ' Parametros
            frmParametros.Show
        Case 36 ' About
            frmAbout.Show 1
        Case 37 ' Materiales de Controles de Eficacia
            frmCE_Materiales_Listado.Show
        Case 38 ' Videos
            frmVideos_Listado.Show
        Case 39 ' Dimensiones
            oform.CODIGO = DECODIFICADORA.DECODIFICADORA_DIMENSIONES
'        Case 40 ' Pinturas
'            frmPinturasListado.Show
        Case 41 ' Etiquetas para soluciones
            frmSoluciones_Etiquetas_Listado.Show
        Case 51 ' Plasma -> Procesos
            frmPlasma_Procesos_Listado.Show
        Case 52 ' Plasma -> Ficha
            frmPlasma_Ficha_Listado.Show
        Case 53 ' Plasma -> Ensayos
            frmPlasma_Ensayos_Listado.Show
        Case 54 ' Plasma -> Recubrimientos
            oform.CODIGO = DECODIFICADORA.DECODIFICADORA_PLASMA_RECUBRIMIENTOS
        Case 55 ' Plasma -> Clientes Internos
            oform.CODIGO = DECODIFICADORA.DECODIFICADORA_PLASMA_CLIENTES_INTERNOS
        Case 56 ' Plasma -> Fabricantes
            oform.CODIGO = DECODIFICADORA.DECODIFICADORA_PLASMA_FABRICANTES
        Case 57 ' Plasma -> Product Type
            oform.CODIGO = DECODIFICADORA.DECODIFICADORA_PLASMA_PRODUCT_TYPE
        Case 58 ' Plasma -> Product Type
            oform.CODIGO = DECODIFICADORA.DECODIFICADORA_PLASMA_NUMBER_AND_TYPE
        Case 60 ' Ba�os -> Instalaciones
            oform.CODIGO = DECODIFICADORA.BANOS_INSTALACIONES
        Case 61 ' Descripciones de productos
            frmDescripcionesProducto.Show
        Case 62 ' Inventario
            frmInventario_Listado.Show
        Case 70 ' Plantas
            oform.CODIGO = DECODIFICADORA.AIRBUS_PLANT
        Case 71 ' Airbus
            frmAirbus_Decodificadora.Show
    End Select
    Select Case ID
     Case 39, 54, 55, 56, 57, 58, 60, 70
        oform.Show
    End Select
End Sub

'M0996-I
Public Sub menuFormacion(ID As Integer)
    Select Case ID
        Case 1: ' Formaci�n anual
            frmFormacion_Listado.Show 0
        Case 2: ' Cursos/Documentaci�n
            frmFormacion_PF_Listado.Show 0
        Case 3: 'Plan de formaci�n anual
            frmFormacion_PFA_Listado.Show 0
        Case 4: 'Certificaci�n de formadores
            frmFormacion_CF_Listado.Show 0
    End Select
End Sub
'M0996-F
'TESORERIA-I
Public Sub menuTesoreria(ID As Integer)
    Dim oform As New frmDecodificadora
    Select Case ID
        Case 1: '
            frmProveedores_Facturas_Listado.Show
        Case 2: ' otros gastos
            frmGastos_Listado.Show
        'M1323-I
        Case 3
            frmContabilidad_Proveedores.Show
        Case 4
            oform.CODIGO = DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS
        Case 5
            oform.CODIGO = DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_PAGOS
        'M1323-F
        'M1362-I
        Case 6
            frmContabilidad_Proveedores_Pagos.Show
        'M1362-F
        Case 7
            frmBancos_Listado.Show
        Case 8
            frmSeguros_Listado.Show
        Case 9
'            frmprestamos_Listado.Show

        Case 10
            frmRemesas_Listado.Show
        Case 11
            frmTesoreria.Show
    End Select
    Select Case ID
     Case 4, 5
        oform.Show
    End Select
End Sub
'TESORERIA-F

Public Sub menuSalir(ID As Integer)
'    Dim oform As New frmDecodificadora
    Select Case ID
        Case 1
            cambiar_usuario
        Case 2
            If MsgBox("�Desea Cerrar Geslab - Gesti�n para laboratorios?", vbQuestion + vbOKCancel, App.Title) = vbOK Then
'                Unload frmErrores
'                Unload frmCambioUsuario
'                Unload frmMEN_Nuevo2
'                Unload frmTareas_Incurrir
'                Unload frmTecladoNumerico
'JGM                Unload frmMenu
                Salir
'                MDIForm_Unload
'                End
            End If
    End Select
End Sub

'Public Sub menuJonathan(ID As Integer)
'
'frm_pruebas_jonathan.Show 1
'
'End Sub

Public Sub cambiar_usuario()
'    If UCase(usuario.getUSUARIO) = "PRUEBA" Then
    If MODO_PRUEBA Then
        MsgBox "En PRUEBA no se puede cambiar de usuario.", vbInformation, App.Title
    Else
        USUARIO.deslogonear (USUARIO.getID_EMPLEADO)
        glogin = 1
        frmLogin.Show 1
        frmMenu.inicializa_ventana
    End If
End Sub

Public Sub barra_vertical(opcion As Integer, subopcion As Integer)
    Select Case (opcion * 10) + subopcion
        ' REGISTRO
        Case 11: menuLaboratorio 1
        Case 12: menuLaboratorio 16
        Case 13: menuLaboratorio 5
        Case 14: menuLaboratorio 4
        Case 15: menuLaboratorio 6
        Case 16: menuLaboratorio 7
        Case 17: menuLaboratorio 8
        Case 18: menuLaboratorio 9
        Case 19: menuLaboratorio 10
        Case 20: menuLaboratorio 11 'Metrohm
        ' OFICINA
        Case 21: menuMantenimiento 23
        Case 22: menuMantenimiento 24
        Case 23: frmListadoAgenda.Show
        Case 24: menuFacturacion 2
        ' LABORATORIO
        Case 31: menuMantenimiento 1 ' TM
        Case 32: menuMantenimiento 2 ' TA
        Case 33: menuMantenimiento 3 ' TD
        Case 34: menuMantenimiento 14 ' Formulas
        Case 35: menuMantenimiento 6 ' Ba�os
        Case 36: menuMantenimiento 18 ' CE
        Case 37: menuMantenimiento 21 ' Sellantes
        Case 38: menuMantenimiento 22 ' Fluidos
        ' OPCIONES
        Case 41: menuSalir 1
        Case 42: menuSalir 2
        ' TABLET
        Case 51: menuTablets 1
    End Select

End Sub
