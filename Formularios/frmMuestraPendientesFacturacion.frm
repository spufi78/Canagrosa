VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmMuestraPendientesFacturacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Muestras pendientes de facturacion"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   120
   ClientWidth     =   13950
   Icon            =   "frmMuestraPendientesFacturacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   13950
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   330
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7110
      Width           =   1410
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   330
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7110
      Width           =   1410
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   12735
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7695
      Width           =   1140
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5010
      Left            =   45
      TabIndex        =   14
      Top             =   2070
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   8837
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   45
      TabIndex        =   12
      Top             =   7470
      Width           =   12615
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Analisis de muestras y Leyenda"
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   4950
         TabIndex        =   30
         Top             =   135
         Width           =   3885
         Begin VB.CommandButton cmdAnalizar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Analizar"
            Height          =   780
            Left            =   2700
            Picture         =   "frmMuestraPendientesFacturacion.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Previsualizar como quedaría la muestra en la factura"
            Top             =   135
            Width           =   1140
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Verde, precio no coincide tarifa"
            ForeColor       =   &H00008000&
            Height          =   240
            Index           =   1
            Left            =   225
            TabIndex        =   34
            Top             =   675
            Width           =   2310
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Azul, muestras con precio 0,00"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   33
            Top             =   450
            Width           =   2310
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Rojo Revisar Facturación"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   225
            TabIndex        =   32
            Top             =   225
            Width           =   2085
         End
      End
      Begin VB.CommandButton cmdprevia 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Previsualizar"
         Height          =   825
         Left            =   3690
         Picture         =   "frmMuestraPendientesFacturacion.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Previsualizar como quedaría la muestra en la factura"
         Top             =   270
         Width           =   1140
      End
      Begin VB.CommandButton cmdbano 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Baño/Analisis"
         Height          =   825
         Left            =   2520
         Picture         =   "frmMuestraPendientesFacturacion.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Recalcula el precio de la muestra seleccionada"
         Top             =   270
         Width           =   1140
      End
      Begin VB.CommandButton cmdlog 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Log"
         Height          =   240
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   45
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdListado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Listado"
         Height          =   825
         Left            =   11385
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton cmdrec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recalcular Seleccionada"
         Height          =   825
         Left            =   1350
         Picture         =   "frmMuestraPendientesFacturacion.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Recalcula el precio de la muestra seleccionada"
         Top             =   270
         Width           =   1140
      End
      Begin VB.CommandButton cmdRecalculo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recalcular todas"
         Height          =   825
         Left            =   180
         Picture         =   "frmMuestraPendientesFacturacion.frx":2BF2
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Recalcula el precio de todas las muestras pendientes de facturar"
         Top             =   270
         Width           =   1140
      End
      Begin VB.CommandButton cmdAlbaran 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Crear &Albaran"
         Enabled         =   0   'False
         Height          =   825
         Left            =   10170
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMuestraPendientesFacturacion.frx":34BC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton cmdFactura 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Crear &Factura"
         Enabled         =   0   'False
         Height          =   825
         Left            =   8955
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMuestraPendientesFacturacion.frx":40FE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   45
      TabIndex        =   6
      Top             =   315
      Width           =   13860
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Asociar pedido a las muestras marcadas"
         Height          =   315
         Left            =   8505
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1080
         Width           =   2985
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   375
         Left            =   1080
         TabIndex        =   22
         Top             =   300
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   661
      End
      Begin VB.CommandButton cmdiniciar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   855
         Left            =   12690
         Picture         =   "frmMuestraPendientesFacturacion.frx":4D40
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   1035
      End
      Begin VB.CheckBox chkCerradas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Muestras abiertas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4770
         TabIndex        =   13
         Top             =   720
         Width           =   2310
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos Clientes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10035
         TabIndex        =   0
         Top             =   315
         Width           =   1440
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   855
         Left            =   11565
         Picture         =   "frmMuestraPendientesFacturacion.frx":560A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker DTPicker_desde 
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   675
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   12632256
         Format          =   53346305
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker DTPicker_hasta 
         Height          =   330
         Left            =   3105
         TabIndex        =   2
         Top             =   675
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   12632256
         Format          =   53346305
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbPedidos 
         Bindings        =   "frmMuestraPendientesFacturacion.frx":5ED4
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   1080
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   20
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   2475
         TabIndex        =   9
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   600
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre el análisis para ver el detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   5175
      TabIndex        =   15
      Top             =   7155
      Width           =   4095
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione Criterio."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      TabIndex        =   11
      Top             =   1800
      Width           =   13905
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Muestras pendientes de facturación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   45
      TabIndex        =   10
      Top             =   0
      Width           =   13875
   End
End
Attribute VB_Name = "frmMuestraPendientesFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_cTT As New cTooltip

Private Sub cmdAnalizar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a realizar un análisis de las muestras de la lista para verificar posibles errores, ¿esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim i As Integer
        Dim omuestra As New clsMuestra
        Dim oBANO As New clsBanos
        Dim oTA As New clsTipos_analisis
        Dim oDeterminacion As New clsDeterminaciones
        Dim oTarifa As New clsTarifas_precios
        Dim ocliente As New clsCliente
        Dim color As Long
        ' ROJO: Muestras con TA, BANO o DETERMICION con check de REVISAR_FACTURA
        ' AZUL: Muestras con precio 0
        ' VERDE: Muestras con precio distinto al de su código de tarifa
        For i = 1 To lista.ListItems.Count
            color = 0
            If omuestra.CargaMuestra(lista.ListItems(i).SubItems(7)) Then
                ' Analizar check de REVISAR_FACTURACION (TA y BANO)
                If omuestra.getBANO_ID = 0 Or omuestra.getANALISIS_MODIFICADO = 2 Then
                    If oTA.revisar_facturacion(omuestra.getTIPO_ANALISIS_ID) Then
                        color = vbRed
                    End If
                Else
                    If oBANO.revisar_facturacion(omuestra.getBANO_ID) Then
                        color = vbRed
                    End If
                End If
                ' Analizar check de REVISAR_FACTURACION (DETERMINACIONES)
                If color = 0 Then
                    If oDeterminacion.revisar_facturacion(lista.ListItems(i).SubItems(7)) Then
                        color = vbRed
                    End If
                End If
                ' Analizar precio de la muestra
                If color = 0 Then
                    If CCur(lista.ListItems(i).SubItems(5)) = 0 Then
                        color = vbBlue
                    End If
                End If
                ' Analizar precio de la tarifa
                If color = 0 Then
                    If lista.ListItems(i).SubItems(10) = 0 Then ' No factura determinaciones
                       ocliente.CargaCliente lista.ListItems(i).SubItems(8)
                       If lista.ListItems(i).SubItems(11) <> 0 And lista.ListItems(i).SubItems(12) <> 2 Then
'                          oBANO.cargar_bano CLng(lista.ListItems(I).SubItems(11))
'                          oTarifa.Carga oBANO.getTARIFA_CODIGO_ID
                          oTarifa.Carga_por_BANO CLng(lista.ListItems(i).SubItems(11)), ocliente.getTARIFA_ID
                       Else
'                          oTA.cargar lista.ListItems(lista.SelectedItem.Index).SubItems(13)
'                          oTarifa.Carga oTA.getTARIFA_CODIGO_ID
                          oTarifa.Carga_por_TA lista.ListItems(i).SubItems(13), ocliente.getTARIFA_ID
                       End If
                       If CCur(lista.ListItems(i).SubItems(5)) <> CCur(oTarifa.getPRECIO) Then
                          color = &H8000&
                       End If
                    End If
                End If
            End If
            If color <> 0 Then
                colorear i, color
            End If
            lista.ListItems(i).EnsureVisible
            DoEvents
        Next
        MsgBox "Análisis finalizado.", vbInformation, App.Title
        lista.Refresh
    End If
End Sub

Public Sub cmdbano_Click()
   On Error GoTo cmdbano_Click_Error

        If lista.ListItems.Count = 0 Then
            Exit Sub
        End If
        Dim omuestra As New clsMuestra
        If omuestra.CargaMuestra(lista.ListItems(lista.SelectedItem.Index).SubItems(7)) Then
            If omuestra.getBANO_ID = 0 Or omuestra.getANALISIS_MODIFICADO = 2 Then
                frmTA_Detalle.PK = omuestra.getTIPO_ANALISIS_ID
                frmTA_Detalle.Show 1
            Else
                frmBANO_Detalle.PK = omuestra.getBANO_ID
                frmBANO_Detalle.Show 1
            End If
'            cmdrec_Click
        End If
        lista_Click
   On Error GoTo 0
   Exit Sub

cmdbano_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdbano_Click of Formulario frmMuestraPendientesFacturacion"
End Sub

Private Sub cmdlog_Click()
        On Error GoTo fallo
        If lista.ListItems.Count = 0 Then
            Exit Sub
        End If
        Dim men As String
        Dim total As Currency
        Dim odd As New clsDeterminaciones_analisis
        Dim ocliente As New clsCliente
        Dim oTarifa As New clsTarifas_precios
        Dim omuestra As New clsMuestra
        If omuestra.CargaMuestra(lista.ListItems(lista.SelectedItem.Index).SubItems(7)) Then
            ocliente.CargaCliente (omuestra.getCLIENTE_ID)
            Dim otar As New clsTarifas
            otar.Carga ocliente.getTARIFA_ID
            men = "TARIFA DEL CLIENTE : " & otar.getNOMBRE & " (" & ocliente.getTARIFA_ID & ")" & vbNewLine
            ' Factura por determinaciones
            If lista.ListItems(lista.SelectedItem.Index).SubItems(10) = 1 Then
                men = men & "El cliente o el análisis se factura por DETERMINACIONES" & vbNewLine
                Dim consulta As String
                consulta = " SELECT td.nombre,tp.PRECIO " & _
                           "  FROM determinaciones d,tipos_determinacion td " & _
                           "  LEFT JOIN tarifas_precios tp on tp.tipo_determinacion_id = td.id_tipo_determinacion " & _
                           " WHERE d.tipo_determinacion_id = td.id_tipo_determinacion" & _
                           "   AND tp.tarifa_id = " & ocliente.getTARIFA_ID & _
                           "   AND d.muestra_id=" & lista.ListItems(lista.SelectedItem.Index).SubItems(7)
                Dim rs As ADODB.RecordSet
                Set rs = datos_bd(consulta)
                men = men & "-------------------------------------" & vbNewLine
                If rs.RecordCount > 0 Then
                    Do
                        men = men & rs(0) & " (Precio : " & moneda(rs(1)) & ")" & vbNewLine
                        rs.MoveNext
                    Loop Until rs.EOF
                Else
                    men = men & "Las determinaciones no tienen el precio introducido." & vbNewLine
                End If
                men = men & "-------------------------------------" & vbNewLine
                men = men & "Precio TOTAL : " & Format(omuestra.ImporteMuestraPorDeterminaciones(lista.ListItems(lista.SelectedItem.Index).SubItems(7), lista.ListItems(lista.SelectedItem.Index).SubItems(8)), "currency")
            Else
            ' No factura por determinaciones
            ' Recuperamos el precio del analisis o bano por tarifa
            ' Miramos si se factura por tipo analisis o control de eficacia
                If omuestra.getBANO_ID = 0 Or omuestra.getANALISIS_MODIFICADO = 2 Then
                    Dim oTA As New clsTipos_analisis
                    oTA.CARGAR omuestra.getTIPO_ANALISIS_ID
                    If omuestra.getANALISIS_MODIFICADO = 2 Then
                        men = men & "CONTROL DEL EFICACIA : " & oTA.getNOMBRE & " (" & omuestra.getTIPO_ANALISIS_ID & ")" & vbNewLine
                    Else
                        men = men & "TIPO DE ANÁLISIS : " & oTA.getNOMBRE & " (" & omuestra.getTIPO_ANALISIS_ID & ")" & vbNewLine
                    End If
                    If oTarifa.Carga_por_TA(omuestra.getTIPO_ANALISIS_ID, ocliente.getTARIFA_ID) Then
                        total = CCur(Replace(oTarifa.getPRECIO, ".", ","))
                        men = men & "PRECIO DEL ANÁLISIS PARA LA TARIFA : " & Format(total, "CURRENCY") & vbNewLine
                    End If
                Else
                    Dim oBANO As New clsBanos
                    oBANO.cargar_bano (omuestra.getBANO_ID)
                    men = men & "BAÑO : " & oBANO.getNOMBRE & " (" & omuestra.getBANO_ID & ")" & vbNewLine
                    If oTarifa.Carga_por_BANO(omuestra.getBANO_ID, ocliente.getTARIFA_ID) Then
                        total = CCur(Replace(oTarifa.getPRECIO, ".", ","))
                        men = men & "PRECIO DEL BAÑO PARA LA TARIFA : " & Format(total, "CURRENCY") & vbNewLine
                    End If
                End If
                ' Recuperamos los datos por defecto del analisis o bano
                If omuestra.getBANO_ID = 0 Then
                    men = men & "PRECIO POR DETERMINACIONES : " & Format(odd.Precio_determinaciones_por_tipo_analisis(omuestra.getTIPO_ANALISIS_ID, lista.ListItems(lista.SelectedItem.Index).SubItems(7), ocliente.getTARIFA_ID), "CURRENCY") & vbNewLine
                    total = total + odd.Precio_determinaciones_por_tipo_analisis(omuestra.getTIPO_ANALISIS_ID, lista.ListItems(lista.SelectedItem.Index).SubItems(7), ocliente.getTARIFA_ID)
                Else
                    men = men & "PRECIO POR DETERMINACIONES : " & Format(odd.Precio_determinaciones_por_bano(omuestra.getBANO_ID, lista.ListItems(lista.SelectedItem.Index).SubItems(7), ocliente.getTARIFA_ID), "CURRENCY") & vbNewLine
                    total = total + odd.Precio_determinaciones_por_bano(omuestra.getBANO_ID, lista.ListItems(lista.SelectedItem.Index).SubItems(7), ocliente.getTARIFA_ID)
                End If
                men = men & "PRECIO TOTAL MUESTRA : " & Format(total, "CURRENCY") & vbNewLine
            End If
        ' Actualizamos el precio de la muestra
'        omuestra.actualizar_precio MUESTRA, Replace(total, ",", ".")
        m_cTT.ToolText(lista) = men
'        MsgBox men
        End If
    Set odd = Nothing
'    Set oDeter = Nothing
    Set omuestra = Nothing
    Exit Sub
fallo:
    MsgBox "Error al obtener el tipo de documento de la muestra.", vbCritical, Err.Description

End Sub

Public Sub cmdprevia_Click()
    If lista.ListItems.Count > 0 Then
        frmDocumento_Previo_Facturacion.PK_MUESTRA = lista.ListItems(lista.SelectedItem.Index).SubItems(7)
        frmDocumento_Previo_Facturacion.FACTURA_DETERMINACIONES = lista.ListItems(lista.SelectedItem.Index).SubItems(10)
        frmDocumento_Previo_Facturacion.Show 1
        buscar CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(7)), lista.SelectedItem.Index
    End If
End Sub

Public Sub cmdrec_Click()
    If lista.ListItems.Count > 0 Then
        Dim omuestra As New clsMuestra
        omuestra.informar_precio_muestra (lista.ListItems(lista.SelectedItem.Index).SubItems(7))
        buscar CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(7)), lista.SelectedItem.Index
    End If
End Sub

Private Sub cmdRecalculo_Click()
    If MsgBox("¿Esta seguro de recalcular el precio de las muestras sin facturar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        Dim omuestra As New clsMuestra
        If omuestra.recalcular_precios_muestras_sin_facturar Then
            Me.MousePointer = 0
            MsgBox "Se han recalculado los precios correctamente.", vbInformation, App.Title
            cmdBuscar_Click
        End If
        Me.MousePointer = 0
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.value = vbChecked Then
        cmbclientes.Limpiar
        cmbclientes.desactivar
    Else
        cmbclientes.activar
    End If
End Sub

Private Sub cmbClientes_change()
    cmbPedidos.Text = ""
    If cmbclientes.getPK_SALIDA <> 0 Then
       pedidos (cmbclientes.getPK_SALIDA)
    End If
End Sub

Private Sub cmdAlbaran_Click()
    Dim strcadena As String
    If contar_marcados = 0 Then
         MsgBox "Debe seleccionar alguna muestra", vbInformation, App.Title
         Exit Sub
    End If
    If contar_marcados = 1 Then
        strcadena = "Va a generar un albaran a 1 muestra. ¿Desea continuar?"
    Else
        strcadena = "Va a generar albaranes a " & contar_marcados & " muestras. ¿Desea continuar?"
    End If
    If MsgBox(strcadena, vbQuestion + vbYesNo, App.Title) = vbYes Then
         generar_documentos (1) ' Factura
    End If
    cmdBuscar_Click
End Sub

Private Sub cmdBuscar_Click()
   buscar 0, 0
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdFactura_Click()
    Dim strcadena As String
    If contar_marcados = 0 Then
        MsgBox "Debe seleccionar alguna muestra", vbInformation, App.Title
        Exit Sub
    End If
    If contar_marcados = 1 Then
        strcadena = "Va a facturar 1 muestra. ¿Desea continuar?"
    Else
        strcadena = "Va a facturar " & contar_marcados & " muestras. ¿Desea continuar?"
    End If
    If MsgBox(strcadena, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        generar_documentos (2) ' Factura
        Me.MousePointer = 0
    End If
    Call cmdBuscar_Click
End Sub

Private Sub cmdiniciar_Click()
    cmbclientes.Limpiar
    chkTodos.value = Unchecked
    chkCerradas.value = Unchecked
    cmbPedidos.BoundText = ""
    DTPicker_desde.value = Date
    DTPicker_hasta.value = Date
    lista.ListItems.Clear
End Sub

Private Sub cmdListado_Click()
    Dim total As Currency
    Dim i As Integer
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        MsgBox "No existen registros para generar el listado.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim rs As New ADODB.RecordSet
    rs.Fields.Append "c1", adChar, 5, adFldUpdatable
    rs.Fields.Append "c2", adChar, 50, adFldUpdatable
    rs.Fields.Append "c3", adChar, 50, adFldUpdatable
    rs.Fields.Append "c4", adChar, 12, adFldUpdatable
    rs.Open
    total = 0
    For i = 1 To lista.ListItems.Count
        rs.AddNew
        rs("c1") = lista.ListItems(i).SubItems(6)
        rs("c2") = Left(lista.ListItems(i).SubItems(1), 50)
        rs("c3") = Left(lista.ListItems(i) & " " & lista.ListItems(i).SubItems(2), 50)
        rs("c4") = lista.ListItems(i).SubItems(5)
        If Trim(lista.ListItems(i).SubItems(5)) <> "" Then
            total = total + Format(lista.ListItems(i).SubItems(5), "currency")
        End If
        rs.Update
    Next
    ' Generar Listado
    Dim Listado As New dataListadoMuestrasPendientes
    ' Cabecera
    With Listado.Sections("cabecera")
        .Controls("lbltitulo").Caption = "Análisis pendientes de facturar del " & Format(DTPicker_desde, "dd/mm/yyyy") & " al " & Format(DTPicker_hasta, "dd/mm/yyyy")
        If chkTodos.value = Checked Then
            .Controls("lblcliente").Caption = "Cliente : *** TODOS ***"
        Else
            Dim ocliente As New clsCliente
            ocliente.CargaCliente cmbclientes.getPK_SALIDA
            .Controls("lblcliente").Caption = "Cliente : " & ocliente.getNOMBRE
            
        End If
    End With
    Set Listado.Sections("cabecera").Controls("logo").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
    'Detalle
    With Listado.Sections("detalle")
        .Controls("c1").DataField = rs.Fields("c1").Name
        .Controls("c2").DataField = rs.Fields("c2").Name
        .Controls("c3").DataField = rs.Fields("c3").Name
        .Controls("c4").DataField = rs.Fields("c4").Name
    End With
    ' Pie de Pagina
    With Listado.Sections("pie")
        .Controls("lbltotal").Caption = Format(total, "currency")
    End With
    Set Listado.DataSource = rs
    Listado.Caption = "Listado de Análisis Pendientes"
    Listado.WindowState = vbNormal
    Listado.Show
    Set rs = Nothing
'    Me.Height = 7890
'    Me.Width = 12780
    Exit Sub
fallo:
    MsgBox "Error al generar el listado de Analisis pendientes.", vbCritical, Err.Description
End Sub

Private Sub Command1_Click()
    If cmbPedidos.Text = "" Then
        MsgBox "Seleccione un pedido para asociar.", vbExclamation, App.Title
    Else
        Dim Msg As Boolean
        Msg = False
        Dim i As Integer
        Dim omuestra As New clsMuestra
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                omuestra.Informar_Pedido lista.ListItems(i).SubItems(7), cmbPedidos.BoundText
                Msg = True
            End If
        Next
        If Msg Then
            MsgBox "Se han informado correctamente los pedidos.", vbInformation, App.Title
            cmbPedidos.Text = ""
        End If
    End If
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.Top = 50
    cargar_combo_clientes
'    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    
    pedidos (0)
    cabecera_grid
    DTPicker_hasta = Now
    DTPicker_desde = Now
'    If UCase(USUARIO.getNOMBRE) = "JULIO" Then
'        cmdlog.Visible = True
'    End If
    tool
End Sub
Private Sub buscar(MUESTRA_ID As Long, linea As Long)  'obtengo el listado de las muestras pendientes de facturacion para el cliente seleccionado y lo vuelco en el listbox
    On Error GoTo fallo
    Dim rs As ADODB.RecordSet
    Dim omuestra As New clsMuestra
    Dim cliente As Long
    If chkTodos.value = 0 And cmbclientes.getTEXTO = "" Then
        MsgBox "Debe seleccionar un cliente.", vbExclamation, App.Title
        Exit Sub
    End If
    Me.MousePointer = 11
    If MUESTRA_ID = 0 Then
        lista.ListItems.Clear
    End If
    If cmbclientes.getTEXTO <> "" Then
        cliente = cmbclientes.getPK_SALIDA
    End If
    Dim PEDIDO As Long
    If cmbPedidos.Text <> "" Then
        PEDIDO = CLng(cmbPedidos.BoundText)
    End If
    Set rs = omuestra.Muestras_pendientes_facturar(MUESTRA_ID, DTPicker_desde.value, DTPicker_hasta.value, cliente, chkCerradas.value, PEDIDO, True)
    If rs.RecordCount > 0 Then
        If MUESTRA_ID = 0 Then
            While Not rs.EOF
                With lista.ListItems.Add(, , rs.Fields(1))
                .SubItems(1) = rs.Fields(2)
                .SubItems(2) = rs.Fields(10)
                .SubItems(3) = rs.Fields(4)
                If Not IsNull(rs.Fields(5)) Then
                    .SubItems(4) = rs.Fields(5)
                End If
                '** rs(9) El cliente factura por determinaciones '** rs(13) El tipo de analisis es por determinaciones
                If rs(9) = 1 Or rs(13) = 1 Then
                    .SubItems(5) = Format(omuestra.ImporteMuestraPorDeterminaciones(rs(8), rs(0)), "currency")
                    .SubItems(10) = 1
                Else
                    If Not IsNull(rs.Fields(7)) Then
                        .SubItems(5) = Format(rs.Fields(7), "currency")
                    End If
                    .SubItems(10) = 0
                End If
                If Not IsNull(rs.Fields(6)) Then
                    .SubItems(6) = Format(rs.Fields(6), "00000")
                End If
                If Not IsNull(rs.Fields(8)) Then
                    .SubItems(7) = rs.Fields(8)
                End If
                .SubItems(8) = rs.Fields(0)
                .SubItems(9) = rs.Fields(12)
                .SubItems(11) = rs(14) ' BANO_ID
                .SubItems(12) = rs(15) ' ANALISIS_MODIFICADO
                .SubItems(13) = rs(3)  ' TIPO_ANALISIS_ID
                End With
                rs.MoveNext
            Wend
            lblmsg.Caption = "Muestras no facturadas entre el " & Format(DTPicker_desde, "dd/mm/yyyy") & " y " & Format(DTPicker_hasta, "dd/mm/yyyy")
            cmdFactura.Enabled = True
            cmdAlbaran.Enabled = True
        Else
            lista.ListItems(linea).Text = rs.Fields(1)
            lista.ListItems(linea).SubItems(1) = rs.Fields(2)
            lista.ListItems(linea).SubItems(2) = rs.Fields(10)
            lista.ListItems(linea).SubItems(3) = rs.Fields(4)
            If Not IsNull(rs.Fields(5)) Then
               lista.ListItems(linea).SubItems(4) = rs.Fields(5)
            End If
            If rs(9) = 1 Or rs(13) = 1 Then
                lista.ListItems(linea).SubItems(5) = Format(omuestra.ImporteMuestraPorDeterminaciones(rs(8), rs(0)), "currency")
                lista.ListItems(linea).SubItems(10) = 1
            Else
                lista.ListItems(linea).SubItems(10) = 0
                If Not IsNull(rs.Fields(7)) Then
                   lista.ListItems(linea).SubItems(5) = Format(rs.Fields(7), "currency")
                End If
            End If
            If Not IsNull(rs.Fields(6)) Then
                   lista.ListItems(linea).SubItems(6) = Format(rs.Fields(6), "00000")
            End If
            If Not IsNull(rs.Fields(8)) Then
                   lista.ListItems(linea).SubItems(7) = rs.Fields(8)
            End If
            lista.ListItems(linea).SubItems(8) = rs.Fields(0)
            lista.ListItems(linea).SubItems(9) = rs.Fields(12)
            lista.ListItems(linea).SubItems(11) = rs(14) ' BANO_ID
            lista.ListItems(linea).SubItems(12) = rs(15) ' ANALISIS_MODIFICADO
            lista.ListItems(linea).SubItems(13) = rs(3)  ' TIPO_ANALISIS_ID
        End If
    Else
        cmdFactura.Enabled = False
        cmdAlbaran.Enabled = False
        lblmsg.Caption = "No existe ninguna muestra por facturar con esos criterios."
    End If
    Dim i As Integer
    If MUESTRA_ID = 0 Then
        For i = 1 To lista.ListItems.Count
            lista.ListItems(i).Checked = True
        Next
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras (frmMuestrasPendientesFacturar).", vbCritical, Err.Description
End Sub

Public Function contar_marcados() As Integer
    Dim i As Integer
    contar_marcados = 0
    For i = 1 To lista.ListItems.Count
       If lista.ListItems(i).Checked = True Then
        contar_marcados = contar_marcados + 1
      End If
    Next
End Function

Public Sub generar_documentos(Tipo_documento As Integer)
    Dim i As Integer
    Dim num_doc As Integer
    Dim cliente_ant As Long
    Dim total_doc As Integer
    total_doc = 0
    cliente_ant = 0
    ReDim documentos_pago(lista.ListItems.Count)
    Dim oDocPago As New clsDocs_pago
    Dim odoc_muestra As New clsDocs_pago_muestras
    Dim omuestra As New clsMuestra
'    Dim oPedido As New clsClientes_pedidos
    'cIVA
    Dim oParametros As New clsParametros
    Dim IVA As Integer
    IVA = recuperaIVA()
    If IVA = 0 Then
        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
        Exit Sub
    End If
    Dim ORDEN As Integer
    ORDEN = 1
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If cliente_ant <> lista.ListItems(i).SubItems(8) Then
                oDocPago.setTIPO = Tipo_documento
                oDocPago.setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
                oDocPago.setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
                oDocPago.setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                oDocPago.setCLIENTE_ID = lista.ListItems(i).SubItems(8)
                oDocPago.setFP_ID = lista.ListItems(i).SubItems(9)
                If cmbPedidos.Text = "" Then
                    oDocPago.setPEDIDO_ID = 0
                Else
                    oDocPago.setPEDIDO_ID = cmbPedidos.BoundText
                End If
'                oDocPago.setPEDIDO_ID = oPedido.Pedido_en_curso(lista.ListItems(i).SubItems(8), Date)
                oDocPago.setTOTAL = "0.00"
                oDocPago.setDESCUENTO = "0.00"
                If Tipo_documento = 2 Then
                    oDocPago.setIVA = IVA
                Else
                    oDocPago.setIVA = 0
                End If
                oDocPago.setPAGADO = 0
                oDocPago.setANULADO = 0
                oDocPago.setFACTURA_CONCEPTOS = 0
                ' Insertamos el documento de pago
                num_doc = oDocPago.InsertarDocPago
                If num_doc = 0 Then
                    MsgBox "Error al generar las facturas, contacte con mantenimiento.", vbCritical, App.Title
                    Exit Sub
                End If
                total_doc = total_doc + 1
                documentos_pago(total_doc) = num_doc
            End If
            ' Insertar el Documento de Pago de Muestras
            odoc_muestra.setDOC_ID = num_doc
            odoc_muestra.setMUESTRA_ID = lista.ListItems(i).SubItems(7)
            odoc_muestra.setTIPO_ANALISIS = lista.ListItems(i).SubItems(2)
            odoc_muestra.setFECHA = Format(lista.ListItems(i).SubItems(4), "yyyy-mm-dd")
            odoc_muestra.setREFERENCIA_CLIENTE = lista.ListItems(i).SubItems(3)
            odoc_muestra.setPRECIO = Replace(Format(lista.ListItems(i).SubItems(5), "0.00"), ",", ".")
            odoc_muestra.setORDEN = ORDEN
            ORDEN = odoc_muestra.Insertar_doc_pago_muestra(lista.ListItems(i).SubItems(10))
            If ORDEN = -1 Then
                MsgBox "Error al generar las facturas (2), contacte con mantenimiento.", vbCritical, App.Title
                Exit Sub
            Else
                ORDEN = ORDEN + 1
            End If
            ' Modificar el documento de pago de la muestra
            If omuestra.Informar_Documento_Pago(lista.ListItems(i).SubItems(7), Tipo_documento) = False Then
                MsgBox "Error al informar el documento de pago, contacte con mantenimiento.", vbCritical, App.Title
                Exit Sub
            End If
            cliente_ant = lista.ListItems(i).SubItems(8)
        End If
    Next
    Set omuestra = Nothing
    Set oDocPago = Nothing
    Set odoc_muestra = Nothing
    Dim stipo As String
    If Tipo_documento = 1 Then
        stipo = "Albaran"
    Else
        stipo = "Factura"
    End If
    If total_doc = 1 Then
        MsgBox "Se ha registrado 1 " & stipo & ".", vbOKOnly + vbInformation, App.Title
    Else
        MsgBox "Se han registrado " & total_doc & " " & stipo & "s.", vbOKOnly + vbInformation, App.Title
    End If
    ' LlamarMas Datos de la factura
    numero_documentos_pago = total_doc
    frmMasDatosFactura.Show 1
End Sub

Public Sub cabecera_grid()
    With lista.ColumnHeaders
        .Add , , "Código", 1200, lvwColumnLeft
        .Add , , "Cliente", 3000, lvwColumnLeft
        .Add , , "Analisis", 3000, lvwColumnLeft
        .Add , , "Ref.Cliente", 3000, lvwColumnLeft
        .Add , , "Fecha", 1300, lvwColumnCenter
        .Add , , "Precio", 1300, lvwColumnCenter
        .Add , , "General", 800, lvwColumnCenter
        .Add , , "ID_MUESTRA", 1, lvwColumnCenter
        .Add , , "CLIENTE_ID", 1, lvwColumnCenter
        .Add , , "FP_ID", 1, lvwColumnCenter
        .Add , , "FACTURA_DETERMINACIONES", 1, lvwColumnCenter
        .Add , , "BANO_ID", 1, lvwColumnCenter
        .Add , , "ANALISIS_MODIFICADO", 1, lvwColumnCenter
        .Add , , "TIPO_ANALISIS_ID", 1, lvwColumnCenter
    End With
End Sub

Public Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        cmdlog_Click
    End If
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub

Public Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.SelectedItem.Index).SubItems(7)
        frmVerMuestra.Show 1
        buscar CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(7)), lista.SelectedItem.Index
        gmuestra = 0
    End If
End Sub
Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub pedidos(cliente As Long)
    Dim oPedido As New clsClientes_pedidos
    Dim anterior As Integer
    If cmbPedidos.Text <> "" Then
        anterior = cmbPedidos.BoundText
    End If
    If cliente = 0 Then
        Set cmbPedidos.RowSource = oPedido.Listado_completo
    Else
        Set cmbPedidos.RowSource = oPedido.Listado_por_Cliente_en_Vigor(cliente, DTPicker_desde, DTPicker_hasta)
    End If
    cmbPedidos.ListField = "CODIGO_LARGO"
    cmbPedidos.DataField = "id_pedido"
    cmbPedidos.BoundColumn = "id_pedido"
    cmbPedidos.BoundText = anterior
End Sub
Private Sub tool()
   On Error GoTo tool_Error

   With m_cTT
    ' Creamos el toolTip pasandole el nombre del Formulario
    Call .Create(Me)
    'Establecemos el Ancho del ToolTip
    .MaxTipWidth = 600
    ' establece los márgenes
    .Margin(ttMarginBottom) = 7
    .Margin(ttMarginTop) = 7
    .Margin(ttMarginLeft) = 5
    .Margin(ttMarginRight) = 5
    ' Establecemos el tiempo que se muestra ( 7 segundos )
    .DelayTime(ttDelayShow) = 10000
    ' Agregamos un ToolTip al FileListBox
    'Para agregar mas controles solo hay que añadir uno por uno
    'Nota: solo es valido usar controles que posean HWND
    .AddTool lista
   End With

   On Error GoTo 0
   Exit Sub

tool_Error:

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tool of Formulario frmMuestraPendientesFacturacion"
End Sub

Private Sub lista_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
'    If Button And vbRightButton Then PopupMenu frmMenu.menuopciones
    If Button And vbRightButton Then cmdprevia_Click
End Sub

Private Sub colorear(fila As Integer, color As Long)
    Dim i As Integer
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
End Sub


Private Sub cargar_combo_clientes()
    Dim consulta As String
    consulta = "SELECT DISTINCT C.ID_CLIENTE,C.NOMBRE " & _
               "  FROM CLIENTES AS C, MUESTRAS AS M " & _
               " WHERE C.ID_CLIENTE = M.CLIENTE_ID " & _
               "   AND M.DOCUMENTO_PAGO=0 AND ANULADA = 0"
    With cmbclientes
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "CLIENTES"
            .setDESCRIPCION = "Clientes"
            .setPK = "ID_CLIENTE"
            .setCAMPO = "NOMBRE"
            .setQUERY = consulta
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmClientes
    End With

End Sub
