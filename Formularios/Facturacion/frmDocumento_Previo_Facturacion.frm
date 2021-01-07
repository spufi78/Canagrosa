VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmDocumento_Previo_Facturacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Previsualización de detalle de facturación de la muestra"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11175
   Icon            =   "frmDocumento_Previo_Facturacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la factura"
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
      Height          =   1275
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   11085
      Begin VB.CheckBox chkDeterminaciones 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "El cliente o la muestra se facturan por determinaciones"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   225
         TabIndex        =   5
         Top             =   810
         Width           =   4560
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   375
         Left            =   900
         TabIndex        =   3
         Top             =   360
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbTarifa 
         Height          =   375
         Left            =   6210
         TabIndex        =   7
         Top             =   765
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarifa"
         Height          =   195
         Index           =   1
         Left            =   5715
         TabIndex        =   6
         Top             =   810
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   420
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6210
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4755
      Left            =   60
      TabIndex        =   0
      Top             =   1395
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   8387
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
   Begin VB.Shape Shape1 
      Height          =   7125
      Left            =   0
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "frmDocumento_Previo_Facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_MUESTRA As Long
Public FACTURA_DETERMINACIONES As Integer

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbTarifa, New clsTarifas, 0, Me, ""
    cabecera
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    chkDeterminaciones.Value = FACTURA_DETERMINACIONES
    If PK_MUESTRA > 0 Then
        Dim oMuestra As New clsMuestra
        Dim oCliente As New clsCliente
        Dim oTA As New clsTipos_analisis
        Dim oBANO As New clsBanos
        Dim oCodigo As New clsTarifas_codigos
        Dim CODIGO As String
        oMuestra.CargaMuestra PK_MUESTRA
        oCliente.CargaCliente oMuestra.getCLIENTE_ID
        cmbclientes.MostrarElemento oMuestra.getCLIENTE_ID
        cmbTarifa.MostrarElemento oCliente.getTARIFA_ID
        If FACTURA_DETERMINACIONES Then
            CODIGO = ""
        Else
           If oMuestra.getBANO_ID <> 0 And oMuestra.getANALISIS_MODIFICADO <> 2 Then
              oBANO.cargar_bano CLng(oMuestra.getBANO_ID)
              oCodigo.Carga oBANO.getTARIFA_CODIGO_ID
           Else
              oTA.CARGAR oMuestra.getTIPO_ANALISIS_ID
              oCodigo.Carga oTA.getTARIFA_CODIGO_ID
           End If
           CODIGO = oCodigo.getCODIGO
        End If
        With lista.ListItems.Add(, , Format(oMuestra.getFECHA_RECEPCION, "dd-mm-yyyy"))
          .SubItems(1) = oMuestra.CodigoParticular(PK_MUESTRA)
          oTA.CARGAR oMuestra.getTIPO_ANALISIS_ID
          .SubItems(2) = oTA.getNOMBRE
          .SubItems(3) = oMuestra.getREFERENCIA_CLIENTE
          .SubItems(4) = CODIGO
          If FACTURA_DETERMINACIONES Then
            .SubItems(5) = moneda(oMuestra.ImporteMuestraPorDeterminaciones(PK_MUESTRA, oMuestra.getCLIENTE_ID))
          Else
            .SubItems(5) = moneda(oMuestra.getPRECIO)
          End If
        End With
        If FACTURA_DETERMINACIONES Then
            If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.SELLANTE Then
                Dim oSR As New clsSellantes_resultados
                Set rs = oSR.Listado_Resultados(PK_MUESTRA)
                If rs.RecordCount > 0 Then
                    Do
                        With lista.ListItems.Add(, , "")
                          .SubItems(1) = ""
                          .SubItems(2) = rs(1)
                          .SubItems(3) = ""
                          .SubItems(4) = ""
                          .SubItems(5) = moneda("0")
                        End With
                        rs.MoveNext
                    Loop Until rs.EOF
                End If
            Else
                Dim PRECIO As String
                Dim oTP As New clsTarifas_precios
                Dim oDeter As New clsDeterminaciones
                Set rs = oDeter.lista_determinaciones(PK_MUESTRA)
                If rs.RecordCount > 0 Then
                 Do
                    If oCodigo.Carga(rs("tarifa_codigo_id")) Then
                        CODIGO = oCodigo.getCODIGO
                    Else
                        CODIGO = ""
                    End If
                    With lista.ListItems.Add(, , "")
                      .SubItems(1) = ""
                      .SubItems(2) = rs("nombre")
                      .SubItems(3) = rs("proc_ref_eads")
                      .SubItems(4) = CODIGO
                      If oTP.Carga_por_determinacion(rs("id_tipo_determinacion"), oCliente.getTARIFA_ID) Then
                          PRECIO = oTP.getPRECIO
                      Else
                          PRECIO = "0"
                      End If
                      .SubItems(5) = moneda(PRECIO)
                    End With
                    rs.MoveNext
                 Loop Until rs.EOF
                End If
            End If
        End If
    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Fecha", 1000, lvwColumnLeft
        .Add , , "NºEnsayo", 1000, lvwColumnCenter
        .Add , , "Tipo de Análisis", 3200, lvwColumnLeft
        .Add , , "Referencia Cliente", 3200, lvwColumnLeft
        .Add , , "Código", 1200, lvwColumnCenter
        .Add , , "Precio", 1200, lvwColumnRight
    End With
End Sub
