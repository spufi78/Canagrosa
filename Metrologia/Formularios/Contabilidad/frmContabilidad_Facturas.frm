VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmContabilidad_Facturas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas pendientes de Contabilizar"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmContabilidad_Facturas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   13635
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Facturas pendientes de contabilizar en "
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
      Height          =   1050
      Left            =   7830
      TabIndex        =   13
      Top             =   450
      Width           =   4515
      Begin VB.CheckBox chkTipoContabilidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BM Conta"
         Height          =   330
         Index           =   1
         Left            =   2835
         TabIndex        =   15
         Top             =   405
         Width           =   1455
      End
      Begin VB.CheckBox chkTipoContabilidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contabilidad tradicional"
         Height          =   330
         Index           =   0
         Left            =   270
         TabIndex        =   14
         Top             =   405
         Value           =   1  'Checked
         Width           =   2130
      End
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Buscar"
      Default         =   -1  'True
      Height          =   960
      Left            =   12465
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   540
      Width           =   1095
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   885
      Index           =   0
      Left            =   1290
      Picture         =   "frmContabilidad_Facturas.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8280
      Width           =   1155
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   885
      Index           =   1
      Left            =   90
      Picture         =   "frmContabilidad_Facturas.frx":1F14
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro de Selección de Facturas"
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
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   405
      Width           =   7695
      Begin VB.CheckBox chkFecha 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   315
         Left            =   210
         TabIndex        =   11
         Top             =   465
         Width           =   285
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1650
         TabIndex        =   7
         Top             =   450
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3825
         TabIndex        =   8
         Top             =   450
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   3210
         TabIndex        =   10
         Top             =   525
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Desde"
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   9
         Top             =   525
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6690
      Left            =   0
      TabIndex        =   0
      Top             =   1530
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   11800
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
      BackColor       =   14609914
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
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   11190
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturas pendientes de Contabilizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   13545
   End
End
Attribute VB_Name = "frmContabilidad_Facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbCliente_change()
    cargar_lista
End Sub

Private Sub cmbObra_change()
    cargar_lista
End Sub

Private Sub cmbTipoFacturacion_Change()
    cargar_lista
End Sub

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        fdesde.Enabled = True
        fhasta.Enabled = True
    Else
        fdesde.Enabled = False
        fhasta.Enabled = False
    End If
End Sub

Private Sub chkTipoContabilidad_Click(Index As Integer)
    Select Case Index
        Case 0
            If chkTipoContabilidad(Index).Value = Unchecked And chkTipoContabilidad(1).Value = Unchecked Then
                MsgBox "Debe seleccionar al menos un tipo de contabilidad", vbExclamation, App.Title
                chkTipoContabilidad(Index) = Checked
            End If
        Case 1
            If chkTipoContabilidad(Index).Value = Unchecked And chkTipoContabilidad(0).Value = Unchecked Then
                MsgBox "Debe seleccionar al menos un tipo de contabilidad", vbExclamation, App.Title
                chkTipoContabilidad(Index) = Checked
            End If
        End Select
        cargar_lista
End Sub

Private Function contabilizarTradicional() As Boolean
    Dim oConta As New clsContabilidad
    Dim oDOC As New clsDocumentos
    Dim Ret As Boolean
    Dim i As Integer
    contabilizarTradicional = False
    On Error GoTo errorContabilidad
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                Ret = oConta.Actualiza_Factura_Particular(lista.ListItems(i).SubItems(6))
            Else
                Ret = oConta.Actualiza_Factura(lista.ListItems(i).SubItems(6))
            End If
            If Ret = True Then
                oDOC.Contabilizar lista.ListItems(i).SubItems(6)
            Else
                Me.MousePointer = 0
                MsgBox "Error al generar la contabilidad tradicional de la factura : " & lista.ListItems(i).SubItems(1), vbExclamation, App.Title
                Exit Function
            End If
        End If
    Next
    contabilizarTradicional = True
    Exit Function
errorContabilidad:
    contabilizarTradicional = False
End Function

Private Function contabilizarBMConta() As Boolean
    Dim oConta As New clsContabilidad_BM
    Dim oDOC As New clsDocumentos
    Dim Ret As Boolean
    Dim i As Integer
    contabilizarBMConta = False
    On Error GoTo errorContabilidad
    
    If oConta.Conectar Then

        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                    Ret = oConta.Actualiza_Factura_Particular(lista.ListItems(i).SubItems(6))
                Else
                    Ret = oConta.Actualiza_Factura(lista.ListItems(i).SubItems(6))
                End If
                If Ret = True Then
                    oDOC.Contabilizar_BM lista.ListItems(i).SubItems(6)
                Else
                    Me.MousePointer = 0
                    MsgBox "Error al generar la contabilidad en BM Conta de la factura : " & lista.ListItems(i).SubItems(1), vbExclamation, App.Title
                    Exit Function
                End If
            End If
        Next
        
        If oConta.DesConectar Then
            contabilizarBMConta = True
        End If
        
    End If
    Exit Function
errorContabilidad:
    oConta.DesConectar True
    contabilizarBMConta = False
    
End Function

Private Sub cmdAceptar_Click()
    ' Validar que hay algo marcado
    Dim i As Integer
    Dim algo As Boolean
    'validad que hay marcado un tipo de contabilidad.
    Dim Mensaje As String
    Dim opcionesMarcadas As Integer
    Dim contTadicional As Boolean
    Dim contBMConta As Boolean
    Dim Ret As Boolean
    contTadicional = False
    contBMConta = False
    
   On Error GoTo cmdAceptar_Click_Error

    If lista.ListItems.Count > 0 Then
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                algo = True
            End If
        Next
        If Not algo Then
            MsgBox "Marque las facturas que desea contabilizar.", vbExclamation, App.Title
            Exit Sub
        End If
    Else
        MsgBox "No existen facturas para contabilizar.", vbExclamation, App.Title
        Exit Sub
    End If

    If chkTipoContabilidad(0).Value = Checked Then
        contTadicional = True
        If chkTipoContabilidad(1).Value = Checked Then
            Mensaje = "Se van a contabilizar las facturas marcadas en las dos contabilidades. ¿Desea continuar?"
            contBMConta = True
        Else
            Mensaje = "Se van a contabilizar las facturas marcadas en la contabilidad tradicional. ¿Desea continuar?"
        End If
    Else
        If chkTipoContabilidad(1).Value = Checked Then
            contBMConta = True
            Mensaje = "Se van a contabilizar las facturas marcadas en BM Conta. ¿Desea continuar?"
        End If
    End If
    
    If MsgBox(Mensaje, vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    ' Recuperar los albaranes a facturar
    Me.MousePointer = 11
    
    If contTadicional Then
        If contabilizarTradicional Then
                Me.MousePointer = 0
               MsgBox "Se ha generado correctamente la contabilidad.", vbInformation, App.Title
        End If
    End If
    Me.MousePointer = 11
    If contBMConta Then
        If contabilizarBMConta Then
            MsgBox "Se ha generado correctamente la contabilidad en BM Conta.", vbInformation, App.Title
            Me.MousePointer = 0
        End If
    End If
        
    cargar_lista

   On Error GoTo 0
   Exit Sub

cmdAceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptar_Click of Formulario frmContabilidad_Facturas"
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdMarcar_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = Index
    Next
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 100
    Me.Top = 100
    fdesde = Date - 31
    fhasta = Date
    cabecera_lista
    cargar_lista
    'Se activa la opción cuando el BMConta esté instalado.
    Me.chkTipoContabilidad.Item(1).Enabled = BMContaInstalado
End Sub
Private Sub cargar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim ESTADO As String
    Dim fecha As String
    Dim contabilidad As String
    
    If chkFecha.Value = Checked Then
        fecha = " AND D.FECHA BETWEEN '" & Format(fdesde, "YYYY-MM-DD") & "' AND '" & Format(fhasta, "YYYY-MM-DD") & "'"
    End If
' Tipo de contabilidad seleccionado
    contabilidad = ""
    If chkTipoContabilidad(0).Value = Checked Then
        contabilidad = " AND D.CONTABILIZADO = 0 "
    End If
    If chkTipoContabilidad(1).Value = Checked Then
        contabilidad = contabilidad & " AND D.CONTABILIZADO_BM = '0000-00-00 00:00:00'"
    End If
        
    tipo = " AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.factura
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT D.FECHA,D.NUMERO,C.NOMBRE,O.NOMBRE,D.TOTAL,D.PORTES,D.ID_DOCUMENTO,TD.ID_TIPO_DOCUMENTO,FP.NOMBRE " & _
               "  FROM DOCUMENTOS D, DOCUMENTOS_TIPOS TD, OBRAS O, CLIENTES C, FORMA_PAGO FP " & _
               " WHERE D.OBRA_ID = O.ID_OBRA " & _
               "   AND O.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "   AND O.FORMA_PAGO_ID =  FP.ID_FORMA_PAGO " & _
               contabilidad & _
               "   AND D.TOTAL <> 0 " & _
               "   AND D.ANULADO = 0 " & _
               tipo & ESTADO & fecha & _
               " ORDER BY D.NUMERO ASC"
    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim total As Currency
        total = 0
        While Not rs.EOF
            With lista.ListItems.Add(, , Format(rs.Fields(0), "yyyy-mm-dd"))
                .SubItems(1) = rs.Fields(1)
                If Not IsNull(rs.Fields(2)) Then
                 .SubItems(2) = rs.Fields(2) ' Numero de factura
                End If
                .SubItems(3) = rs.Fields(3)
                .SubItems(4) = Format(Replace(rs.Fields(4), ".", ","), "currency")
                .SubItems(5) = Format(Replace(rs.Fields(5), ".", ","), "currency")
                .SubItems(6) = rs.Fields(6)
                .SubItems(7) = rs.Fields(7)
                .SubItems(8) = rs.Fields(8)
            End With
            rs.MoveNext
        Wend
'        lista.SetFocus
    Else
        MsgBox "No existen facturas pendientes de contabilizar.", vbInformation, App.Title
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Documentos : " & Err.Description, vbCritical, Err.Description
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

Private Sub cabecera_lista()
    ' Pendientes
    With lista.ColumnHeaders
        .Add , , "Fecha", 1300, lvwColumnLeft
        .Add , , "Numero", 1000, lvwColumnCenter
        .Add , , "Cliente", 3200, lvwColumnLeft
        .Add , , "Obra", 3200, lvwColumnLeft
        .Add , , "Base", 1200, lvwColumnRight
        .Add , , "Portes", 1200, lvwColumnRight
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "TIPO_ID", 1, lvwColumnCenter
        .Add , , "Forma Pago", 2000, lvwColumnLeft
    End With
End Sub
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
            frmDocumento.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
            frmDocumento.Show 1
            actualizar_lista
    End If
End Sub

Private Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
   On Error GoTo actualizar_lista_Error

    consulta = "SELECT D.FECHA,D.NUMERO,C.NOMBRE,O.NOMBRE,D.TOTAL,D.PORTES,D.ID_DOCUMENTO,TD.ID_TIPO_DOCUMENTO, FP.NOMBRE " & _
               "  FROM DOCUMENTOS D, DOCUMENTOS_TIPOS TD, OBRAS O, CLIENTES C, FORMA_PAGO FP " & _
               " WHERE D.OBRA_ID = O.ID_OBRA " & _
               "   AND O.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "   AND O.FORMA_PAGO_ID =  FP.ID_FORMA_PAGO " & _
               "   AND D.ID_DOCUMENTO = " & lista.ListItems(lista.SelectedItem.Index).SubItems(6)
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        With lista.ListItems(lista.SelectedItem.Index)
            .Text = Format(rs.Fields(0), "yyyy-mm-dd")
            .SubItems(1) = rs.Fields(1)
            If Not IsNull(rs.Fields(2)) Then
             .SubItems(2) = rs.Fields(2) ' Numero de factura
            End If
            .SubItems(3) = rs.Fields(3)
            .SubItems(4) = Format(Replace(rs.Fields(4), ".", ","), "currency")
            .SubItems(5) = Format(Replace(rs.Fields(5), ".", ","), "currency")
            .SubItems(6) = rs.Fields(6)
            .SubItems(7) = rs.Fields(7)
            .SubItems(8) = rs.Fields(8)
        End With
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

actualizar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_lista of Formulario frmContabilidad_Facturas"
End Sub
