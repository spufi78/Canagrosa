VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmErrores 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas"
   ClientHeight    =   9570
   ClientLeft      =   2430
   ClientTop       =   1785
   ClientWidth     =   11490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmErrores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   11490
   WindowState     =   1  'Minimized
   Begin VB.CommandButton cmdCorregida 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Corregida"
      Height          =   855
      Left            =   7875
      Picture         =   "frmErrores.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   0
      Width           =   1125
   End
   Begin VB.CommandButton cmdScripts 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scripts"
      Height          =   855
      Left            =   9090
      Picture         =   "frmErrores.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   0
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
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
      Height          =   960
      Left            =   45
      TabIndex        =   20
      Top             =   8550
      Width           =   8070
      Begin VB.CommandButton cmdrestore 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   765
         Left            =   7110
         Picture         =   "frmErrores.frx":15D6
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   135
         Width           =   810
      End
      Begin VB.TextBox txtbusqueda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3735
         TabIndex        =   26
         Top             =   540
         Width           =   3165
      End
      Begin VB.CheckBox chkmis 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar solo mis consultas"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   270
         Value           =   1  'Checked
         Width           =   2670
      End
      Begin VB.CheckBox chkcorregidos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar consultas resueltas"
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   585
         Width           =   2805
      End
      Begin MSDataListLib.DataCombo cmbtipofiltro 
         Height          =   315
         Left            =   3735
         TabIndex        =   23
         Top             =   180
         Width           =   3150
         _ExtentX        =   5556
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
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Asunto"
         Height          =   195
         Index           =   2
         Left            =   3150
         TabIndex        =   25
         Top             =   630
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   1
         Left            =   3150
         TabIndex        =   24
         Top             =   270
         Width           =   315
      End
   End
   Begin VB.CheckBox chkcorregido 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Corregida"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      TabIndex        =   12
      Top             =   3870
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la consulta/Error"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   45
      TabIndex        =   7
      Top             =   900
      Width           =   11400
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Insertar"
         Height          =   945
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar Campos"
         Height          =   945
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   405
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox txtTexto 
         Height          =   1200
         Left            =   135
         TabIndex        =   1
         Top             =   1665
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   2117
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmErrores.frx":1EA0
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
      Begin RichTextLib.RichTextBox txtError 
         Height          =   345
         Left            =   135
         TabIndex        =   0
         Top             =   990
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   609
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmErrores.frx":1F22
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
      Begin MSDataListLib.DataCombo cmbtipo 
         Height          =   315
         Left            =   6480
         TabIndex        =   13
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo cmbPrioridad 
         Height          =   315
         Left            =   3060
         TabIndex        =   14
         Top             =   360
         Width           =   2880
         _ExtentX        =   5080
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
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   810
         TabIndex        =   15
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
         CalendarTitleBackColor=   14737632
         Format          =   52494337
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha_correccion 
         Height          =   330
         Left            =   1350
         TabIndex        =   19
         Top             =   2925
         Width           =   1335
         _ExtentX        =   2355
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
         CalendarTitleBackColor=   14737632
         Format          =   52494337
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   6075
         TabIndex        =   18
         Top             =   405
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Prioridad"
         Height          =   195
         Index           =   13
         Left            =   2340
         TabIndex        =   17
         Top             =   405
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Alta"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   16
         Top             =   405
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Asunto"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   1440
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   9135
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8640
      Width           =   1125
   End
   Begin VB.CommandButton cmdMinimizar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Minimizar"
      Height          =   855
      Left            =   10305
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8640
      Width           =   1125
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4230
      Left            =   45
      TabIndex        =   4
      Top             =   4275
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   7461
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión de consultas y errores en Geslab"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   11
      Top             =   90
      Width           =   4305
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10890
      Picture         =   "frmErrores.frx":1FA4
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique todos los datos necesarios con la mayor claridad posible"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   405
      Width           =   4530
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   11505
   End
End
Attribute VB_Name = "frmErrores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTipoFiltro_Change()
    cargar_lista
End Sub

Private Sub cmdImprimir_Click()
    If lista.ListItems.Count > 0 Then
        Dim consulta As String
        consulta = "{decodificadora.CODIGO}=" & DECODIFICADORA.ERRORES_PRIORIDADES & " AND {decodificadora_2.CODIGO}=" & DECODIFICADORA.ERRORES_TIPOS
        Dim tipo As String
        Dim tarea As String
        If cmbTipoFiltro.Text <> "" Then
            consulta = consulta & " AND {errores.TIPO_ID} = " & cmbTipoFiltro.BoundText
        End If
        consulta = consulta & " AND {errores.ERROR} like ""*" & txtbusqueda & "*"""
        If chkmis.value = Checked Then
            consulta = consulta & " AND {errores.EMPLEADO_ID} = " & usuario.getID_EMPLEADO
        End If
        If chkcorregidos.value = Unchecked Then
            consulta = consulta & " AND {errores.CORREGIDO} = 0"
        End If
        With frmReport
            .iniciar
            .informe = "\General\rpterrores_Listado"
            .criterio = consulta
            .imprimir = False
            .generar
            .Visible = True
        End With
    Else
        MsgBox "No existen datos en la lista.", vbExclamation, App.Title
    End If

End Sub

Private Sub cmdLimpiar_Click()
    borrar_campos
End Sub

Private Sub cmdMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub cmdrestore_Click()
    cmbTipoFiltro.Text = ""
    txtbusqueda = ""
    chkmis.value = Checked
    chkcorregidos.value = Unchecked
    cargar_lista
End Sub

Private Sub chkcorregidos_Click()
    cargar_lista
End Sub

Private Sub chkmis_Click()
    cargar_lista
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo fallo
    If validar = False Then
        Exit Sub
    End If
    If MsgBox("Va a dar de alta la consulta. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oError As New clsErrores
        With oError
            .setPRIORIDAD_ID = cmbPrioridad.BoundText
            .setTIPO_ID = cmbTipo.BoundText
            .setERROR = txtError.Text
            .setTEXTO = txttexto.Text
            .setEMPLEADO_ID = usuario.getID_EMPLEADO
            .setFECHA_ALTA = Format(Date, "yyyy-mm-dd")
            .setHORA_ALTA = Format(Time, "hh:mm:ss")
            .setCORREGIDO = 0
            .setFECHA_CORRECCION = "9999-12-31"
            .Insertar
        End With
        Me.MousePointer = 11
        ' Enviar a la web de proyectos
'        enviar_web (cod)
        enviar_mail
        cargar_lista
        Me.MousePointer = 0
        MsgBox "Incidencia generada correctamente.", vbInformation, App.Title
    End If
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al generar la consulta.", vbCritical, Err.Description
End Sub

Private Sub cmdCorregida_Click()
    If lista.ListItems.Count > 0 Then
        Dim oError As New clsErrores
        oError.Corregir lista.ListItems(lista.selectedItem.Index)
        Set oError = Nothing
        borrar_campos
        cargar_lista
    End If
End Sub

Private Sub cmdScripts_Click()
    frmScripts.Show 1
End Sub


Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 50
    Me.top = 50
    cargar_botones Me
    cargar_combos
    cabecera
    cargar_lista
    If UCase(usuario.getUSUARIO) = "JULIO" Then
        cmdCorregida.Visible = True
        cmdScripts.Visible = True
    Else
        cmdCorregida.Visible = False
        cmdScripts.Visible = False
    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 550, lvwColumnLeft
        .Add , , "Fecha", 1250, lvwColumnCenter
        .Add , , "Prioridad", 1250, lvwColumnCenter
        .Add , , "Tipo", 1250, lvwColumnCenter
        .Add , , "Error", 5300, lvwColumnLeft
        .Add , , "Descripcion", 1, lvwColumnLeft
        .Add , , "Usuario", 1500, lvwColumnCenter
        .Add , , "Corregida", 1, lvwColumnCenter
        .Add , , "Fecha_Correccion", 1, lvwColumnCenter
        .Add , , "Prioridad_id", 1, lvwColumnCenter
        .Add , , "Tipo_id", 1, lvwColumnCenter
    End With
End Sub
Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim corregidos As String
    Dim mios As String
    If chkcorregidos.value = Checked Then
        corregidos = ""
    Else
        corregidos = " and e.corregido = 0 "
    End If
    If chkmis.value = Checked Then
        mios = " and e.empleado_id = " & usuario.getID_EMPLEADO
    Else
        mios = ""
    End If
    Dim oErrores As New clsErrores
    Set rs = oErrores.Listado(chkmis.value, chkcorregidos.value, cmbTipoFiltro.BoundText, txtbusqueda.Text)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs.Fields(0), "0000"))
            .SubItems(1) = Format(rs.Fields(1), "dd-mm-yyyy")
            .SubItems(2) = rs.Fields(2)
            .SubItems(3) = rs.Fields(3)
            .SubItems(4) = rs.Fields(4)
            .SubItems(5) = rs.Fields(5)
            .SubItems(6) = UCase(rs.Fields(6))
            .SubItems(7) = rs.Fields(7)
            .SubItems(8) = rs.Fields(8)
            .SubItems(9) = rs.Fields(9)
            .SubItems(10) = rs.Fields(10)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    borrar_campos
    Set rs = Nothing
End Sub
Public Function enviar_web(incidencia As Long) As Boolean
    On Error GoTo falloConexion
    Dim consulta As String
    Dim cod As Integer
    Dim rs As New ADODB.Recordset
    Dim cw As ADODB.Connection
    Set cw = New ADODB.Connection
    cw.ConnectionString = "DRIVER=" & ReadINI(App.Path + "\config.ini", "SERVER", "DRIVER") & ";" _
                         & "SERVER=ixitec.net;" _
                         & "DATABASE=ixiteca_dprj1;" _
                         & "UID=ixiteca_ixiteca;" _
                         & "PWD=bde40ea;" _
                         & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    cw.Open
    consulta = "select max(task_id) from tasks"
    rs.ActiveConnection = cw
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
    rs.Open consulta
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        cod = 1
    Else
        cod = rs.Fields(0) + 1
    End If
    Set rs = Nothing
    consulta = "Insert into tasks " & _
               " values(" & cod & ",'" & CStr(incidencia) & ". " & txtError.Text & "'," & cod & ",0,1,2,'" & _
               Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss") & "',3,1,0,'" & Format(Date + 3, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss") & "',0,0,0,'" & _
               txttexto.Text & " (Reportado por " & usuario.getUSUARIO & ")',0,'',2,0,0,0,0,1,'','','',0)"
    cw.Execute consulta
    consulta = "Insert into user_tasks " & _
               " values(2,0," & cod & ",100,0)"
    cw.Execute consulta
    cw.Close
    enviar_web = True
    Exit Function
falloConexion:
    enviar_web = False
    MsgBox "Error al enviar a la web. " & Err.Description, vbCritical, App.Title
End Function
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        fecha = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        cmbPrioridad.BoundText = lista.ListItems(lista.selectedItem.Index).SubItems(9)
        cmbTipo.BoundText = lista.ListItems(lista.selectedItem.Index).SubItems(10)
        txtError = lista.ListItems(lista.selectedItem.Index).SubItems(4)
        txttexto = lista.ListItems(lista.selectedItem.Index).SubItems(5)
        chkcorregido.value = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        If lista.ListItems(lista.selectedItem.Index).SubItems(8) <> "9999-12-31" Then
            fecha_correccion.Visible = False
        Else
            fecha_correccion.Visible = True
            fecha_correccion = lista.ListItems(lista.selectedItem.Index).SubItems(8)
        End If
    End If
End Sub

Private Sub lista_KeyDown(KeyCode As Integer, Shift As Integer)
    lista_Click
End Sub

Public Sub enviar_mail()
   On Error Resume Next
   ret = Enviar_Mail_CDO(BUZON_CORREO_LOG, txtError.Text, txttexto.Text & "(Reportado por : " & usuario.getUSUARIO & ")", vbNullString)

End Sub

Private Function validar() As Boolean
    validar = True
    If cmbPrioridad.BoundText = "" Then
        MsgBox "La prioridad debe estar informada.", vbCritical, "Error"
        cmbPrioridad.SetFocus
        validar = False
        Exit Function
    End If
    If cmbTipo.BoundText = "" Then
        MsgBox "El tipo debe estar informado.", vbCritical, "Error"
        cmbTipo.SetFocus
        validar = False
        Exit Function
    End If
    If txtError.Text = "" Then
        MsgBox "El error debe estar informado.", vbCritical, "Error"
        txtError.SetFocus
        validar = False
        Exit Function
    End If
End Function

Private Sub borrar_campos()
    fecha = Date
    cmbTipo.Text = ""
    cmbPrioridad.Text = ""
    txttexto = ""
    txtError = ""
End Sub

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
   On Error GoTo cargar_combos_Error

    oDeco.cargar_combo cmbPrioridad, DECODIFICADORA.ERRORES_PRIORIDADES
    oDeco.cargar_combo cmbTipo, DECODIFICADORA.ERRORES_TIPOS
    oDeco.cargar_combo cmbTipoFiltro, DECODIFICADORA.ERRORES_TIPOS

   On Error GoTo 0
   Exit Sub

cargar_combos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_combos of Formulario frmErrores"
End Sub

Private Sub txtbusqueda_Change()
    cargar_lista
End Sub
