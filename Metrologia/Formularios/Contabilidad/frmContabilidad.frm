VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmContabilidad 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de facturas para contabilidad"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   Icon            =   "frmContabilidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   10545
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   9360
      Picture         =   "frmContabilidad.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7650
      Width           =   1050
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   45
      TabIndex        =   3
      Top             =   7425
      Width           =   7860
      Begin VB.CommandButton cmdruta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Abrir ruta de ficheros generados"
         Height          =   870
         Left            =   4230
         Picture         =   "frmContabilidad.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   225
         Width           =   1725
      End
      Begin VB.CommandButton cmdno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar a NO contabilizada"
         Enabled         =   0   'False
         Height          =   870
         Left            =   2250
         Picture         =   "frmContabilidad.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   225
         Width           =   1950
      End
      Begin VB.CommandButton cmdgenera 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar Fichero "
         Enabled         =   0   'False
         Height          =   870
         Left            =   5985
         Picture         =   "frmContabilidad.frx":1D68
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   1770
      End
      Begin VB.CommandButton cmdDesmarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ninguna"
         Height          =   870
         Left            =   1170
         Picture         =   "frmContabilidad.frx":2632
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1050
      End
      Begin VB.CommandButton cmdMarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todas"
         Height          =   870
         Left            =   90
         Picture         =   "frmContabilidad.frx":293C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   45
      TabIndex        =   1
      Top             =   360
      Width           =   10440
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contabilizadas"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   5355
         TabIndex        =   13
         Top             =   630
         Width           =   1410
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin contabilizar"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   5355
         TabIndex        =   8
         Top             =   225
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   780
         Left            =   9180
         Picture         =   "frmContabilidad.frx":2D7E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1155
         TabIndex        =   9
         Top             =   405
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   20578305
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3555
         TabIndex        =   10
         Top             =   405
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   20578305
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbSubtipo 
         Height          =   315
         Left            =   7200
         TabIndex        =   17
         Top             =   495
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   8
         Left            =   7200
         TabIndex        =   18
         Top             =   270
         Width           =   315
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta el"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   12
         Top             =   450
         Width           =   585
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde el"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   11
         Top             =   480
         Width           =   645
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5985
      Left            =   45
      TabIndex        =   7
      Top             =   1395
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10557
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
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Listado de facturas para contabilidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   10500
   End
End
Attribute VB_Name = "frmContabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    Call buscar
End Sub

Private Sub buscar()
    Dim importe As Currency
    Dim base As Currency
    Dim IVA As Currency
    On Error GoTo fallo
    If cmbSubtipo.Text = "" Then
        MsgBox "Seleccione un tipo de factura.", vbExclamation, App.Title
        Exit Sub
    End If
    lista.ListItems.Clear
    Dim rs As ADODB.Recordset
    Dim oDOCUMENTO As New clsDocumentos
    Me.MousePointer = 11
    Set rs = oDOCUMENTO.Listado_para_Contabilidad(fdesde, fhasta, Option1(1).Value, cmbSubtipo.BoundText)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs.Fields(0))
                    .SubItems(1) = rs.Fields(1)
                    .SubItems(2) = rs.Fields(2)
                    importe = rs.Fields(3)
                    IVA = (importe * 16) / 100
                    .SubItems(3) = Format(importe, "currency")
                    .SubItems(4) = Format(IVA, "currency")
                    .SubItems(5) = Format(importe + IVA, "currency")
                    .SubItems(6) = rs.Fields(4)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    If Option1(0).Value = True Then
        cmdno.Enabled = False
        cmdgenera.Enabled = True
    Else
        cmdno.Enabled = True
        cmdgenera.Enabled = False
    End If
    
    Me.MousePointer = 0
    Set oDOCUMENTO = Nothing
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al buscar las facturas.", vbCritical, Err.Description
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdgenera_Click()
   On Error GoTo cmdgenera_Click_Error

    If lista.ListItems.Count > 0 Then
        If contar_marcados = 0 Then
            MsgBox "Marque las facturas que quiere exportar a contaplus.", vbInformation, App.Title
        Else
            Dim oContabilidad As New clsContabilidad
            Dim i As Integer
            On Error Resume Next
            Dim documento As String
            If Dir(ReadINI(App.Path + "\config.ini", "documentos", "contabilidad")) = "" Then
                MkDir ReadINI(App.Path + "\config.ini", "documentos", "contabilidad")
            End If
            documento = ReadINI(App.Path + "\config.ini", "documentos", "contabilidad") & "\" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & "-" & cmbSubtipo.Text & "-" & usuario.getUSUARIO & ".txt"
            On Error GoTo cmdgenera_Click_Error
            oContabilidad.documento = documento
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
'                    If oContabilidad.verificacion_previa(lista.ListItems(i).SubItems(9)) = False Then
'                        If MsgBox("Existen conceptos o tipos de muestra con la familia sin informar en la factura " & lista.ListItems(i).Text & ", ¿Esta seguro de contabilizar?", vbExclamation + vbYesNo, App.Title) = vbYes Then
'                            oContabilidad.genera_contabilidad_por_documento lista.ListItems(i).SubItems(9)
'                        End If
'                    Else
                        oContabilidad.genera_contabilidad_por_documento lista.ListItems(i).SubItems(6)
'                    End If
                End If
            Next
            MsgBox "Proceso terminado correctamente.", vbInformation, App.Title
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & documento, vbMaximizedFocus)
            cmdBuscar_Click
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdgenera_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdgenera_Click of Formulario frmContabilidad"
End Sub

Private Sub cmdmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdno_Click()
    If lista.ListItems.Count > 0 Then
        If contar_marcados = 0 Then
            MsgBox "Marque las facturas para las que quiere anular la contabilidad.", vbInformation, App.Title
        Else
            If MsgBox("¿Esta seguro de anular la contabilidad?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                Dim oDOCUMENTO As New clsDocumentos
                Dim i As Integer
                For i = 1 To lista.ListItems.Count
                    If lista.ListItems(i).Checked = True Then
                        oDOCUMENTO.no_contabilizar lista.ListItems(i).SubItems(6)
                    End If
                Next
                MsgBox "Proceso terminado correctamente.", vbInformation, App.Title
                cmdBuscar_Click
            End If
        End If
    End If

End Sub

Private Sub cmdruta_Click()
    On Error Resume Next
    If Dir(ReadINI(App.Path + "\config.ini", "documentos", "contabilidad")) = "" Then
        MkDir ReadINI(App.Path + "\config.ini", "documentos", "contabilidad")
    End If
    r = Shell("explorer.exe " & ReadINI(App.Path + "\config.ini", "documentos", "contabilidad"), vbNormalFocus)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Me.SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 50
    Me.Top = 50
    fdesde = Date
    fhasta = Date
    cargar_combo cmbSubtipo, New clsDOCUMENTOS_SUBTIPOS
    
    cabecera
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "NºDoc", 800, lvwColumnLeft)
        .Tag = "NºDoc"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 4500, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1300, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Base", 1200, lvwColumnRight)
        .Tag = "Base"
    End With
    With lista.ColumnHeaders.Add(, , "Cuota I.V.A.", 1200, lvwColumnRight)
        .Tag = "Cuota I.V.A."
    End With
    With lista.ColumnHeaders.Add(, , "Total", 1200, lvwColumnRight)
        .Tag = "Total"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnLeft)
        .Tag = "ID"
    End With
End Sub
Public Function contar_marcados() As Integer
    Dim i As Integer
    Dim cont As Integer
    cont = 0
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            cont = cont + 1
        End If
    Next
    contar_marcados = cont
End Function
Private Sub lista_DblClick()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oDOCUMENTO As New clsDocumentos
    oDOCUMENTO.Imprimir CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(6)), False
End Sub

