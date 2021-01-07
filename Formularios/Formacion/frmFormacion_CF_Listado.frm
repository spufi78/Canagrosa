VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmFormacion_CF_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificados de Formación"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   12900
   Begin VB.TextBox txtPREFIX 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   4545
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Certificados de Formación"
      Top             =   90
      Width           =   3750
   End
   Begin VB.Frame Frame1 
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
      Height          =   1905
      Left            =   45
      TabIndex        =   5
      Top             =   630
      Width           =   12840
      Begin VB.TextBox txtCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1395
         TabIndex        =   7
         Top             =   225
         Width           =   1950
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   690
         Left            =   11925
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   900
         Width           =   825
      End
      Begin MSComCtl2.DTPicker fechaPrevistaI 
         Height          =   360
         Left            =   9225
         TabIndex        =   8
         Top             =   225
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   635
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
         Format          =   52887553
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaPrevistaF 
         Height          =   360
         Left            =   11160
         TabIndex        =   9
         Top             =   225
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   635
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
         Format          =   52887553
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbCertificador 
         Height          =   330
         Left            =   1395
         TabIndex        =   16
         Top             =   675
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbDepartamento 
         Height          =   330
         Left            =   1395
         TabIndex        =   17
         Top             =   1080
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbEmpleado 
         Height          =   330
         Left            =   1395
         TabIndex        =   18
         Top             =   1485
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   582
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Empleado:"
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   21
         Top             =   1530
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable:"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   20
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Certificador:"
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   19
         Top             =   765
         Width           =   840
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código:"
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   12
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Certificación:"
         Height          =   195
         Index           =   10
         Left            =   7695
         TabIndex        =   11
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
         Height          =   195
         Index           =   7
         Left            =   10845
         TabIndex        =   10
         Top             =   315
         Width           =   120
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8370
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Formulario"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8370
      Width           =   1230
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8370
      Width           =   1230
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Certificado"
      Height          =   870
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8370
      Width           =   1230
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5775
      Left            =   45
      TabIndex        =   4
      Top             =   2565
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   10186
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
   Begin VB.Image Image2 
      Height          =   480
      Left            =   12285
      Picture         =   "frmFormacion_CF_Listado.frx":0000
      Top             =   90
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   12945
   End
   Begin VB.Label lblSubTitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   14
      Top             =   375
      Width           =   510
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registros de Formación (RFI)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   45
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   12330
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frmFormacion_CF_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID. ", 1, lvwColumnLeft
        .Add , , "CERTIFICADO ", 1400, lvwColumnCenter
        .Add , , "Nombre", 3550, lvwColumnCenter
        .Add , , "Puesto", 1, lvwColumnCenter
        .Add , , "Departamento", 3900, lvwColumnCenter
        .Add , , "Materia", 2500, lvwColumnCenter
        .Add , , "Fecha Certificación", 1200, lvwColumnCenter
    End With
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdModificar_Click()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    frmFormacion_CF_Detalle.CUALIFICACION = 0
    frmFormacion_CF_Detalle.ID_DOC = 0
    frmFormacion_CF_Detalle.PK = CLng(lista.ListItems(lista.selectedItem.Index).Text)
    frmFormacion_CF_Detalle.Show 1
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmFormacion_CF_Listado"
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    fechaPrevistaI.value = Format(Date, "yyyy-01-01")
    fechaPrevistaF.value = Format(Date, "yyyy-12-31")
    cmdImprimir.Enabled = False
    cabecera
    cargar_lista
    cargar_combos
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
 
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a ELIMINAR el certificado Nº " & lista.ListItems(lista.selectedItem.Index) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oCertificado As New clsFormacion_certificados
        oCertificado.Eliminar CLng(lista.selectedItem.Text)
        Set oCertificado = Nothing
    End If
    cargar_lista
    lista.SetFocus
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Eliminar of Formulario frmFormacion_CF_Listado"
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oCertificados As New clsFormacion_certificados
    Dim cod As Long
       
    If txtCod.Text <> "" Then
       cod = CLng(Trim(txtCod.Text))
    Else
       cod = 0
    End If
    Set rs = oCertificados.ListadoFiltro(cod, Format(fechaPrevistaI.value, "yyyy-mm-dd"), Format(fechaPrevistaF.value, "yyyy-mm-dd"), CLng(cmbCertificador.getPK_SALIDA), CLng(cmbDepartamento.getPK_SALIDA), CLng(cmbEmpleado.getPK_SALIDA))
    lista.ListItems.Clear
     
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("ID_FORMACION_CERTIFICADO"))
                .Bold = True
                .ForeColor = vbBlue
                 .SubItems(1) = "CF " & rs("CODIGO") & "-" & rs("ANYO")
                 .ListSubItems(1).Bold = True
                 .ListSubItems(1).ForeColor = vbBlue
                 .SubItems(2) = rs("NOMBRE")
                 .SubItems(3) = " -- "
                 .SubItems(4) = DEPARTAMENTOS(rs("DEPARTAMENTOS"))
                 .SubItems(5) = rs("COD_DOC")
                 .SubItems(6) = Format(rs("FECHA_CERTIFICACION"), "yyyy-mm-dd")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
     
    Set oCertificados = Nothing
    Set rs = Nothing
End Sub

Private Function DEPARTAMENTOS(texto As String) As String
      'Obtención de la lista de departamentos (DESCRIPCIONES)
                                
        Dim strDepartamentos() As String
        Dim intCount As Integer
        Dim VALOR As Long
        Dim oDepart As New clsDecodificadora
        DEPARTAMENTOS = ""
        strDepartamentos = Split(Trim(texto), ";")
                
        For intCount = LBound(strDepartamentos) To UBound(strDepartamentos)

            If strDepartamentos(intCount) <> "" Then 'Para prevenir el caso de encontrar un ; al final de la línea de parámetros
                VALOR = CLng(Solo_Numeros(strDepartamentos(intCount)))
                oDepart.Carga_valor 50, VALOR
                'intcount: número de parámetros
                If intCount > LBound(strDepartamentos) Then
                    DEPARTAMENTOS = DEPARTAMENTOS & ", "
                End If
                DEPARTAMENTOS = DEPARTAMENTOS & oDepart.getDESCRIPCION
            End If
        Next intCount
End Function

Private Sub cargar_combos()
    llenar_combo cmbCertificador, New clsEmpleados, 0, frmEmpleados_Gestion, ""
    llenar_combo cmbDepartamento, New clsEmpleados, 0, frmEmpleados_Gestion, ""
    llenar_combo cmbEmpleado, New clsEmpleados, 0, frmEmpleados_Gestion, ""
End Sub

Private Sub lista_Click()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    cmdImprimir.Enabled = True
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lista_Click of Formulario frmFormacion_CF_Listado"
End Sub

Private Sub lista_DblClick()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    frmFormacion_CF_Detalle.CUALIFICACION = 0
    frmFormacion_CF_Detalle.ID_DOC = 0
    frmFormacion_CF_Detalle.PK = CLng(lista.ListItems(lista.selectedItem.Index).Text)
    frmFormacion_CF_Detalle.Show 1
    
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lista_DblClick of Formulario frmFormacion_CF_Listado"
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_SoloNumerico(txtCod, KeyAscii)
    txtCod.Refresh
    'txtCod = txtCod & Chr(KeyAscii)
    'cargar_lista
End Sub

 
