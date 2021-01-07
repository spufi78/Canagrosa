VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpleados_Control 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de empleado"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   Icon            =   "frmEmpleados_Control.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7380
      Width           =   1155
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7380
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7410
      Width           =   1155
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7380
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   90
      TabIndex        =   15
      Top             =   6570
      Width           =   7935
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   6750
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   23
         Top             =   270
         Width           =   1110
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   12
         Top             =   270
         Width           =   1110
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   2925
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   11
         Top             =   270
         Width           =   1110
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   855
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   10
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Falta"
         Height          =   195
         Index           =   7
         Left            =   6300
         TabIndex        =   24
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kilometraje"
         Height          =   195
         Index           =   6
         Left            =   4185
         TabIndex        =   18
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Horas Extras"
         Height          =   195
         Index           =   5
         Left            =   1935
         TabIndex        =   17
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ausencia"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del control"
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
      Height          =   2040
      Left            =   45
      TabIndex        =   14
      Top             =   405
      Width           =   7980
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   885
         Left            =   6885
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   630
         Width           =   975
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   930
         Index           =   1
         Left            =   1080
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1035
         Width           =   5550
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   3
         Top             =   675
         Width           =   2595
      End
      Begin VB.CheckBox chkjus 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Justificada"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5445
         TabIndex        =   2
         Top             =   270
         Width           =   1140
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   2700
         TabIndex        =   1
         Top             =   225
         Width           =   2700
         _ExtentX        =   4763
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
      Begin MSComCtl2.DTPicker Fecha 
         Height          =   345
         Left            =   765
         TabIndex        =   0
         Top             =   225
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
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
         Format          =   16777217
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "* Nº de horas, kilometros, etc..."
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   8
         Left            =   3780
         TabIndex        =   25
         Top             =   720
         Width           =   2190
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comentario"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   22
         Top             =   1395
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   21
         Top             =   315
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   20
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   4
         Left            =   2250
         TabIndex        =   19
         Top             =   315
         Width           =   315
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4035
      Left            =   45
      TabIndex        =   5
      Top             =   2475
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   7117
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Control de empleado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   60
      TabIndex        =   13
      Top             =   30
      Width           =   7935
   End
End
Attribute VB_Name = "frmEmpleados_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTipo_Change()
'    borrar_campos
    If cmbTipo.Text = "" Then
        Exit Sub
    End If
    If cmbTipo.BoundText = 1 Or cmbTipo.BoundText = 4 Then
        chkjus.Enabled = True
    Else
        chkjus.Enabled = False
        chkjus.Value = Unchecked
    End If
    If cmbTipo.BoundText = 4 Then
        txtdatos(0).Enabled = False
        txtdatos(0) = ""
    Else
        txtdatos(0).Enabled = True
    End If
End Sub

Private Sub cmdAnadir_Click()
    If gOperario > 0 Then
        insertar_control
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        pregunta = "Va a ELIMINAR un control de empleado. ¿Esta seguro?"
        If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oc As New clsEmpleados_Control
            If oc.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
                MsgBox "El control de empleado se ha eliminado correctamente.", vbInformation, App.Title
                listado_control
            End If
        End If
    End If
End Sub

Private Sub cmdLimpiar_Click()
    borrar_campos
    Fecha = Date
    cmbTipo.Text = ""
End Sub

Private Sub cmdModificar_Click()
    If gOperario > 0 Then
        modificar_control
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combo cmbTipo, New clsEmpleados_Control_Tipos
    titulo_ventana
    cabecera_lista
    Fecha = Date
    If gOperario > 0 Then
        listado_control
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmEmpleados_Control = Nothing
End Sub

Private Sub lista_Click()
    consulta_control
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub
Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 1 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtDatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = &HFFFFFF
End Sub
Public Sub borrar_campos()
    Dim i As Integer
    For i = 0 To 1
        txtdatos(i) = ""
    Next
End Sub
Public Sub insertar_control()
    If valida_datos = False Then
        Exit Sub
    End If
    pregunta = "Va a dar de alta un nuevo control de empleado. ¿Esta seguro?"
    If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Set oExpediente = mover_datos
        If oExpediente.Insertar > 0 Then
            MsgBox "El control de empleado se ha insertado correctamente.", vbInformation, App.Title
            listado_control
        End If
    End If
End Sub
Public Sub modificar_control()
    If valida_datos() = False Then
        Exit Sub
    End If
    pregunta = "Va a modificar los datos del control de empleado. ¿Esta seguro?"
    If MsgBox(pregunta, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Set oExpediente = mover_datos
        If oExpediente.Modificar(lista.ListItems(lista.SelectedItem.Index)) = True Then
            listado_control
'            With lista.ListItems(lista.SelectedItem.Index)
'                .SubItems(1) = Format(Fecha, "dd/mm/yyyy")
'                If op(0).Value = True Then
'                    .SubItems(2) = "AUSENCIA"
'                ElseIf op(1).Value = True Then
'                    .SubItems(2) = "HORAS EXTRAS"
'                ElseIf op(2).Value = True Then
'                    .SubItems(2) = "KILOMETRAJE"
'                Else
'                    .SubItems(3) = "FALTA"
'                End If
'                .SubItems(3) = txtdatos(0)
'                .SubItems(4) = txtdatos(1)
'            End With
'            calcular_totales
            MsgBox "Control de empleado modificado correctamente.", vbInformation, App.Title
        Else
            MsgBox "Error al modificar el control de empleado.", vbInformation, App.Title
        End If
        Set oExpediente = Nothing
    End If
End Sub
Public Function valida_datos() As Boolean
    valida_datos = True
    If cmbTipo.Text <> "" Then
        If cmbTipo.BoundText > 0 And cmbTipo.BoundText < 4 And txtdatos(0) = "" Then
            MsgBox "El campo valor no puede estar en blanco.", vbCritical, "Error"
            txtdatos(0).SetFocus
            valida_datos = False
            Exit Function
        End If
    Else
       MsgBox "Introduzca el tipo de control.", vbCritical, "Error"
       txtdatos(0).SetFocus
       valida_datos = False
       Exit Function
    End If
End Function
Public Sub consulta_control()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oexp As New clsEmpleados_Control
    oexp.cargar (lista.ListItems(lista.SelectedItem.Index))
    With oexp
'        op(.getTIPO - 1).Value = True
        Fecha = .getFECHA
        txtdatos(0) = .getVALOR
        txtdatos(1) = .getCOMENTARIO
        cmbTipo.BoundText = .getTIPO
        If .getJUSTIFICADA = 1 Then
            chkjus.Value = Checked
        Else
            chkjus.Value = Unchecked
        End If
    End With
    Set oexp = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar el control de empleado.", vbCritical, Err.Description
End Sub
Public Function mover_datos() As clsEmpleados_Control
    On Error GoTo fallo
    Dim oexp As New clsEmpleados_Control
    With oexp
        .setEMPLEADO_ID = gOperario
        .setTIPO = cmbTipo.BoundText
        If chkjus.Value = Checked Then
            .setJUSTIFICADA = 1
        Else
            .setJUSTIFICADA = 0
        End If
        If cmbTipo.BoundText = 4 Then
            .setVALOR = 1
        Else
            .setVALOR = txtdatos(0)
        End If
        .setCOMENTARIO = txtdatos(1)
        .setFECHA = Format(Fecha.Value, "yyyy-mm-dd")
    End With
    Set mover_datos = oexp
    Set oexp = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del control de empleado.", vbCritical, Err.Description
End Function
Public Sub listado_control()
    Dim oexp As New clsEmpleados_Control
    Dim rs As ADODB.Recordset
    Set rs = oexp.Listado(gOperario)
    borrar_campos
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(0))
            .SubItems(1) = Format(rs(1), "dd/mm/yyyy")
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
            .SubItems(4) = rs(4)
           End With
           rs.MoveNext
        Loop Until rs.EOF
        calcular_totales
        consulta_control
    End If
End Sub
Public Sub cabecera_lista()
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1100, lvwColumnLeft)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo", 1400, lvwColumnCenter)
        .Tag = "Tipo"
    End With
    With lista.ColumnHeaders.Add(, , "Valor", 1200, lvwColumnCenter)
        .Tag = "Valor"
    End With
    With lista.ColumnHeaders.Add(, , "Comentario", 3800, lvwColumnLeft)
        .Tag = "Comentario"
    End With
End Sub
Public Sub titulo_ventana()
    Fecha = Date
    If gOperario > 0 Then
        Dim operario As New clsEmpleados
        operario.cargar (gOperario)
        lbltitulo.Caption = "Control de empleado : " & operario.getNOMBRE
        Me.Caption = lbltitulo.Caption
    End If
End Sub

Public Sub calcular_totales()
    Dim i As Integer
    Dim oc As New clsEmpleados_Control
    For i = 1 To 4
        txtdatos(i + 1) = CStr(oc.Totales(gOperario, i))
    Next
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
