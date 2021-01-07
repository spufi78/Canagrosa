VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDatosEspecificos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Específicos"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "frmDatosEspecificos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSubir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Subir datos"
      Height          =   825
      Left            =   8010
      Picture         =   "frmDatosEspecificos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   1005
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6315
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   7995
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6315
      Width           =   1050
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar dato"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6300
      Width           =   1635
   End
   Begin VB.TextBox txtvalor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1635
      Left            =   750
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4545
      Width           =   7170
   End
   Begin MSComctlLib.ListView datos 
      Height          =   3645
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   6429
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
      NumItems        =   0
   End
   Begin MSDataListLib.DataCombo cmbDatos 
      Height          =   315
      Left            =   750
      TabIndex        =   3
      Top             =   4140
      Width           =   7170
      _ExtentX        =   12647
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
      Caption         =   "Dato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   5
      Top             =   4185
      Width           =   420
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   5220
      Width           =   450
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Datos Específicos"
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
      Left            =   45
      TabIndex        =   1
      Top             =   60
      Width           =   9090
   End
End
Attribute VB_Name = "frmDatosEspecificos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_MUESTRA As Long
Public PK_BANO As Long
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If datos.ListItems.Count > 0 Then
        datos.ListItems.Remove (datos.SelectedItem.Index)
        If datos.ListItems.Count > 0 Then
            datos_Click
        End If
    End If
End Sub

Private Sub cmdok_Click()
    If MsgBox("¿Modificar los datos especificos?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim consulta As String
        consulta = "delete from Datos_valores where muestra_id = " & PK_MUESTRA
        execute_bd consulta
        Dim j As Integer
        Dim ovalbano As New clsDatos_valores
        For j = 1 To datos.ListItems.Count
             ovalbano.setMUESTRA_ID = PK_MUESTRA
             ovalbano.setBANO_ID = PK_BANO
             ovalbano.setTIPO_DATO_ID = datos.ListItems(j).SubItems(2)
             ovalbano.setVALOR = datos.ListItems(j).SubItems(1)
             ovalbano.setORDEN = j
             If ovalbano.Insertar = False Then
                MsgBox "Error al insertar los datos específicos.", vbCritical, App.Title
                Exit Sub
             End If
        Next
'        imprimir gmuestra, 10, False
        MsgBox "Se han generado los datos específicos.", vbInformation, App.Title
        Unload Me
    End If
End Sub

Private Sub cmdSubir_Click()
   On Error GoTo cmdSubir_Click_Error

    If cmbDatos.Text <> "" Then
        Dim i As Integer
        Dim sw As Boolean
        sw = False
        If datos.ListItems.Count > 0 Then
            For i = 1 To datos.ListItems.Count
                If UCase(datos.ListItems(i).Text) = UCase(cmbDatos.Text) Then
                    datos.ListItems(i).SubItems(1) = txtvalor
                    sw = True
                End If
            Next
        End If
        If sw = False Or datos.ListItems.Count = 0 Then
            With datos.ListItems.Add(, , cmbDatos.Text)
                .SubItems(1) = txtvalor
                .SubItems(2) = cmbDatos.BoundText
            End With
        End If
        
        If datos.ListItems.Count > datos.SelectedItem.Index Then
            Set datos.SelectedItem = datos.ListItems(datos.SelectedItem.Index + 1)
            datos_Click
        Else
            txtvalor = ""
            datos.SetFocus
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdSubir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSubir_Click of Formulario frmDatosEspecificos"
End Sub
Private Sub datos_Click()
    If datos.ListItems.Count > 0 Then
        cmbDatos.Text = datos.ListItems(datos.SelectedItem.Index).Text
        txtvalor = datos.ListItems(datos.SelectedItem.Index).SubItems(1)
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_lista
    cargar_combo cmbDatos, New clsTipos_dato
    If PK_MUESTRA <> 0 Then
        cargar_muestra
    End If
    If datos.ListItems.Count > 0 Then
        datos_Click
    End If
End Sub

Public Sub cargar_lista()
    ' Datos
    With datos.ColumnHeaders.Add(, , "Dato", 3900, lvwColumnLeft)
        .Tag = "Dato"
    End With
    With datos.ColumnHeaders.Add(, , "Valor", 3800, lvwColumnLeft)
        .Tag = "Valor"
    End With
    With datos.ColumnHeaders.Add(, , "ID", 500, lvwColumnLeft)
        .Tag = "ID"
    End With
End Sub

Public Sub cargar_muestra()
    Dim ovb As New clsDatos_valores
    Dim rs As ADODB.RecordSet
    Set rs = ovb.datos_muestra_completo(PK_MUESTRA)
    If rs.RecordCount <> 0 Then
        Do
            With datos.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(1)
                 .SubItems(2) = rs(2)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set ovb = Nothing
End Sub

Private Sub txtvalor_GotFocus()
    txtvalor.BackColor = &H80C0FF
End Sub
Private Sub txtvalor_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        cmdSubir_Click
'        ' Pasar al siguiente campo
'        If datos.ListItems.Count > datos.SelectedItem.Index Then
'            Set datos.SelectedItem = datos.ListItems(datos.SelectedItem.Index + 1)
'            datos_Click
'        Else
'            txtvalor = ""
'            datos.SetFocus
'        End If
'    End If
End Sub
Private Sub txtvalor_LostFocus()
    txtvalor.BackColor = vbWhite
End Sub
