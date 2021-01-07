VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTDE_Analisis 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Específicos"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTDE_Analisis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5895
      Picture         =   "frmTDE_Analisis.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4500
      Width           =   645
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5220
      Picture         =   "frmTDE_Analisis.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4500
      Width           =   645
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5100
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   5475
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5100
      Width           =   1050
   End
   Begin MSComctlLib.ListView datos 
      Height          =   3645
      Left            =   45
      TabIndex        =   0
      Top             =   765
      Width           =   6060
      _ExtentX        =   10689
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
      Left            =   1035
      TabIndex        =   1
      Top             =   4590
      Width           =   3975
      _ExtentX        =   7011
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
   Begin VB.Image flecha 
      Height          =   480
      Index           =   1
      Left            =   6120
      Picture         =   "frmTDE_Analisis.frx":11A0
      Top             =   2745
      Width           =   480
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   0
      Left            =   6120
      Picture         =   "frmTDE_Analisis.frx":16E0
      Top             =   1935
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento de datos específicos"
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
      Left            =   90
      TabIndex        =   6
      Top             =   120
      Width           =   3750
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   6030
      Picture         =   "frmTDE_Analisis.frx":1C1C
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento de datos específicos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   420
      Width           =   2565
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo de Dato"
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   2
      Top             =   4635
      Width           =   930
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   6675
   End
End
Attribute VB_Name = "frmTDE_Analisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_ANALISIS As Long
Public PK_BANO As Long
Const campos = 2

Private Sub cmdAnadir_Click()
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
                .SubItems(1) = Format(cmbDatos.BoundText, "00")
            End With
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If datos.ListItems.Count > 0 Then
        datos.ListItems.Remove (datos.SelectedItem.Index)
    End If
End Sub

Private Sub cmdOk_Click()
   On Error GoTo cmdok_Click_Error

    If MsgBox("¿Desea actualizar los datos especificos?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oTDA As New clsTipos_datos_analisis
        oTDA.Eliminar PK_ANALISIS, PK_BANO
        Dim i As Integer
        For i = 1 To datos.ListItems.Count
            With oTDA
                .setTIPO_DATO_ID = datos.ListItems(i).SubItems(1)
                .setTIPO_ANALISIS_ID = PK_ANALISIS
                .setBANO_ID = PK_BANO
                .setORDEN = i
                .Insertar
            End With
        Next
        MsgBox "Se han actualizado los datos específicos correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmTDE_Analisis", vbCritical, App.Title
End Sub
Private Sub datos_Click()
    If datos.ListItems.Count > 0 Then
        cmbDatos.Text = datos.ListItems(datos.SelectedItem.Index).Text
        txtvalor = datos.ListItems(datos.SelectedItem.Index).SubItems(1)
    End If
End Sub


Private Sub flecha_Click(Index As Integer)
    Dim aux As String
    Dim i As Integer
    If datos.ListItems.Count > 0 Then
        If Index = 0 Then 'Subir
           If datos.SelectedItem.Index > 1 Then
              aux = datos.ListItems(datos.SelectedItem.Index - 1).Text
              datos.ListItems(datos.SelectedItem.Index - 1).Text = datos.ListItems(datos.SelectedItem.Index).Text
              datos.ListItems(datos.SelectedItem.Index).Text = aux
              For i = 1 To campos - 1
                  aux = datos.ListItems(datos.SelectedItem.Index - 1).SubItems(i)
                  datos.ListItems(datos.SelectedItem.Index - 1).SubItems(i) = datos.ListItems(datos.SelectedItem.Index).SubItems(i)
                  datos.ListItems(datos.SelectedItem.Index).SubItems(i) = aux
              Next
              Set datos.SelectedItem = datos.ListItems(datos.SelectedItem.Index - 1)
           End If
        Else ' Bajar
           If datos.SelectedItem.Index < datos.ListItems.Count Then
              aux = datos.ListItems(datos.SelectedItem.Index + 1).Text
              datos.ListItems(datos.SelectedItem.Index + 1).Text = datos.ListItems(datos.SelectedItem.Index).Text
              datos.ListItems(datos.SelectedItem.Index).Text = aux
              For i = 1 To campos - 1
                  aux = datos.ListItems(datos.SelectedItem.Index + 1).SubItems(i)
                  datos.ListItems(datos.SelectedItem.Index + 1).SubItems(i) = datos.ListItems(datos.SelectedItem.Index).SubItems(i)
                  datos.ListItems(datos.SelectedItem.Index).SubItems(i) = aux
              Next
              Set datos.SelectedItem = datos.ListItems(datos.SelectedItem.Index + 1)
           End If
        End If
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    Cargar_Combo cmbDatos, New clsTipos_dato
    cargar_datos
End Sub

Public Sub cabecera()
    With datos.ColumnHeaders
        .Add , , "Tipo de Dato", 5500, lvwColumnLeft
        .Add , , "ID", 500, lvwColumnLeft
    End With
End Sub

Public Sub cargar_datos()
    Dim oTDA As New clsTipos_datos_analisis
    Dim rs As ADODB.Recordset
    If PK_ANALISIS <> 0 Then
        Dim oTA As New clsTipos_analisis
        If oTA.cargar(PK_ANALISIS) = True Then
            lblsubtitulo = "Datos específicos de análisis : " & UCase(oTA.getNOMBRE)
        End If
        Set rs = oTDA.Listado_por_tipo_analisis(PK_ANALISIS)
    Else
        Dim oBANO As New clsBanos
        If oBANO.cargar_bano(PK_BANO) = True Then
            lblsubtitulo = "Datos específicos del baño : " & UCase(oBANO.getNOMBRE)
        End If
        Set rs = oTDA.Listado_por_bano(PK_BANO)
    End If
    If rs.RecordCount <> 0 Then
        Do
            With datos.ListItems.Add(, , rs(1))
                 .SubItems(1) = Format(rs(0), "00")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oTDA = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PK_ANALISIS = 0
    PK_BANO = 0
End Sub
