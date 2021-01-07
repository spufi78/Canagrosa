VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmEmpleados_Categorias_Historia 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categorias"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12480
   Icon            =   "frmEmpleados_Categorias_Historia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   12480
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la Categoría"
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
      Height          =   1725
      Left            =   45
      TabIndex        =   5
      Top             =   5400
      Width           =   10770
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver Ficha"
         Height          =   840
         Left            =   6165
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   810
         Width           =   1095
      End
      Begin VB.CheckBox chkActual 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Actualidad"
         Height          =   195
         Left            =   3195
         TabIndex        =   15
         Top             =   1305
         Width           =   2670
      End
      Begin VB.CommandButton cmdmodificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   840
         Left            =   8415
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   810
         Width           =   1095
      End
      Begin VB.CommandButton cmdeliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   840
         Left            =   9540
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   810
         Width           =   1095
      End
      Begin VB.CommandButton cmdanadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   840
         Left            =   7290
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   810
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   360
         Left            =   1305
         TabIndex        =   10
         Top             =   855
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   60358657
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   360
         Left            =   4320
         TabIndex        =   12
         Top             =   855
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   60358657
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbPuestos 
         Height          =   330
         Left            =   1305
         TabIndex        =   14
         Top             =   360
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   582
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de Fin"
         Height          =   195
         Index           =   2
         Left            =   3195
         TabIndex        =   13
         Top             =   945
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   11
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Puesto"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   405
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Listado de Empleados"
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
      Height          =   4530
      Left            =   45
      TabIndex        =   3
      Top             =   855
      Width           =   12390
      Begin MSComctlLib.ListView lista 
         Height          =   4140
         Left            =   90
         TabIndex        =   4
         Top             =   225
         Width           =   12240
         _ExtentX        =   21590
         _ExtentY        =   7303
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
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   11250
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6255
      Width           =   1155
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Relación de departamentos y distinas categorías dentro de la empresa"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   4980
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11880
      Picture         =   "frmEmpleados_Categorias_Historia.frx":08CA
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos históricos de Categorías"
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
      TabIndex        =   0
      Top             =   120
      Width           =   3225
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   12465
   End
End
Attribute VB_Name = "frmEmpleados_Categorias_Historia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmdMostrar_Click()
    If cmbPuestos.getTEXTO <> "" Then
        Dim oDoc As New clsCa_documentos
        Dim oCa As New clsEmpleados_categorias
        oCa.Carga cmbPuestos.getPK_SALIDA
        oDoc.mostrar oCa.getDOCUMENTO_ID, True
        Set oDoc = Nothing
        Set oCa = Nothing
    End If
End Sub

Private Sub chkActual_Click()
    If chkActual.value = Checked Then
        fhasta.Enabled = False
    Else
        fhasta.Enabled = True
    End If
End Sub

Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If validar Then
        Dim oCat As New clsEmpleados_categorias_historia
            With oCat
                .setEMPLEADO_ID = PK
                .setCATEGORIA_ID = cmbPuestos.getPK_SALIDA
                .setFECHA_INICIO = Format(fdesde, "yyyy-mm-dd")
                If chkActual.value = Unchecked Then
                    .setFECHA_FIN = Format(fhasta, "yyyy-mm-dd")
                Else
                    .setFECHA_FIN = "0000-00-00"
                End If
                .setORDEN = lista.ListItems.Count + 1
                .insertar
            End With
        Set oCat = Nothing
        cargar_categorias
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdanadir_Click of Formulario frmEmpleados_Categorias_Historia"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDepartamentos_Click()
    Dim oform As New frmDecodificadora
    oform.CODIGO = decodificadora.EMPLEADOS_DEPARTAMENTOS
    oform.Show
    Set oform = Nothing
End Sub

Private Sub cmdeliminar_Click()
   On Error GoTo cmdEliminar_Click_Error
    If lista.ListItems.Count > 0 Then
        Dim oCat As New clsEmpleados_categorias_historia
        oCat.EliminarCategoria PK, CInt(lista.ListItems(lista.SelectedItem.Index).SubItems(5))
        Set oCat = Nothing
        cargar_categorias
    End If

   On Error GoTo 0
   Exit Sub

cmdEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdeliminar_Click of Formulario frmEmpleados_Categorias_Historia"

End Sub
Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_Error
    If validar Then
        Dim oCat As New clsEmpleados_categorias_historia
            With oCat
                .setCATEGORIA_ID = cmbPuestos.getPK_SALIDA
                .setFECHA_INICIO = Format(fdesde, "yyyy-mm-dd")
                If chkActual.value = Unchecked Then
                    .setFECHA_FIN = Format(fhasta, "yyyy-mm-dd")
                Else
                    .setFECHA_FIN = "0000-00-00"
                End If
                .Modificar PK, lista.ListItems(lista.SelectedItem.Index).SubItems(5)
            End With
        Set oCat = Nothing
        cargar_categorias
    End If


   On Error GoTo 0
   Exit Sub

cmdModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdmodificar_Click of Formulario frmEmpleados_Categorias_Historia"

End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargar_combos
    fdesde = Date
    fhasta = Date
    chkActual.value = Checked
    If PK > 0 Then
        cargar_categorias
    End If
End Sub

Private Sub cargar_categorias()
    Dim oCat As New clsEmpleados_categorias_historia
    Dim rs As ADODB.RecordSet
    Set rs = oCat.Listado(PK)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0)) ' puesto
                .SubItems(1) = rs(1) ' departamento
                .SubItems(2) = Format(rs(2), "dd-mm-yyyy") ' fdesde
                If IsNull(rs(3)) Then ' fhasta
                    .SubItems(3) = "Actualidad"
                Else
                    If rs(3) = "0000-00-00" Then
                        .SubItems(3) = "Actualidad"
                    Else
                        .SubItems(3) = Format(rs(3), "dd-mm-yyyy")
                    End If
                End If
                .SubItems(4) = rs(4) ' id_categoria
                .SubItems(5) = rs(5) ' orden
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oCat = Nothing
End Sub
Private Sub cargar_combos()
    llenar_combo cmbPuestos, New clsEmpleados_categorias, 0, Me, ""
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Puesto", 5500, lvwColumnLeft
        .Add , , "Departamento", 4000, lvwColumnLeft
        .Add , , "F.Inicio", 1300, lvwColumnCenter
        .Add , , "F.Fin", 1300, lvwColumnCenter
        .Add , , "ID_CATEGORIA", 1, lvwColumnLeft
        .Add , , "ORDEN", 1, lvwColumnLeft
    End With
End Sub

Private Function validar() As Boolean
    validar = True
    If cmbPuestos.getTEXTO = "" Then
        MsgBox "Debe indicar la categoria.", vbCritical, App.Title
        cmbPuestos.SetFocus
        validar = False
        Exit Function
    End If
End Function
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        cmbPuestos.MostrarElemento lista.ListItems(lista.SelectedItem.Index).SubItems(4)
        fdesde = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
        If lista.ListItems(lista.SelectedItem.Index).SubItems(3) = "Actualidad" Then
            chkActual.value = Checked
        Else
            chkActual.value = Unchecked
            fhasta = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
        End If
    End If
End Sub
