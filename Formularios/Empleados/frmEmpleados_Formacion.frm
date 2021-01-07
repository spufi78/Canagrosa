VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmpleados_Formacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formación de Empleados"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   Icon            =   "frmEmpleados_Formacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   915
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8955
      Width           =   1275
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   8310
      Left            =   45
      TabIndex        =   0
      Top             =   585
      Width           =   10455
      Begin VB.Frame frmAdjunto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Adjunto"
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
         Left            =   135
         TabIndex        =   12
         Top             =   7155
         Width           =   10230
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   870
            Left            =   9270
            Picture         =   "frmEmpleados_Formacion.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   180
            Width           =   870
         End
         Begin VB.TextBox datos 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   2520
            TabIndex        =   19
            Top             =   180
            Visible         =   0   'False
            Width           =   3540
         End
         Begin VB.CommandButton cmdAdjuntar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Adjuntar"
            Height          =   870
            Index           =   0
            Left            =   8370
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   180
            Width           =   870
         End
         Begin VB.TextBox datos 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataSource      =   "Adodc1"
            Height          =   330
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1170
            TabIndex        =   14
            Top             =   450
            Width           =   6120
         End
         Begin VB.CommandButton cmdEXplorar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Explorar"
            Height          =   870
            Index           =   0
            Left            =   7470
            Picture         =   "frmEmpleados_Formacion.frx":1194
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   180
            Width           =   870
         End
         Begin VB.Label lblCampos 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Adjunto"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   15
            Top             =   540
            Width           =   555
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos de la formación a adjuntar"
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
         Height          =   1185
         Left            =   135
         TabIndex        =   5
         Top             =   5895
         Width           =   10230
         Begin VB.CommandButton cmdmodificar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Modificar"
            Height          =   840
            Left            =   8370
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   225
            Width           =   870
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   840
            Left            =   9270
            Picture         =   "frmEmpleados_Formacion.frx":1A5E
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   225
            Width           =   870
         End
         Begin VB.CommandButton cmdDocumentacion 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            Height          =   840
            Left            =   7470
            Picture         =   "frmEmpleados_Formacion.frx":2328
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   225
            Width           =   870
         End
         Begin VB.TextBox txtdatos 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1170
            TabIndex        =   6
            Top             =   720
            Width           =   6135
         End
         Begin MSComCtl2.DTPicker fecha 
            Height          =   360
            Left            =   1170
            TabIndex        =   9
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   75431937
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Descripción"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   8
            Top             =   765
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F. Obtención"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView lista 
         Height          =   5310
         Left            =   75
         TabIndex        =   1
         Top             =   180
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   9366
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Doble-Click para ver el Adjunto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3555
         TabIndex        =   17
         Top             =   5580
         Width           =   3525
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   7110
      Top             =   9090
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3870
      Top             =   9135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Formacion.frx":2BF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del tipo de análisis"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   315
      Width           =   1830
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Formación del Empleado"
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
      TabIndex        =   2
      Top             =   45
      Width           =   2625
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmEmpleados_Formacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal Hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public PK As Long

Private Sub cmdAdjuntar_Click(Index As Integer)
   On Error GoTo cmdAdjuntar_Click_Error

    ' Validar
    If datos(4) = "" Then
        MsgBox "Escriba una ruta.", vbInformation, App.Title
        Exit Sub
    End If
    If Dir(datos(4)) = "" Then
        MsgBox "La ruta introducida no existe.", vbInformation, App.Title
        Exit Sub
    End If
'    adjuntar PK
    Me.MousePointer = 11
    Dim oEF As New clsEmpleados_formacion
    With oEF
        .setRUTA = datos(0)
        .ModificarAdjunto lista.ListItems(lista.selectedItem.Index).Text
    End With
    Set oEF = Nothing
    Dim oRR As New clsRRHH
    oRR.SubirFormacion CLng(lista.ListItems(lista.selectedItem.Index).Text), datos(4), datos(0)
    Set oRR = Nothing
    Me.MousePointer = 0
    MsgBox "El archivo se ha adjuntado correctamente.", vbOKOnly + vbInformation, App.Title
    cargar_lista
   On Error GoTo 0
   Exit Sub

cmdAdjuntar_Click_Error:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntar_Click of Formulario frmEmpleados_formacion"
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDocumentacion_Click()
   On Error GoTo cmdDocumentacion_Click_Error

    If txtdatos = "" Then
        MsgBox "Introduzca una descripción para el archivo, por ejemplo, Curriculum, etc...", vbExclamation, App.Title
        Exit Sub
    End If
'    On Error Resume Next
'    cd.DialogTitle = "Adjuntar archivo al empleado..."
'    cd.ShowOpen
'    If cd.FileName <> "" Then
        Dim oEF As New clsEmpleados_formacion
        With oEF
            .setEMPLEADO_ID = PK
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setDESCRIPCION = txtdatos
            .setRUTA = ""
'            .setRUTA = Replace(cd.FileName, "\", "/")
            .Insertar
        End With
        Set oEF = Nothing
        MsgBox "Registro insertado correctamente.", vbInformation, App.Title
        cargar_lista
'    End If

   On Error GoTo 0
   Exit Sub

cmdDocumentacion_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdDocumentacion_Click of Formulario frmEmpleados_Formacion"
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
'    cd.InitDir = "c:\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(4).Text = cd.FileName  ' cd.FileTitle
        datos(0).Text = cd.FileTitle
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If txtdatos = "" Then
        MsgBox "Introduzca una descripción, por ejemplo, Curriculum, etc...", vbExclamation, App.Title
        Exit Sub
    End If
    Dim oEF As New clsEmpleados_formacion
    With oEF
        .setFECHA = Format(fecha, "yyyy-mm-dd")
        .setDESCRIPCION = txtdatos
        .Modificar lista.ListItems(lista.selectedItem.Index).Text
    End With
    Set oEF = Nothing
    MsgBox "Registro modificado correctamente.", vbInformation, App.Title
    cargar_lista
End Sub

Private Sub Command1_Click()
   On Error GoTo Command1_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    If lista.ListItems(lista.selectedItem.Index).SubItems(3) = "" Then
        MsgBox "El registro no tiene adjunto.", vbExclamation, App.Title
        Exit Sub
    End If
    Me.MousePointer = 11
    Dim oEF As New clsEmpleados_formacion
    With oEF
        .setRUTA = ""
        .ModificarAdjunto lista.ListItems(lista.selectedItem.Index).Text
    End With
    Set oEF = Nothing
    Me.MousePointer = 0
    MsgBox "El adjunto se ha eliminado correctamente.", vbOKOnly + vbInformation, App.Title
    cargar_lista

   On Error GoTo 0
   Exit Sub

Command1_Click_Error:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command1_Click of Formulario frmEmpleados_Formacion"
End Sub

Private Sub Command2_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Esta seguro de eliminar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oEF As New clsEmpleados_formacion
            oEF.Eliminar CLng(lista.ListItems(lista.selectedItem.Index).Text)
            Set oEF = Nothing
            cargar_lista
        End If
    End If
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargar_lista
    fecha = Date
End Sub

Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oEF As New clsEmpleados_formacion
    Dim objLitem As ListItem
    Set rs = oEF.Listado(PK)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs("ID"))
                .SubItems(1) = Format(rs("FECHA"), "dd-mm-yyyy")
                .SubItems(2) = rs("DESCRIPCION")
                .SubItems(3) = Replace(rs("RUTA"), "/", "\")
                If rs("ruta") <> "" Then
                    Set objLitem = lista.ListItems(lista.ListItems.Count)
                    objLitem.SmallIcon = 1
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "", 300, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnLeft
        .Add , , "Descripcion", lista.Width - 1700, lvwColumnLeft
        .Add , , "Ruta", 1, lvwColumnLeft
    End With
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtdatos = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        fecha = lista.ListItems(lista.selectedItem.Index).SubItems(1)
    End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
'        Dim ruta As String
'        ruta = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\calidad\formacion\" & PK & "\" & lista.ListItems(lista.selectedItem.Index).SubItems(3)
'        If Trim(ruta) = "" Then
'            MsgBox "No tiene asignado adjunto.", vbExclamation, App.Title
'            Exit Sub
'        End If
'        If Dir(ruta) = "" Then
'            MsgBox "El documento adjunto no existe.", vbExclamation, App.Title
'            Exit Sub
'        End If
'        Dim iret As Long
'        iret = ShellExecute(Me.Hwnd, vbNullString, ruta, vbNullString, "c:", SW_SHOWNORMAL)
        If lista.ListItems(lista.selectedItem.Index).SubItems(3) <> "" Then
            Dim oRR As New clsRRHH
            oRR.CargarFormacion CLng(lista.ListItems(lista.selectedItem.Index).Text), True
            Set oRR = Nothing
        End If
    End If
End Sub
'Private Sub adjuntar(ID As Long)
'    If copiar(ID) = False Then
'        Me.MousePointer = 0
'        Exit Sub
'    End If
'    Me.MousePointer = 0
'    Exit Sub
'fallo:
'    Me.MousePointer = 0
'    error_grave "Error al adjuntar el archivo. En Funcion frmEmpleados_Formacion.adjuntar"
'End Sub
'Private Function copiar(EMPLEADO As Long) As Boolean
'    Dim origen As String
'    Dim destino As String
'    origen = datos(4)
'    On Error Resume Next
'    copiar = False
'    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\calidad\formacion"
'    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\calidad\formacion\" & EMPLEADO
'    destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\calidad\formacion\" & EMPLEADO & "\" & datos(0)
'    On Error GoTo fallo
'    If Not gFSO.FileExists(origen) Then
'        error_grave "Error al adjuntar el archivo (Copiar). NO EXISTE EL ORIGEN. " & vbCrLf & " Origen: " & origen & vbCrLf & "Destino: " & destino & vbCrLf
'        copiar = False
'        Exit Function
'    End If
'    If Trim(origen) <> Trim(destino) Then
'        gFSO.CopyFile origen, destino
'    End If
'    If Not gFSO.FileExists(destino) Then
'        error_grave "Error al adjuntar el archivo en funcion frmEmpledos_Formacion(Copiar). NO EXISTE EL DESTINO. " & vbCrLf & " Origen: " & origen & vbCrLf & "Destino: " & destino & vbCrLf
'        copiar = False
'        Exit Function
'    End If
'    copiar = True
'    Exit Function
'fallo:
'    copiar = False
'    error_grave "Error al adjuntar el archivo en funcion frmEmpledos_Formacion.Copiar " & vbCrLf & " Nº Muestra: " & MUESTRA & vbCrLf & "Origen: " & origen & vbCrLf & "Destino: " & destino & vbCrLf & Err.Description
'End Function

