VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCE_Imagenes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Imagenes de la Muestra"
   ClientHeight    =   10905
   ClientLeft      =   60
   ClientTop       =   195
   ClientWidth     =   12645
   Icon            =   "frmCE_Imagenes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10905
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicOriginal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7155
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   14
      Top             =   7650
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   45
      ScaleHeight     =   6435
      ScaleWidth      =   12510
      TabIndex        =   13
      Top             =   630
      Width           =   12570
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documento adjunto"
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
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   9765
      Width           =   11355
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1035
         TabIndex        =   12
         Top             =   45
         Visible         =   0   'False
         Width           =   7065
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Anular"
         Height          =   870
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   960
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar"
         Height          =   870
         Left            =   9315
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   960
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1035
         TabIndex        =   7
         Top             =   630
         Width           =   7065
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   285
         Index           =   0
         Left            =   8145
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   915
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1035
         TabIndex        =   0
         Top             =   315
         Width           =   7065
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fichero"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   6
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Leyenda"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   675
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11565
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9900
      Width           =   1005
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8370
      Top             =   9900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lista 
      Height          =   2550
      Left            =   90
      TabIndex        =   3
      Top             =   7155
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   4498
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Se muestras las imagenes adjuntas a la muestra"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   9
      Top             =   315
      Width           =   3375
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imagenes adjuntas "
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
      TabIndex        =   8
      Top             =   45
      Width           =   2040
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   12705
   End
End
Attribute VB_Name = "frmCE_Imagenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Dim ancho As Single, alto As Single, porcentaje As Single
Dim imagen As IPictureDisp


Private Sub cmdAdjuntar_Click()
    On Error GoTo fallo
    Me.MousePointer = 11
    ' Validar
    If validar = False Then
        Me.MousePointer = 0
        Exit Sub
    End If
    Dim ruta As String
    Dim nombre As String
    ruta = datos(2)
    nombre = datos(0)
    Dim Insertar As Boolean
    Insertar = True
    ' Convertir a jpg si es un bmp
    If UCase(Right(nombre, 3)) = "BMP" Then
        PicOriginal.Picture = LoadPicture(ruta) 'Cargamos el Picture
        nombre = Replace(nombre, ".bmp", ".jpg")
        ruta = Replace(ruta, ".bmp", ".jpg")
        Dim Conversor As Class1
        Set Conversor = New Class1
        Conversor.GrabarJpg PicOriginal.Image, ruta, CByte(70)
        Set Conversor = Nothing
    End If
    If Insertar Then
        Dim od As New clsDocumentacion
'        salida = od.SubirMuestraImagen(PK, lista.ListItems.Count + 1, ruta, nombre, datos(1))
        salida = od.SubirMuestraImagen(PK, ruta, nombre, datos(1))
        Set od = Nothing
        If salida <> "" Then
            Me.MousePointer = 0
            MsgBox "Error al adjuntar el archivo : " & salida, vbExclamation, App.Title
            Exit Sub
        End If
    End If
    
    Dim oce_recepcion As New clsCe_recepcion
    oce_recepcion.Incluir_Imagenes PK, 1
    Set oce_recepcion = Nothing
            
    Me.MousePointer = 0
    datos(0) = ""
    datos(1) = ""
    datos(2) = ""
    MsgBox "El archivo se ha adjuntado correctamente.", vbOKOnly + vbInformation, App.Title
    
    cargar_lista
    Exit Sub
fallo:
    Me.MousePointer = 0
    error_grave "Error al adjuntar el archivo. " & Err.Description

End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Seguro de anular la imagen?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim od As New clsDocumentacion
            od.EliminarMuestrasImagen PK, CInt(lista.ListItems(lista.selectedItem.Index).Text)
            Set od = Nothing
'            Dim oCE_imagenes As New clsCe_imagenes
'            oCE_imagenes.Eliminar PK, CInt(lista.ListItems(lista.selectedItem.Index).Text)
'            Set oCE_imagenes = Nothing
            cargar_lista
            If lista.ListItems.Count = 0 Then
                Dim oce_recepcion As New clsCe_recepcion
                oce_recepcion.Incluir_Imagenes PK, 0
                Set oce_recepcion = Nothing
                Picture1.Cls
            End If
        End If
    End If
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir imagen"
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(0).Text = cd.FileTitle
        datos(2).Text = cd.FileName
        Set previa.Picture = LoadPicture(datos(2).Text)
    End If
End Sub

Private Sub datos_GotFocus(Index As Integer)
    datos(Index).BackColor = &HC0FFFF
    datos(Index).SelStart = 0
    datos(Index).SelLength = Len(datos(Index))
    
End Sub

Private Sub datos_LostFocus(Index As Integer)
    datos(Index).BackColor = vbWhite
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_lista
End Sub

Private Sub cabecera()
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra PK
    lbltitulo(0) = "Imagenes adjuntas a la muestra : " & oMuestra.getID_GENERAL & "/" & oMuestra.getANNO
    Set oMuestra = Nothing
    With lista.ColumnHeaders
        .Add , , "Orden", 1, lvwColumnLeft
        .Add , , "Leyenda", (lista.Width / 2) - 150, lvwColumnCenter
        .Add , , "Fichero", (lista.Width / 2) - 150, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    If PK > 0 Then
'        Dim oCE_imagenes As New clsCe_imagenes
        Dim oDocumentacion As New clsDocumentacion
        Dim rs As adodb.Recordset
'        Set rs = oCE_imagenes.Listado(PK)
        Set rs = oDocumentacion.ListadoMuestrasImagenes(PK)
        lista.ListItems.Clear
        If rs.RecordCount > 0 Then
            Do
                With lista.ListItems.Add(, , rs("ORDEN"))
                     .SubItems(1) = rs("LEYENDA")
'                     .SubItems(2) = rs("RUTA")
                     .SubItems(2) = rs("FILE_NAME")
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        If lista.ListItems.Count > 0 Then
            lista_Click
        End If
    End If
End Sub
Private Function validar() As Boolean
    validar = False
    If datos(0) = "" Then
        MsgBox "Escriba una ruta.", vbExclamation, App.Title
        datos(0).SetFocus
        Exit Function
    End If
    If Dir(datos(0)) = "" Then
        MsgBox "La ruta introducida no existe.", vbExclamation, App.Title
        datos(0).SetFocus
        Exit Function
    End If
    If datos(1) = "" Then
        MsgBox "Escriba una leyenda para la imagen.", vbExclamation, App.Title
        datos(1).SetFocus
        Exit Function
    End If
    validar = True
End Function

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        On Error Resume Next
        Dim oDoc As New clsDocumentacion
        Dim ruta As String
        ruta = oDoc.CargarMuestrasImagen(PK, lista.ListItems(lista.selectedItem.Index).Text, False)
'        ruta = Replace(lista.ListItems(lista.selectedItem.Index).SubItems(2), "/", "\")
        If Dir(ruta) <> "" Then
        
            Picture1.AutoRedraw = True
'            Picture1.PaintPicture LoadPicture(ruta), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
'            Set previa.Picture = LoadPicture(ruta)

            
            Picture1.Cls
            Set imagen = LoadPicture(ruta)
            ancho = imagen.Width
            alto = imagen.Height
            
            If ancho > Picture1.Width Or alto > Picture1.Height Then
                If ancho - 1000 > alto Then
                    porcentaje = (Picture1.Width * 100) / ancho
                Else
                    porcentaje = (Picture1.Height * 100) / alto
                End If
                CentrarPicture
                Exit Sub
            End If
            
            If ancho <= Picture1.Width Or alto <= Picture1.Height Then
            If ancho > alto Then
            porcentaje = (Picture1.Width * 100) / ancho
            Else
            porcentaje = (Picture1.Width * 100) / alto
            End If
            CentrarPicture
            End If
        End If
    End If
End Sub


Public Sub CentrarPicture()
    Dim centro1 As Single, centro2 As Single
    ancho = (ancho * porcentaje) / 100
    alto = (alto * porcentaje) / 100
    centro1 = (Picture1.Width - ancho) / 2
    centro2 = (Picture1.Height - alto) / 2
    Picture1.PaintPicture imagen, centro1, centro2, ancho, alto
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        Dim oDoc As New clsDocumentacion
        oDoc.CargarMuestrasImagen PK, lista.ListItems(lista.selectedItem.Index).Text, True
        Set oDoc = Nothing
    End If
End Sub
