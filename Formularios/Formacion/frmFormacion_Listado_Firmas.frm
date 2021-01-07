VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFormacion_Listado_Firmas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de firmantes"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   Icon            =   "frmFormacion_Listado_Firmas.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   915
      Left            =   8775
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7110
      Width           =   1275
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6135
      Left            =   45
      TabIndex        =   0
      Top             =   900
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   10821
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   45
      Top             =   7425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Listado_Firmas.frx":6852
            Key             =   "estado_ok2"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Listado_Firmas.frx":6C33
            Key             =   "estado_preaviso2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Listado_Firmas.frx":7018
            Key             =   "nada"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Listado_Firmas.frx":735B
            Key             =   "estado_pendiente2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Listado_Firmas.frx":773F
            Key             =   "estado_ok"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Listado_Firmas.frx":7974
            Key             =   "estado_preavisox"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Listado_Firmas.frx":7BFD
            Key             =   "estado_pendiente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Listado_Firmas.frx":7DA1
            Key             =   "estado_preaviso"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Listado_Firmas.frx":8043
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9495
      Picture         =   "frmFormacion_Listado_Firmas.frx":E8A5
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   7110
      Width           =   8550
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   420
      Left            =   45
      TabIndex        =   3
      Top             =   180
      Width           =   7815
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   585
      Width           =   7815
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   915
      Left            =   -45
      Top             =   0
      Width           =   10225
   End
End
Attribute VB_Name = "frmFormacion_Listado_Firmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'M0996: Formulario creado para MANTIS 996


Public PK As Long
Private oCurso As New clsFormacion_cursos

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    
    cabecera
    cargar_lista
    
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID. FIRMA ", 0, lvwColumnLeft
        .Add , , "ID_TIPO", 0, lvwColumnCenter
        .Add , , "TIPO", 1200, lvwColumnCenter
        .Add , , "ID. Empleado ", 0, lvwColumnCenter
        .Add , , "Nombre", 4150, lvwColumnCenter
        .Add , , "Rol", 1500, lvwColumnCenter
        .Add , , "Firma", 600, lvwColumnCenter
        .Add , , "Fecha de firma", 2230, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()

    Dim NoFirmados As Integer

    Dim rs As New ADODB.Recordset
    
    Dim oFirma As New clsFirmas
    Dim oempleado As New clsEmpleados
    
    oCurso.Carga PK
    
    'etiquetas superiores
    'M1106-I
    'If oCurso.getTIPO_MODALIDAD_ID = 0 Then
        lbltitulo.Caption = " Curso: RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & oCurso.getANYO
    'Else
    '    lbltitulo.Caption = " Curso: 0301-" & Format(oCurso.getCOD_CURSO)
    'End If
    'M1106-F
    
    'Carga de la lista
    NoFirmados = 0
    Set rs = oFirma.ListadoCurso(PK)
 
    lista.ListItems.Clear
    lblsubtitulo.Caption = " Se han encontrado " & rs.RecordCount & " firmantes"
    
    If rs.RecordCount <> 0 Then
        Do
            'M1106-I
            Dim oDecodificadora As New clsDecodificadora
            'M1106-F
            
            With lista.ListItems.Add(, , rs("ID_FIRMA"))
                  
                 oempleado.CARGAR rs("USUARIO_ID")
                 'M1106-I
                 oDecodificadora.Carga_valor PARAM_TOBJETO, rs("TOBJETO")
                 'M1106-F
                 
                 .SubItems(1) = rs("TOBJETO")
                 
                 'M1106-I
                 'If rs("TOBJETO") = 17 Then
                 '   .SubItems(2) = "Finalización"
                 'Else
                 '   .SubItems(2) = "Asistencia"
                 'End If
                 '.SubItems(5) = rs("ROL")
                 
                 Select Case rs("TOBJETO")
                 Case TOBJETO_ASISTENCIA_CURSO_ASISTENTES To TOBJETO_ASISTENCIA_CURSO_CALIDAD
                 
                      .SubItems(2) = "Finalización"
                      
                      'Obtención del rol desde la tabla decodificadora.
                      'Este valor deber ir contenido entre paréntesis en la descripción del parámetro.
                      'PosX: primer paréntesis
                      'PosY: segundo paréntesis
                      
                      Dim strRol As String
                      Dim posX As Integer
                      Dim posY As Integer
                    
                      posX = InStr(1, oDecodificadora.getDESCRIPCION, "(") + 1
                      posY = InStr(posX, oDecodificadora.getDESCRIPCION, ")")
                    
                      If posY > posX Then
                         .SubItems(5) = Mid$(oDecodificadora.getDESCRIPCION, posX, posY - posX)
                      End If
                    
                  Case TOBJETO_INVITACION_CURSO
                      .SubItems(2) = "Asistencia"
                      .SubItems(5) = "Alumno"
                      
                  End Select
                 'M1106-F
                 
                 .SubItems(3) = oempleado.getID_EMPLEADO
                 .SubItems(4) = oempleado.getNOMBRE
                 'M1106-F
                 
                 If rs("FIRMADA") = 0 Then
                    .SubItems(6) = " "
                    NoFirmados = NoFirmados + 1
                 Else
                    .SubItems(6) = "Sí"
                    .SubItems(7) = rs("FTIMESTP")
                 End If

            End With
      
            rs.MoveNext
            Set oDecodificadora = Nothing
        Loop Until rs.EOF
    End If
    
    If NoFirmados <> 0 Then
        lblMensaje.Caption = "Quedan " & NoFirmados & " personas por firmar el curso"
    End If
    
    Set oFirma = Nothing
    Set oempleado = Nothing
    Set rs = Nothing
    
End Sub

Private Sub lista_Click()

    If lista.ListItems.Count = 0 Then Exit Sub
    If lista.ListItems(lista.selectedItem.Index).SubItems(1) >= TOBJETO.TOBJETO_ASISTENCIA_CURSO_ASISTENTES And lista.ListItems(lista.selectedItem.Index).SubItems(1) <= TOBJETO.TOBJETO_ASISTENCIA_CURSO_CALIDAD And Trim(lista.ListItems(lista.selectedItem.Index).SubItems(6)) <> "" Then
        lblMensaje.Caption = "Doble click muestra la evaluación del curso"
    Else
        lblMensaje.Caption = " "
    End If
    
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        'M1106-I
        'If lista.ListItems(lista.selectedItem.Index).SubItems(1) = 17 And lista.ListItems(lista.selectedItem.Index).SubItems(6) = "Sí" Then
        'M1341-I
        'If lista.ListItems(lista.selectedItem.Index).SubItems(1) >= TOBJETO.TOBJETO_ASISTENCIA_CURSO_ASISTENTES And lista.ListItems(lista.selectedItem.Index).SubItems(1) <= TOBJETO.TOBJETO_ASISTENCIA_CURSO_CALIDAD And lista.ListItems(lista.selectedItem.Index).SubItems(6) = "Sí" Then
        If lista.ListItems(lista.selectedItem.Index).SubItems(1) = TOBJETO.TOBJETO_ASISTENCIA_CURSO_ASISTENTES And lista.ListItems(lista.selectedItem.Index).SubItems(6) = "Sí" Then
        'M1341-F
        'M1106-F
            frmFormacion_Evaluacion.PK = PK
            frmFormacion_Evaluacion.ID_ASISTENTE = lista.ListItems(lista.selectedItem.Index).SubItems(3)
            frmFormacion_Evaluacion.Show 1
            
        End If
    End If
End Sub
