VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_pruebas_jonathan 
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageCombo icombo 
      Height          =   330
      Left            =   6705
      TabIndex        =   21
      Top             =   810
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
      ImageList       =   "ImageList1"
   End
   Begin MSComctlLib.ListView lista2 
      Height          =   4620
      Left            =   6705
      TabIndex        =   22
      Top             =   1170
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   8149
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      TextBackground  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   7380
      TabIndex        =   20
      Top             =   7290
      Width           =   3765
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Vincular"
      Height          =   555
      Left            =   5940
      TabIndex        =   19
      Top             =   6930
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6345
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Limpiar"
      Height          =   420
      Left            =   4320
      TabIndex        =   18
      Top             =   7245
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Añadir"
      Height          =   465
      Left            =   4320
      TabIndex        =   17
      Top             =   6750
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selección del Certificado Externo"
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
      Height          =   1230
      Left            =   4320
      TabIndex        =   9
      Top             =   5490
      Width           =   7890
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar"
         Height          =   825
         Left            =   6930
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   270
         Width           =   810
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   825
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   540
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   270
         Width           =   3765
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   825
         Index           =   0
         Left            =   4365
         Picture         =   "frm_Pruebas_Jonathan.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   270
         Width           =   810
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar"
         Height          =   825
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   270
         Width           =   810
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escaner"
         Height          =   825
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   270
         Width           =   825
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ruta"
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
         Index           =   8
         Left            =   90
         TabIndex        =   15
         Top             =   585
         Width           =   405
      End
   End
   Begin VB.TextBox txtRutaXml 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4170
      TabIndex        =   8
      Text            =   "c:\archivo.xml"
      Top             =   5160
      Width           =   10755
   End
   Begin VB.TextBox txtSQL 
      Appearance      =   0  'Flat
      Height          =   1245
      Left            =   4170
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Text            =   "frm_Pruebas_Jonathan.frx":030A
      Top             =   3900
      Width           =   10755
   End
   Begin VB.CommandButton cmdSQL2XML 
      Caption         =   "Exportar SQL a XML Local"
      Height          =   1545
      Left            =   120
      TabIndex        =   6
      Top             =   3900
      Width           =   4005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   525
      Left            =   4710
      TabIndex        =   5
      Top             =   780
      Width           =   1245
   End
   Begin VB.CommandButton cmdEquiposMtoPte 
      Caption         =   "4.-Equipos con Mantenimientos Pendientes"
      Height          =   765
      Left            =   120
      TabIndex        =   4
      Top             =   3090
      Width           =   4005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3.-Desmarcar CVM cuando no tienen"
      Height          =   765
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   4005
   End
   Begin VB.CommandButton cmdCrearOperacionesPendientes 
      Caption         =   "2.-Crear Operaciones Pendientes Equipos"
      Height          =   765
      Left            =   120
      TabIndex        =   2
      Top             =   1470
      Width           =   4005
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1.-Revisar Firmas en Recepcion de muestras"
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   750
      Width           =   4005
   End
   Begin MSComctlLib.ListView lista 
      Height          =   2100
      Left            =   135
      TabIndex        =   16
      Top             =   5490
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   3704
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6345
      Top             =   7245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   5460
      Top             =   6510
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "En este formulario hay botones para lanzar distintos scripts"
      Height          =   525
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   14805
   End
End
Attribute VB_Name = "frm_pruebas_jonathan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub adjuntar()
   On Error GoTo cmdVincular_Click_Error

    If datos(0) = "" Then
        MsgBox "Por favor, indique el certificado a vincular.", vbExclamation, App.Title
        Exit Sub
    End If
    If Dir(datos(0)) = "" Then
        MsgBox "El documento vinculado no existe en la ruta.", vbExclamation, App.Title
        Exit Sub
    End If
    On Error Resume Next
    Dim RUTA As String
    RUTA = ReadINI(App.Path + "\config.ini", "Documentos", "rex_certificados")
    MkDir RUTA & "\" & CStr(PK)
    On Error GoTo cmdVincular_Click_Error
    FileCopy datos(0), RUTA & "\" & CStr(PK) & "\" & datos(1)
    Dim oBote As New clsBotes_ex
    oBote.setCERTIFICADO_EXTERNO = RUTA & "\" & CStr(PK) & "\" & datos(1)
    oBote.InformarRutaCertificado CLng(PK)
    Set oBote = Nothing
    datos(0) = RUTA & "\" & CStr(PK) & "\" & datos(1)
    MsgBox "Certificado vinculado correctamente.", vbInformation, App.Title
   On Error GoTo 0
   Exit Sub

cmdVincular_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdVincular_Click of Formulario frmREX_Bote_Modificacion"
    

End Sub

Private Sub cmdEscaner_Click()
    frmEscaner.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = ""
        nombreNuevo = InputBox("Introduzca nombre para el archivo Escaneado, sin ninguna extensión (SOLO LETRAS Y NUMEROS).", "Escaneando Archivo", , Screen.Width / 3, (Screen.Height / 3))
        nombreNuevo = Eliminar_Caracteres_Archivo(nombreNuevo)
        If Trim(nombreNuevo) <> "" Then
            datos(0).Text = documento_escaner
            datos(1).Text = nombreNuevo & ".pdf"
            adjuntar
        End If
    End If
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
'    cd.InitDir = "c:\"
    cd.ShowOpen
    If cd.FileName <> "" Then
'        datos(4).Text = cd.FileTitle 'cd.FileName  '
        datos(0).Text = cd.FileName
        datos(1).Text = cd.FileTitle
    End If
End Sub

Private Sub Command3_Click()
    Dim s As String
    s = InputBox("Numero:")
    With lista.ListItems.Add(, , s)
    End With
End Sub

Private Sub Command4_Click()
    lista.ListItems.Clear
End Sub

Private Sub Command5_Click()
    Dim I As Integer
    Dim PK As Long
    Dim RUTA As String
    RUTA = ReadINI(App.Path + "\config.ini", "Documentos", "rex_certificados")
        Dim oBote As New clsBotes_ex
    For I = 1 To lista.ListItems.Count
    PK = lista.ListItems(I).Text
        On Error Resume Next
        MkDir RUTA & "\" & CStr(PK)
        FileCopy datos(0), RUTA & "\" & CStr(PK) & "\" & datos(1)
        oBote.setCERTIFICADO_EXTERNO = RUTA & "\" & CStr(PK) & "\" & datos(1)
        oBote.InformarRutaCertificado CLng(PK)
'        datos(0) = RUTA & "\" & CStr(PK) & "\" & datos(1)
    Next
        Set oBote = Nothing
    MsgBox "Certificado vinculado correctamente.", vbInformation, App.Title

End Sub

Private Sub datos_GotFocus(Index As Integer)
    datos(Index).BackColor = &HC0FFFF
End Sub

Private Sub datos_LostFocus(Index As Integer)
    datos(Index).BackColor = vbWhite
End Sub
Private Sub cmd1_Click()
Dim rs As ADODB.RecordSet
Dim RUTA As String, tx As TextStream
Dim cad As String
Dim cont As Long
Dim Conversor As Class1
RUTA = "\\servidor\Canagrosa\documentos\FIRMAS\"
'ruta = "c:\temp\FIRMAS\"

'cad = "update muestras set firma='' where firma <> ''"
cont = 0

    'execute_bd cad

    'cad = "select id_muestra, firma, fecha_recepcion, cliente_id from muestras where firma <> '' and id_muestra in (99523,99518) order by id_muestra asc"
    cad = "select id_muestra, firma, fecha_recepcion, cliente_id from muestras where firma <> '' order by id_muestra asc"
    Set rs = datos_bd(cad)
Set Conversor = New Class1
    
    Set tx = gFSO.CreateTextFile(App.Path & "\arreglo_firma_muestras.sql", True)
    
    rs.MoveFirst
    While Not rs.EOF
        DoEvents
        If gFSO.FileExists(RUTA & CStr(rs!ID_MUESTRA) & ".bmp") Then
            tx.WriteLine "update muestras set firma = '" & rs!ID_MUESTRA & ".jpg' where id_muestra = " & CStr(rs!ID_MUESTRA) & ";" & vbCrLf
            
            Set Image1.Picture = LoadPicture(RUTA & CStr(rs!ID_MUESTRA) & ".bmp")
            Conversor.GrabarJpg Image1, RUTA & CStr(rs!ID_MUESTRA) & ".jpg", CByte(70)
            DoEvents
            cont = cont + 1
        End If
        rs.MoveNext
    Wend

   
    tx.Close
    Set Conversor = Nothing
    Set tx = Nothing
    Set rs = Nothing
    MsgBox "FIN: " & cont & " muestras procesadas"

    

End Sub




Private Sub cmdCrearOperacionesPendientes_Click()

' comienza con las calibraciones

Dim oOP As New clsEquiposOperacionesPendientes
Dim rs As ADODB.RecordSet, rs2 As ADODB.RecordSet
Dim consulta As String
Dim ocal As New clsEquipoCalibracion, over As New clsEquipoVerificacion
Dim lista() As String

' se hará de la siguiente manera
' 1.- se buscan la última calibracion realizada por equipo y periodicidad
' 2.- a partir de esta, según la fecha próxima que tenga puesta, se crea la siguiente (prevista) y la tarea pendiente

consulta = ""
consulta = consulta & " select * from eq_calibracion_equipos where id_calibracion in("
consulta = consulta & " select id from ("
consulta = consulta & " select max(id_calibracion) id, equipo_id, periodicidad_id"
consulta = consulta & " from eq_calibracion_equipos where estado=1 and periodicidad_id>0"
consulta = consulta & " group by equipo_id, periodicidad_id"
consulta = consulta & " order by equipo_id) as t)"
Set rs = datos_bd(consulta)


If rs.RecordCount <> 0 Then
    rs.MoveFirst
    While Not rs.EOF
        'If rs!EQUIPO_ID = 28 Then
        '    MsgBox ""
        'End If
        ocal.Carga rs!ID_CALIBRACION
        ocal.crear_calibracion_pendiente rs!ID_CALIBRACION, rs!EQUIPO_ID
        rs.MoveNext
        DoEvents
    Wend
End If


' Para las verificaciones
consulta = ""

consulta = consulta & " select * from eq_verificacion_equipos where id_verificacion in("
consulta = consulta & " select id from ("
consulta = consulta & " select max(id_verificacion) id, equipo_id, periodicidad_id"
consulta = consulta & " from eq_verificacion_equipos where estado=1 and periodicidad_id>0"
consulta = consulta & " group by equipo_id, periodicidad_id"
consulta = consulta & " order by equipo_id) as t)"
Set rs = datos_bd(consulta)


If rs.RecordCount <> 0 Then
    rs.MoveFirst
    While Not rs.EOF
        'If rs!EQUIPO_ID = 28 Then
        '    MsgBox ""
        'End If

        over.Carga rs!ID_VERIFICACION
        over.crear_verificacion_pendiente rs!ID_VERIFICACION, rs!EQUIPO_ID
        rs.MoveNext
        DoEvents
    Wend
End If


'MANTENIMIENTOS
ReDim Preserve lista(0)
Set rs = datos_bd("select distinct equipo_id, planmto_id from eq_mantenimiento_equipos where planmto_id <> 0 order by equipo_id, planmto_id ")
If rs.RecordCount <> 0 Then
    rs.MoveFirst
    While Not rs.EOF
        If UBound(lista) < rs!EQUIPO_ID Then ReDim Preserve lista(rs!EQUIPO_ID)
        lista(rs!EQUIPO_ID) = lista(rs!EQUIPO_ID) & ":" & CStr(rs!PLANMTO_ID) & ":"
        rs.MoveNext
    Wend
End If

Set rs = datos_bd("select distinct equipo_id from eq_mantenimiento_equipos where planmto_id <> 0 order by equipo_id")
If rs.RecordCount <> 0 Then
    rs.MoveFirst
    While Not rs.EOF
        oOP.crear_mantenimientos_pendiente_para_nuevas_fechas_planes rs!EQUIPO_ID, lista(rs!EQUIPO_ID)
        rs.MoveNext
    Wend
End If


MsgBox "fin"


End Sub


Private Sub cmdEquiposMtoPte_Click()

    frmEquipoListadoMtoPte.Show 1
    
End Sub

Private Sub cmdSQL2XML_Click()
Dim rs As New ADODB.RecordSet
Dim obj_DOMDocument As DOMDocument60

    ' Graba el contenido del Recordset en el Obj DOMDocument.
'    Set rs = datos_bd(Replace(txtSQL.Text, vbCrLf, ""))
    Set rs = datos_bd(txtSQL.Text)
    
    Set obj_DOMDocument = New DOMDocument60
    rs.Save obj_DOMDocument, adPersistXML
      
    'Cierra el recordset y la conexión a la base de datos
    Set rs = Nothing
    
      
    ' Genera el archivo xml
    obj_DOMDocument.Save txtRutaXml.Text
    
    
    MsgBox "Salvado"

End Sub

Private Sub Command1_Click()
Dim consulta As String

consulta = "update equipos Set CON_CALIBRACION = 0 where id_equipo not in (select distinct equipo_id from eq_calibracion_equipos) and CON_CALIBRACION = 1"
Set rs = datos_bd(consulta)

consulta = "update equipos Set CON_VERIFICACION = 0 where id_equipo not in (select distinct equipo_id from eq_verificacion_equipos) and CON_VERIFICACION = 1"
Set rs = datos_bd(consulta)


consulta = "update equipos Set CON_MANTENIMIENTO = 0 where id_equipo not in (select distinct equipo_id from eq_mantenimiento_equipos) and CON_mantenimiento = 1"
Set rs = datos_bd(consulta)

MsgBox "Fin"
End Sub




Private Sub txtTb_Change()



With grdResultados
    .Columns(1).RefetchCell
grdResultados.Text = txtTb.Text
End With
End Sub


Private Sub Command2_Click()
Dim rs As ADODB.RecordSet

Set rs = datos_bd("select * from ca_documentos where ruta <> ''")
If rs.RecordCount > 0 Then

    Do
        origen = rs("ruta")
'        destino = Replace(rs("ruta"), "/SOFTWARE/", "/FORMATO/")
        If Not gFSO.FileExists(Replace(origen, "/", "\")) Then
            MsgBox rs("nombre")
        End If
'        Kill destino
'        FileCopy Replace(origen, "/", "\"), Replace(destino, "/", "\")
'        execute_bd "update ca_documentos set ruta = '" & destino & "' where id_documento = " & rs("id_documento")
        rs.MoveNext
    Loop Until rs.EOF
End If


End Sub


Private Sub Form_Load()
    With lista.ColumnHeaders.Add(, , "BOTE", 1200, lvwColumnLeft)
        .Tag = "BOTE"
    End With
'    With lista2.ColumnHeaders
'        .Add , , "Nombre1", 2200, lvwColumnLeft
'        .Add , , "Nombre2", 2200, lvwColumnLeft
'        .Add , , "Nombre3", 2200, lvwColumnLeft
'    End With
    icombo.ComboItems.Add , , "PRUEBA1", 1
    icombo.ComboItems.Add , , "PRUEBA2", 2
    With lista2.ListItems.Add(, , "PRUEBA1", 1)
        .Tag = 1
    End With
    With lista2.ListItems.Add(, , "PRUEBA2", 2)
        .Tag = 2
    End With
    With lista2.ListItems.Add(, , "PRUEBA3", 3)
        .Tag = 3
    End With

End Sub

Private Sub lista2_DblClick()
    If lista2.ListItems.Count > 0 Then
        MsgBox lista2.ListItems(lista2.SelectedItem.Index).Tag
    End If

End Sub
