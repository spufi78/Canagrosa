VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAlodine_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Lotes de Alodine"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   Icon            =   "frmAlodine_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   13215
   Begin VB.CommandButton cmdClientes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clientes"
      Height          =   870
      Left            =   6795
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7770
      Width           =   1275
   End
   Begin VB.Frame Frame1 
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
      Height          =   1005
      Left            =   45
      TabIndex        =   13
      Top             =   675
      Width           =   13155
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   855
         TabIndex        =   0
         Top             =   405
         Width           =   1545
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   3420
         TabIndex        =   1
         Top             =   405
         Width           =   1770
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   825
         Left            =   12105
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1005
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   7080
         TabIndex        =   2
         Top             =   405
         Width           =   4665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "NºLOTE"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   16
         Top             =   450
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   195
         Index           =   0
         Left            =   6300
         TabIndex        =   15
         Top             =   450
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   1
         Left            =   2790
         TabIndex        =   14
         Top             =   450
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdTerminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lote Terminado"
      Height          =   870
      Left            =   5490
      Picture         =   "frmAlodine_Listado.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7770
      Width           =   1275
   End
   Begin VB.CheckBox chkterminado 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incluir los lotes terminados"
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
      Height          =   285
      Left            =   8445
      TabIndex        =   11
      Top             =   8055
      Width           =   3105
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7770
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7770
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7770
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7770
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7770
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7770
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6015
      Left            =   60
      TabIndex        =   4
      Top             =   1710
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   10610
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
      Caption         =   "Mantenimiento de Lotes de Alodine"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   18
      Top             =   360
      Width           =   2490
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12645
      Picture         =   "frmAlodine_Listado.frx":1B3C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Lotes Alodine"
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
      Left            =   45
      TabIndex        =   17
      Top             =   45
      Width           =   2595
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   13185
   End
End
Attribute VB_Name = "frmAlodine_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkterminado_Click()
    cargar_lista
End Sub

Private Sub cmdClientes_Click()
    If lista.ListItems.Count > 0 Then
        gAlodine = lista.ListItems(lista.selectedItem.Index).Text
        frmAlodine_Clientes.Show 1
        actualizar_lista
        gAlodine = 0
    End If
End Sub

Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
'    If lista.ListItems(lista.SelectedItem.Index).SubItems(7) = 0 Then
'        MsgBox "Solo se pueden duplicar LOTES/MADRES que se encuentran terminados.", vbExclamation, App.Title
'        Exit Sub
'    End If
        
    If MsgBox("Va a duplicar el tipo de Alodine. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
      Dim alodine As Long
      Dim oalodine As New clsAlodine
      Dim oAlodine_nuevo As New clsAlodine
      Dim RS As ADODB.Recordset
      If oalodine.Carga(lista.ListItems(lista.selectedItem.Index).Text) = True Then
          With oAlodine_nuevo
            .setPRODUCTO = oalodine.getPRODUCTO
            .setCODIGO = oalodine.getCODIGO
'            .setLOTE = ""
            .setLOTE = oalodine.getLOTE
            
            .setDESCRIPCION = oalodine.getDESCRIPCION
            .setPROCEDIMIENTO = oalodine.getPROCEDIMIENTO
            .setTIPO_CADUCIDAD_ID = oalodine.getTIPO_CADUCIDAD_ID
            .setFECHA_CREACION = Format(Date, "yyyy-mm-dd")
            .setFECHA_TERMINACION = "0000-00-00"
            .setTERMINADO = 0
            alodine = .Insertar
            If alodine = 0 Then
               MsgBox "Error al insertar el alodine duplicado.", vbCritical, App.Title
               Exit Sub
            End If
          End With
          ' ALODINE_CLIENTES
          Dim oAC As New clsAlodine_clientes
          Set RS = oAC.Listado(lista.ListItems(lista.selectedItem.Index).Text)
          Do While Not RS.EOF
             With oAC
                .setALODINE_ID = alodine
                .setCLIENTE_ID = RS("CLIENTE_ID")
                .setEADS = RS("EADS")
                .setCAPACIDAD_ID = RS("CAPACIDAD_ID")
                .setETIQUETA_ID = RS("ETIQUETA_ID")
                .setNUMERO_BOTES = RS("NUMERO_BOTES")
                .setPRECIO = Replace(Format(RS("PRECIO"), "0.00"), ",", ".")
                .setPEDIDO = RS("PEDIDO")
                .setPEDIDO_ID = RS("PEDIDO_ID")
                .setNORMA = RS("NORMA")
                .setNORMA_ETIQUETA = RS("NORMA_ETIQUETA")
                If .Insertar = 0 Then
                    MsgBox "Error al insertar los clientes del alodine.", vbCritical, App.Title
                    Exit Sub
                End If
             End With
            RS.MoveNext
          Loop
          ' ALODINE_PARAMETROS
          Dim OAP As New clsAlodine_parametros
          Set RS = OAP.Listado(lista.ListItems(lista.selectedItem.Index).Text)
          Do While Not RS.EOF
             With OAP
                .setID_PARAMETRO = 0
                .setALODINE_ID = alodine
                .setPARAMETRO = RS("PARAMETRO")
                .setRANGO = RS("RANGO")
                .setPROCEDIMIENTO = RS("PROCEDIMIENTO")
                .setUNIDAD_ID = RS("UNIDAD_ID")
                
                If .Insertar = 0 Then
                    MsgBox "Error al insertar los parametros del alodine.", vbCritical, App.Title
                    Exit Sub
                End If
             End With
            RS.MoveNext
          Loop
          ' ALODINE_NORMAS
          Dim oAN As New clsAlodine_normas
          oAN.duplicar lista.ListItems(lista.selectedItem.Index).Text, alodine
          Set oAN = Nothing
          ' ALODINE_ETIQUETAS
          Dim oAE As New clsAlodine_etiqueta
          oAE.Carga lista.ListItems(lista.selectedItem.Index).Text, "ES"
          oAE.setALODINE_ID = alodine
          oAE.Insertar
          oAE.Carga lista.ListItems(lista.selectedItem.Index).Text, "EN"
          oAE.setALODINE_ID = alodine
          oAE.Insertar
          Set oAE = Nothing
          MsgBox "El alodine se ha duplicado correctamente.NO OLVIDE INFORMAR EL LOTE/MADRE.", vbOKOnly + vbInformation, App.Title
          cargar_lista
      End If
    End If
    Exit Sub
fallo:
    MsgBox "Error en el proceso de duplicación.", vbCritical, App.Title
End Sub

Private Sub cmdAnadir_Click()
    gAlodine = 0
    frmAlodine_Alodine.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el Alodine : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oalodine As New clsAlodine
            If oalodine.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub
'Private Sub cmdImprimir_Click_old()
'    On Error GoTo fallo
'    Dim i As Integer
'    ' Generamos los datos del listado
'    Dim rs As New ADODB.RecordSet
'    rs.Fields.Append "c1", adChar, 50, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 50, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 15, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 5, adFldUpdatable
'    rs.Open
'    For i = 1 To lista.ListItems.Count
'        rs.AddNew
'        rs("c1") = Left(lista.ListItems(i), 50)
'        If Trim(lista.ListItems(i).SubItems(1)) <> "" Then
'            rs("c2") = Left(lista.ListItems(i).SubItems(1), 50)
'        End If
'        If Trim(lista.ListItems(i).SubItems(2)) <> "" Then
'            rs("c3") = Left(lista.ListItems(i).SubItems(2), 15)
'        End If
'        If Trim(lista.ListItems(i).SubItems(3)) <> "" Then
'            rs("c4") = Left(lista.ListItems(i).SubItems(3), 5)
'        End If
'        rs.Update
'    Next
'    ' Generar Listado
'    Dim Listado As New rptListado
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Tipos de Alodine"
'        .Controls("etiqueta4").Caption = "ID"
'        .Controls("etiqueta5").Caption = "Producto"
'        .Controls("etiqueta10").Caption = "Codigo"
'        .Controls("etiqueta11").Caption = "Procedimiento"
'    End With
'    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("d1").DataField = rs.Fields("c4").Name
'        .Controls("d2").DataField = rs.Fields("c1").Name
'        .Controls("d3").DataField = rs.Fields("c2").Name
'        .Controls("d4").DataField = rs.Fields("c3").Name
'    End With
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & Usuario.getNOMBRE
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Tipos de Alodine"
''    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado.", vbCritical, Err.Description
'End Sub

Private Sub cmdImprimir_Click()
    Dim objAlo As New clsAlodine
    
    objAlo.Imprimir_Listado
    
    Set objAlo = Nothing
    
End Sub

Private Sub cmdLimpiar_Click()
    txtFiltro(0) = ""
    txtFiltro(1) = ""
    txtFiltro(2) = ""
    cargar_lista
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        gAlodine = lista.ListItems(lista.selectedItem.Index).Text
        frmAlodine_Alodine.Show 1
        actualizar_lista
        gAlodine = 0
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTerminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a terminar el Alodine : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oalodine As New clsAlodine
            If oalodine.Terminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
    
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "NºLOTE", 1000, lvwColumnCenter)
        .Tag = "NºLOTE"
    End With
    With lista.ColumnHeaders.Add(, , "Producto", 4100, lvwColumnLeft)
        .Tag = "Producto"
    End With
    With lista.ColumnHeaders.Add(, , "Codigo", 1800, lvwColumnCenter)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Lote Componente", 2350, lvwColumnCenter)
        .Tag = "Lote"
    End With
    With lista.ColumnHeaders.Add(, , "Procedimiento", 1400, lvwColumnCenter)
        .Tag = "Procedimiento"
    End With
    With lista.ColumnHeaders.Add(, , "F.Creación", 1050, lvwColumnCenter)
        .Tag = "F.Creación"
    End With
    With lista.ColumnHeaders.Add(, , "F.Finalización", 1050, lvwColumnCenter)
        .Tag = "F.Finalización"
    End With
    With lista.ColumnHeaders.Add(, , "TERMINADO", 0, lvwColumnCenter)
        .Tag = "TERMINADO"
    End With
End Sub

Private Sub cargar_lista()
    Dim RS As New ADODB.Recordset
    Dim oalodine As New clsAlodine
    lista.ListItems.Clear
    Set RS = oalodine.Listado(chkterminado.Value, txtFiltro(0), txtFiltro(1), txtFiltro(2))
    If RS.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(RS("ID_ALODINE"), "0000"))
             .SubItems(1) = Format(RS("MADRE"), "0000")
             .SubItems(2) = RS("PRODUCTO")
             .SubItems(3) = RS("CODIGO")
             .SubItems(4) = RS("LOTE")
             .SubItems(5) = RS("PROCEDIMIENTO")
             .SubItems(6) = Format(RS("FECHA_CREACION"), "DD-MM-YYYY")
             If RS("terminado") = 1 Then
                .SubItems(7) = Format(RS("FECHA_TERMINACION"), "DD-MM-YYYY")
             Else
                .SubItems(7) = ""
             End If
             .SubItems(8) = RS("TERMINADO")
            End With
            RS.MoveNext
        Loop Until RS.EOF
    End If
    Set oalodine = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
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
Public Sub actualizar_lista()
    Dim oalodine As New clsAlodine
    With oalodine
        If .Carga(lista.ListItems(lista.selectedItem.Index).Text) = True Then
            lista.ListItems(lista.selectedItem.Index).SubItems(1) = Format(.getMADRE, "0000")
            lista.ListItems(lista.selectedItem.Index).SubItems(2) = .getPRODUCTO
            lista.ListItems(lista.selectedItem.Index).SubItems(3) = .getCODIGO
            lista.ListItems(lista.selectedItem.Index).SubItems(4) = .getLOTE
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = .getPROCEDIMIENTO
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = Format(.getFECHA_CREACION, "DD-MM-YYYY")
            If .getTERMINADO = 1 Then
                lista.ListItems(lista.selectedItem.Index).SubItems(7) = Format(.getFECHA_TERMINACION, "DD-MM-YYYY")
            End If
            lista.ListItems(lista.selectedItem.Index).SubItems(8) = .getTERMINADO
        End If
    End With
    Set oalodine = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
