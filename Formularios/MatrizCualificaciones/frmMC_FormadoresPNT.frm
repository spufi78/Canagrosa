VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmMC_FormadoresPNT 
   Caption         =   "Usuarios Formadores en PNT"
   ClientHeight    =   10905
   ClientLeft      =   2790
   ClientTop       =   1035
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10905
   ScaleWidth      =   14460
   Begin VB.CommandButton cmdImprimir2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Formadores PNT"
      Height          =   915
      Left            =   11070
      Picture         =   "frmMC_FormadoresPNT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9960
      Width           =   1080
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PNT Formadores"
      Height          =   915
      Left            =   12180
      Picture         =   "frmMC_FormadoresPNT.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9960
      Width           =   1080
   End
   Begin MSFlexGridLib.MSFlexGrid glista 
      Height          =   8145
      Left            =   0
      TabIndex        =   7
      Top             =   1770
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   14367
      _Version        =   393216
      BackColor       =   12640511
      BackColorFixed  =   -2147483639
      BackColorSel    =   8553090
      BackColorBkg    =   12632256
      HighLight       =   2
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   1500
      Top             =   10200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMC_FormadoresPNT.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMC_FormadoresPNT.frx":164D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   13290
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9960
      Width           =   1125
   End
   Begin VB.Frame fraDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   14475
      Begin VB.TextBox txtFiltroNombre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3150
         TabIndex        =   13
         Top             =   960
         Width           =   6975
      End
      Begin VB.CommandButton cmdQuitarFiltro 
         Caption         =   "Quitar Filtro"
         Height          =   345
         Left            =   11730
         TabIndex        =   11
         Top             =   930
         Width           =   1515
      End
      Begin VB.TextBox txtFiltroCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3150
         TabIndex        =   10
         Top             =   660
         Width           =   6975
      End
      Begin VB.CommandButton cmdFiltrar 
         Caption         =   "Aplicar Filtro"
         Height          =   345
         Left            =   10170
         TabIndex        =   9
         Top             =   930
         Width           =   1515
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar Usuario Seleccionado"
         Height          =   525
         Left            =   12870
         TabIndex        =   5
         Top             =   210
         Width           =   1515
      End
      Begin VB.CommandButton cmdAnadir 
         Caption         =   "Añadir Usuario Seleccionado"
         Height          =   525
         Left            =   11310
         TabIndex        =   4
         Top             =   210
         Width           =   1515
      End
      Begin pryCombo.miCombo cmbUsuarios 
         Height          =   315
         Left            =   1590
         TabIndex        =   3
         Top             =   300
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   556
      End
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Filtrar (Nombre PNT)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   885
         TabIndex        =   12
         Top             =   990
         Width           =   2265
      End
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Filtrar (Código PNT)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   870
         TabIndex        =   8
         Top             =   690
         Width           =   2175
      End
      Begin VB.Label lblValor 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   2
         Top             =   315
         Width           =   975
      End
   End
   Begin VB.Label lblPNT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   0
      TabIndex        =   14
      Top             =   9930
      Width           =   10905
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuarios Formadores en PNT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14475
   End
End
Attribute VB_Name = "frmMC_FormadoresPNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs_pnt As ADODB.RecordSet, rs_usuarios As ADODB.RecordSet
Dim quitar_cargando As Boolean

Private objMC As New clsMc_cualificaciones_pnt
Private WithEvents oformCargando As frmCargando
Attribute oformCargando.VB_VarHelpID = -1


Private Sub cabecera_usuarios()

Dim nombre As String


    With glista
        .Rows = 1
        .Cols = 2
        
        
        .TextMatrix(0, 0) = "id_documento"
        .TextMatrix(0, 1) = "Doc. Calidad"
        .ColWidth(0) = 0
        .ColWidth(1) = glista.Width * 0.4
        
        If rs_usuarios.RecordCount <> 0 Then
            rs_usuarios.MoveFirst
            
            While Not rs_usuarios.EOF
                .Cols = .Cols + 2
                nombre = UCase(rs_usuarios("NOMBRE") & " " & rs_usuarios("APELLIDOS"))
                .TextMatrix(0, .Cols - 2) = rs_usuarios("ID_EMPLEADO")
                If Len(nombre) > 28 Then
                    .Col = .Cols - 1
                    .Row = 0
                    .CellAlignment = flexAlignLeftCenter
                    nombre = Left(nombre, 26) & "..."
                Else
                    .ColAlignment(.Cols - 1) = flexAlignCenterCenter
                End If
                .TextMatrix(0, .Cols - 1) = nombre
                .ColWidth(.Cols - 2) = 0
                .ColWidth(.Cols - 1) = 3000
                
                
                rs_usuarios.MoveNext
            Wend
        End If
    
    If .Cols > 2 Then .FixedCols = 2
    
    End With

    

End Sub

Private Sub cargar_datos()
    Dim oPNT As New clsCa_documentos
    
    activar_cargando "Cargando lista de Cualificaciones. Por favor, espere..."
    
    Set rs_usuarios = objMC.Listado_usuarios_cualificados
    
    Set rs_pnt = oPNT.Listado("0", "0", "0", txtFiltroNombre.Text, txtFiltroCod.Text, 0, 0, 0, 0, "0", 0)

End Sub
Private Sub activar_cargando(mensaje)
    quitar_cargando = False
    Set oformCargando = New frmCargando
    
    oformCargando.lblMensaje.Caption = mensaje
    
    oformCargando.Show
End Sub
Private Sub des_marca_usuario(ID_USUARIO As Long, ID_DOCUMENTO As Long, fila As Long, Col As Long)

    Dim estado As Boolean
    
    estado = GridBooleanCell_Estado(glista, fila, Col + 1)
    
    If estado Then
        ' está marcado, lo que significa que lo desmarcará
        
        If objMC.comprobar_usuario_descualificable_en_pnt(ID_USUARIO, ID_DOCUMENTO) Then
            objMC.Eliminar ID_USUARIO, ID_DOCUMENTO
            mostrar_estado_celda fila, Col + 1, 0
        Else
            objMC.desmarcar_usuario ID_USUARIO, ID_DOCUMENTO, USUARIO.getID_EMPLEADO
            mostrar_estado_celda fila, Col + 1, 2
        End If
    Else ' cuando no está marcado, o no tiene texto
        objMC.setDOCUMENTO_ID = ID_DOCUMENTO
        objMC.setUSUARIO_ID = ID_USUARIO
        objMC.setES_FORMADOR = 1
        If Trim(glista.TextMatrix(fila, Col + 1)) = "" Then
            objMC.Insertar USUARIO.getID_EMPLEADO
        Else
            objMC.Modificar USUARIO.getID_EMPLEADO
        End If
        mostrar_estado_celda fila, Col + 1, 1
    End If

End Sub

Private Sub mostrar_estado_celda(fila, columna, tipo_celda)

    
    glista.Row = fila
    glista.Col = columna
    
    If tipo_celda = 0 Then
        glista.TextMatrix(fila, columna) = ""
        glista.CellBackColor = vbWhite
    ElseIf tipo_celda = 1 Then
        GridBooleanCell glista, fila, columna, True
        glista.CellBackColor = vbWhite
    Else
        GridBooleanCell glista, fila, columna, False
        glista.CellBackColor = vbBlue
    End If

End Sub



Private Sub presenta_datos()
    

    Dim rs_usu_pnt As ADODB.RecordSet
    
    Dim iCont As Integer
    
    glista.Rows = 1
        
    If rs_pnt.RecordCount = 0 Then
        quitar_cargando = True
        Exit Sub
    End If
        
        
    Set rs_usu_pnt = objMC.Listado
    
    rs_pnt.MoveFirst
        
    While Not rs_pnt.EOF
        
        iCont = 0
        
        DoEvents
        
        glista.Rows = glista.Rows + 1
        glista.TextMatrix(glista.Rows - 1, 0) = rs_pnt("id_documento")
        glista.TextMatrix(glista.Rows - 1, 1) = "[" & rs_pnt("CODIGO") & "]" & " " & rs_pnt("NOMBRE")
        glista.Row = glista.Rows - 1
        glista.Col = 1
        glista.CellBackColor = vbWhite
        
        ' Filtra el RS de usuarios por si este pnt tuviera formadores
        If rs_usuarios.RecordCount <> 0 Then
            rs_usuarios.MoveFirst
            While Not rs_usuarios.EOF
                iCont = iCont + 2
                
                DoEvents
                
                glista.TextMatrix(glista.Rows - 1, iCont) = rs_usuarios("id_empleado") & "/" & rs_pnt("id_documento")
                                
                ' filtra por si existe
                rs_usu_pnt.Filter = "USUARIO_ID = '" & rs_usuarios("ID_EMPLEADO") & "' AND DOCUMENTO_ID = '" & rs_pnt("id_documento") & "'"
                
                If rs_usu_pnt.RecordCount <> 0 Then
                    mostrar_estado_celda glista.Rows - 1, iCont + 1, IIf(CInt(rs_usu_pnt!ES_FORMADOR) = 1, 1, 2)
                Else
                    mostrar_estado_celda glista.Rows - 1, iCont + 1, 0
                End If
                
                rs_usu_pnt.Filter = ""
                
                        
'                If Not objMC.Carga(rs_usuarios("id_empleado"), rs_pnt("id_documento")) Then
'                    mostrar_estado_celda glista.Rows - 1, iCont + 1, 0
'                Else
'                    mostrar_estado_celda glista.Rows - 1, iCont + 1, IIf(objMC.getES_FORMADOR = 1, 1, 2)
'                End If
                rs_usuarios.MoveNext
            Wend
        End If

        rs_pnt.MoveNext
    Wend


    quitar_cargando = True

End Sub


Private Sub cmdAnadir_Click()

Dim x As Long, id As Long, nombre As String

    id = cmbUsuarios.getPK_SALIDA
    nombre = UCase(cmbUsuarios.getTEXTO)

    If id <= 0 Then Exit Sub


    If MsgBox("Va a añdir al usuario " & nombre & " a la matriz de cualificaciones, ¿Está seguro?", vbQuestion + vbYesNo, "Añadir Usuario a la Matriz de Cualificaciones") = vbNo Then Exit Sub
    
    
    ' Primero Recorre las columnas por si ya ha sido eliminado dicho usuario
    For x = 2 To glista.Cols - 1 Step 2
        If CLng(glista.TextMatrix(0, x)) = id Then
            ' lo ha localizado, por lo tanto, lo unico que hace es agrandar la columna del nombres
            glista.ColWidth(x + 1) = 3000
            Exit Sub
        End If
    Next x
    
    
    'llegado a este punto, no ha encontrado a ese usuario
    
    activar_cargando "Añadiendo nuevo usuario a la lista de Cualificaciones. Por favor, espere..."
    
    With glista
        .Cols = .Cols + 2
        .TextMatrix(0, .Cols - 2) = id
        If Len(nombre) > 28 Then
            .Col = .Cols - 1
            .Row = 0
            .CellAlignment = flexAlignLeftCenter
            nombre = Left(nombre, 26) & "..."
        Else
            .ColAlignment(.Cols - 1) = flexAlignCenterCenter
        End If
        .TextMatrix(0, .Cols - 1) = nombre
        .ColWidth(.Cols - 2) = 0
        .ColWidth(.Cols - 1) = 3000
        
        
    End With
    ' Ahora recorre todas las filas colocando los pares de identificador ID_USUARIO/ID_PNT
    For x = 1 To glista.Rows - 1
        ' en la fila del id
        glista.TextMatrix(x, glista.Cols - 2) = CStr(id) & "/" & glista.TextMatrix(x, 0)
        mostrar_estado_celda x, glista.Cols - 1, 0
    Next x
    
    quitar_cargando = True


End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdEliminar_Click()
Dim x As Long, id As Long, col_localizada As Long


    id = cmbUsuarios.getPK_SALIDA
    
    col_localizada = -1
    
    If id <= 0 Then Exit Sub

    ' primero comprueba que el usuario esté en la lista de cualificados
    
    For x = 2 To glista.Cols - 1 Step 2
        If CLng(glista.TextMatrix(0, x)) = id Then
            ' lo ha localizado, por lo tanto, lo unico que hace es agrandar la columna del nombres
            col_localizada = x
        End If
    Next x
    
    
    If col_localizada = -1 Then
        MsgBox "El usuario " & cmbUsuarios.getTEXTO & " no se encuentra en la Matriz de Cualificaciones", vbInformation, "Eliminar Usuario de la Matriz de Cualificaciones"
        Exit Sub
    End If
    
    ' Ahora comprueba si se puede eliminar o no
    If Not objMC.comprobar_usuario_descualificable(id) Then
        MsgBox "El usuario " & cmbUsuarios.getTEXTO & " no se puede eliminar de la Matriz de Cualificaciones, dado que aparece como FORMADOR en al menos un ensayo.", vbInformation, "Eliminar Usuario de la Matriz de Cualificaciones"
        Exit Sub
    End If


    If MsgBox("ATENCION: Va a eliminar al usuario " & UCase(cmbUsuarios.getTEXTO) & " de la matriz de cualificaciones." & vbCrLf & "SI ELIMINA ESTE USUARIO DE LA LISTA, PERDERÁ TODAS LAS CUALIFICACIONES QUE POSEE PARA DICHO USUARIO" & vbCrLf & "  ¿Está seguro que desea Eliminarlo?", vbCritical + vbYesNo, "Eliminar Usuario de la Matriz de Cualificaciones") = vbNo Then Exit Sub
    
    
    activar_cargando "Eliminando usuario de la Matriz de Cualificaciones. Por favor, espere..."
    
    ' Ahora recorre todas las filas, y para aquellos pnt's en los que esté cualificado, lo elimina
    For x = 1 To glista.Rows - 1
        If GridBooleanCell_Estado(glista, x, col_localizada + 1) Then
            objMC.Eliminar id, CLng(glista.TextMatrix(x, 0))
            glista.TextMatrix(x, col_localizada + 1) = ""
        End If
    Next x
    
    ' pone la columna a 0 de ancho, para así ocultarlo
    glista.ColWidth(col_localizada + 1) = 0
    
    quitar_cargando = True

End Sub

Private Sub cmdFiltrar_Click()
    
    
    cargar_datos
    
    cabecera_usuarios
    
    presenta_datos

    glista_RowColChange
    
End Sub

Private Sub cmdImprimir_Click()
    With frmReport
        .iniciar
        .informe = "/MC/rptMCRelacion_PNT_Formador"
        .CRITERIO = "" ' "{procnc.ID_PROCNC} = " & id_pnc & " and {decodificadora.CODIGO}=110" '"{decodificadora.codigo}=11 and {nc.ID_NC} in [" & Left(strIncidencias, Len(strIncidencias) - 1) & "]"
        .imprimir = False
        .generar
        .Visible = True
    End With
End Sub

Private Sub cmdImprimir2_Click()
    With frmReport
        .iniciar
        .informe = "/MC/rptMCRelacion_Formador_PNT"
        .CRITERIO = "" ' "{procnc.ID_PROCNC} = " & id_pnc & " and {decodificadora.CODIGO}=110" '"{decodificadora.codigo}=11 and {nc.ID_NC} in [" & Left(strIncidencias, Len(strIncidencias) - 1) & "]"
        .imprimir = False
        .generar
        .Visible = True
    End With
End Sub

Private Sub cmdQuitarFiltro_Click()

    txtFiltroCod.Text = ""
    txtFiltroNombre.Text = ""
    
    cmdFiltrar_Click
    
    glista_RowColChange
End Sub


Private Sub Form_Load()
        
    log Me.Name
    
    cargar_botones Me
    
    llenar_combo cmbUsuarios, New clsUsuarios, 0, frmUsuarios, ""
    
    cargar_datos

    cabecera_usuarios

    presenta_datos
    


End Sub


Private Sub Form_Resize()

    glista.Width = Me.ScaleWidth
    glista.Left = 0
    
    cmdcancel.Top = Me.ScaleHeight - 60 - cmdcancel.Height
    cmdcancel.Left = Me.ScaleWidth - 60 - cmdcancel.Width
    
    cmdImprimir.Top = Me.ScaleHeight - 60 - cmdImprimir.Height
    cmdImprimir.Left = cmdcancel.Left - 60 - cmdImprimir.Width
    
    cmdImprimir2.Top = Me.ScaleHeight - 60 - cmdImprimir2.Height
    cmdImprimir2.Left = cmdImprimir.Left - 60 - cmdImprimir2.Width
    
    glista.Height = Me.ScaleHeight - glista.Top - cmdcancel.Height - 120
    
    fraDatos.Width = Me.ScaleWidth
    
    glista.ColWidth(1) = glista.Width * 0.4

    lblPNT.Left = 30
    lblPNT.Top = glista.Top + glista.Height + 30
    lblPNT.Width = cmdcancel.Left - 60

End Sub

Private Sub glista_DblClick()
Dim fila As Long, Col As Long
Dim strCad As String
Dim iusu As Long, idoc As Long

'MsgBox "Fila: " & glista.Row & vbCrLf & "Col: " & glista.Col & vbCrLf & "Texto: " & glista.TextMatrix(glista.Row, glista.Col) & ""

On Error GoTo glista_DblClick_Error
    
fila = glista.Row
Col = glista.Col

' Primero se cerciora que está en la columna correcta.

If Col <= 1 Then Exit Sub

If Col Mod 2 <> 0 Then
    ' cuando está en una columna impar, se va siempre a la anterior
    Col = Col - 1
End If


strCad = glista.TextMatrix(fila, Col)

iusu = CLng(Split(strCad, "/")(0))
idoc = CLng(Split(strCad, "/")(1))


des_marca_usuario iusu, idoc, fila, Col


   
On Error GoTo 0
    Exit Sub
glista_DblClick_Error:
    Exit Sub

End Sub



Private Sub glista_RowColChange()
On Error Resume Next

If Not quitar_cargando Then Exit Sub

If glista.ColWidth(glista.Col) = 0 Or glista.Col <= 1 Then
    lblPNT.Caption = "DOCUMENTO: " & glista.TextMatrix(glista.Row, 1) & vbCrLf & CStr(glista.Rows - 1) & " DOCUMENTOS ENCONTRADOS"
Else
    lblPNT.Caption = "DOCUMENTO: " & glista.TextMatrix(glista.Row, 1) & vbCrLf & "USUARIO: " & glista.TextMatrix(0, glista.Col) & vbCrLf & CStr(glista.Rows - 1) & " DOCUMENTOS ENCONTRADOS"
End If

End Sub

Private Sub oformCargando_pasarela(Cancel As Boolean)
    Cancel = quitar_cargando
End Sub


