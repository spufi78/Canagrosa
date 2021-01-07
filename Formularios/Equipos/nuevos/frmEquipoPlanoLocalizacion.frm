VERSION 5.00
Begin VB.Form frmEquipoPlanoLocalizacion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plano Localización Equipos"
   ClientHeight    =   9825
   ClientLeft      =   1275
   ClientTop       =   885
   ClientWidth     =   15900
   Icon            =   "frmEquipoPlanoLocalizacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   15900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPanel 
      BorderStyle     =   0  'None
      Height          =   9405
      Left            =   30
      TabIndex        =   5
      Top             =   390
      Width           =   14445
      Begin VB.Image imgEquipo 
         Height          =   495
         Left            =   11250
         Stretch         =   -1  'True
         Top             =   60
         Width           =   495
      End
      Begin VB.Image imgTodosEquipos 
         Height          =   495
         Index           =   0
         Left            =   13680
         Stretch         =   -1  'True
         Top             =   90
         Width           =   495
      End
      Begin VB.Image imgPlano 
         Height          =   19335
         Left            =   0
         Picture         =   "frmEquipoPlanoLocalizacion.frx":08CA
         Top             =   0
         Width           =   14415
      End
   End
   Begin VB.VScrollBar navigator 
      Height          =   9405
      LargeChange     =   300
      Left            =   14490
      Max             =   9405
      SmallChange     =   150
      TabIndex        =   4
      Top             =   360
      Width           =   315
   End
   Begin VB.CommandButton cmdGuardarPosiciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar Posiciones"
      Height          =   870
      Left            =   14820
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   14820
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8010
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   14820
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8940
      Width           =   1050
   End
   Begin VB.ComboBox cmbMaquinaria 
      Height          =   315
      ItemData        =   "frmEquipoPlanoLocalizacion.frx":1F74E
      Left            =   30
      List            =   "frmEquipoPlanoLocalizacion.frx":1F750
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   30
      Width           =   14445
   End
   Begin VB.Menu mnumaquina 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuinfo 
         Caption         =   "Ver Info"
      End
   End
End
Attribute VB_Name = "frmEquipoPlanoLocalizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private mvarcolEquipos As clsGenericCollection
Private mvarobjEquipo As clsEquipos

Private lngContadorElementos As Long
Private mvarlngSelMaquina As Long

Private difxPos As Long
Private difyPos As Long
Private mvarblnVerTodo As Boolean
Private mvarblnResultado As Boolean

Private mvarColIntCX() As Long
Private mvarColIntCY() As Long

Private mvarblnEnCarga As Boolean
Private Sub cmdcancel_Click()
    mvarblnResultado = False
    Unload Me
End Sub

Private Sub cmdGuardarPosiciones_Click()

Call guardarPosiciones(True)

'cmdGuardarPosiciones.Enabled = False

End Sub

Private Sub cmdOk_Click()

    If cmdGuardarPosiciones.Enabled Then
'        If MsgBox("Ha modificado la posición en el plano de algun/os Equipo/s. ¿Desea hacer permanentes ahora esas modificaciones de posición?", vbYesNo, "¿Guardar Nuevas Posiciones?") = vbYes Then
            Call guardarPosiciones(True)
'        Else
'            Call guardarPosiciones(False)
'        End If
    End If
    

    mvarblnResultado = True
    Unload Me
    
    

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjEquipo = Nothing

End Sub





Public Property Get EQUIPO() As clsEquipos

    Set EQUIPO = mvarobjEquipo

End Property

Public Property Set EQUIPO(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Private Sub cmbMaquinaria_Click()
If cmbMaquinaria.ItemData(cmbMaquinaria.ListIndex) = 0 Then
    mvarblnVerTodo = True
    Call CargarImagenesTodosEquipos
Else
    mvarblnVerTodo = False
    Call CargarImagenEquipoUnico
End If
End Sub
Private Sub cmbMaquinaria_Change()
    If cmbMaquinaria.ItemData(cmbMaquinaria.ListIndex) = 0 Then
        Call CargarImagenesTodosEquipos
    Else
        Call CargarImagenEquipoUnico
    End If
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    log Me.Name
    cargar_botones Me
    
    Dim strRutaImagen As String

    mvarblnEnCarga = True
    
    If mvarobjEquipo Is Nothing Then
        If PK <> 0 Then
                    
            Dim oE As New clsEquipos
            oE.Carga PK
            Set mvarobjEquipo = oE
            With cmbMaquinaria
                .Clear
                .AddItem mvarobjEquipo.getNOMBRE & "[" & mvarobjEquipo.getSERIE & "]"
                .ItemData(.ListCount - 1) = mvarobjEquipo.getID_EQUIPO
                .AddItem "Todos los Equipos"
                .ItemData(.ListCount - 1) = 0
                .ListIndex = 0
            End With
        
        Else
            With cmbMaquinaria
                .Clear
                .AddItem "Todos los Equipos"
                .ItemData(.ListCount - 1) = 0
                .ListIndex = 0
            End With
            mvarblnVerTodo = True
        End If
    Else
        With cmbMaquinaria
            .Clear
            .AddItem mvarobjEquipo.getNOMBRE & "[" & mvarobjEquipo.getSERIE & "]"
            .ItemData(.ListCount - 1) = mvarobjEquipo.getID_EQUIPO
            .AddItem "Todos los Equipos"
            .ItemData(.ListCount - 1) = 0
            .ListIndex = 0
        End With
    End If
    
    
    mvarblnEnCarga = False
    
    If mvarblnVerTodo = False Then
        Call CargarImagenEquipoUnico
    Else
        Call CargarImagenesTodosEquipos
    End If

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmEquipoPlanoLocalizacion"
           
End Sub



Private Sub CargarImagenEquipoUnico()

    Dim strRutaImagen As String
    
   On Error GoTo CargarImagenEquipoUnico_Error

    strRutaImagen = ReadINI(App.Path + "\config.ini", "Documentos", "familias_equipos_img")
    
    If strRutaImagen <> "" Then
        If mvarobjEquipo.getFAMILIA_ID < 0 Then
            strRutaImagen = strRutaImagen & "\0.jpg"
        Else
            strRutaImagen = strRutaImagen & "\" & mvarobjEquipo.getFAMILIA_ID & ".jpg"
        End If
        imgEquipo.Picture = LoadPicture(strRutaImagen)
    End If
    
    
    ReDim Preserve mvarColIntCX(1)
    ReDim Preserve mvarColIntCY(1)
    mvarColIntCX(0) = mvarobjEquipo.getCoordx
    mvarColIntCY(0) = mvarobjEquipo.getCoordY
    imgEquipo.Left = mvarobjEquipo.getCoordx
    imgEquipo.Top = mvarobjEquipo.getCoordY
    
    imgEquipo.Tag = "id=" & CStr(mvarobjEquipo.getID_EQUIPO)

   On Error GoTo 0
   Exit Sub

CargarImagenEquipoUnico_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarImagenEquipoUnico of Formulario frmEquipoPlanoLocalizacion"

End Sub

Private Sub imgEquipo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo imgEquipo_MouseUp_Error

    Source.Left = Source.Left - (mvarColIntCX(0) - x)
    Source.Top = Source.Top - (mvarColIntCY(0) - y)
    
    Exit Sub
    If x < mvarColIntCX(0) Then
        Source.Left = Source.Left - x
    Else
        Source.Left = Source.Left + x
    End If
    
    If y < mvarColIntCY(0) Then
        Source.Top = Source.Top - y
    Else
        Source.Top = Source.Top + y
    End If
    
    cmdGuardarPosiciones.Enabled = True

   On Error GoTo 0
   Exit Sub

imgEquipo_MouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgEquipo_MouseUp of Formulario frmEquipoPlanoLocalizacion"
End Sub


Private Sub imgPlano_DragDrop(Source As Control, x As Single, y As Single)

If Source.Name = "imgEquipo" Then
     If imgEquipo.Left < imgPlano.Left Then
         imgEquipo.Left = imgPlano.Left
     ElseIf imgEquipo.Left > imgPlano.Left + imgPlano.Width Then
         imgEquipo.Left = imgPlano.Left + imgPlano.Width
     Else
         imgEquipo.Left = imgPlano.Left + x - difxPos
     End If
    
     imgEquipo.Top = imgPlano.Top + y - difyPos
     If imgEquipo.Top < imgPlano.Top Then
         imgEquipo.Top = imgPlano.Top
     ElseIf imgEquipo.Top > imgPlano.Top + imgPlano.Height Then
         imgEquipo.Top = imgPlano.Top + imgPlano.Height
     Else
         imgEquipo.Top = imgPlano.Top + y - difyPos
     End If
     imgEquipo.Drag vbEndDrag
     imgEquipo.ZOrder
     mvarlngSelMaquina = 0
Else ' es imgTodosEquipos
     If imgTodosEquipos(mvarlngSelMaquina).Left < imgPlano.Left Then
         imgTodosEquipos(mvarlngSelMaquina).Left = imgPlano.Left
     ElseIf imgTodosEquipos(mvarlngSelMaquina).Left > imgPlano.Left + imgPlano.Width Then
         imgTodosEquipos(mvarlngSelMaquina).Left = imgPlano.Left + imgPlano.Width
     Else
         imgTodosEquipos(mvarlngSelMaquina).Left = imgPlano.Left + x - difxPos
     End If
    
     imgTodosEquipos(mvarlngSelMaquina).Top = imgPlano.Top + y - difyPos
     If imgTodosEquipos(mvarlngSelMaquina).Top < imgPlano.Top Then
         imgTodosEquipos(mvarlngSelMaquina).Top = imgPlano.Top
     ElseIf imgTodosEquipos(mvarlngSelMaquina).Top > imgPlano.Top + imgPlano.Height Then
         imgTodosEquipos(mvarlngSelMaquina).Top = imgPlano.Top + imgPlano.Height
     Else
         imgTodosEquipos(mvarlngSelMaquina).Top = imgPlano.Top + y - difyPos
     End If
     imgTodosEquipos(mvarlngSelMaquina).Drag vbEndDrag
     imgTodosEquipos(mvarlngSelMaquina).ZOrder
     mvarlngSelMaquina = 0
End If
    
    
End Sub




Private Sub imgEquipo_DblClick()

On Error GoTo imgEquipo_DblClick_Error

    mvarlngSelMaquina = 0
    Call mnuinfo_Click

On Error GoTo 0
    Exit Sub
imgEquipo_DblClick_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgEquipo_DblClick of Formulario frmEquipoPlanoLocalizacion"
    
End Sub

Private Sub imgEquipo_DragDrop(Source As Control, x As Single, y As Single)
    
On Error GoTo imgEquipo_DragDrop_Error

    Source.Left = Source.Left - (mvarColIntCX(0) - x)
    Source.Top = Source.Top - (mvarColIntCY(0) - y)
    
    Exit Sub
    If x < mvarColIntCX(0) Then
        Source.Left = Source.Left - x
    Else
        Source.Left = Source.Left + x
    End If
    
    If y < mvarColIntCY(0) Then
        Source.Top = Source.Top - y
    Else
        Source.Top = Source.Top + y
    End If
    
    cmdGuardarPosiciones.Enabled = True
    

On Error GoTo 0
    Exit Sub
imgEquipo_DragDrop_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgEquipo_DragDrop of Formulario frmEquipoPlanoLocalizacion"

End Sub

Private Sub imgEquipo_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    
On Error GoTo imgEquipo_DragOver_Error

    If State = 0 Then
        mvarColIntCX(0) = x
        mvarColIntCY(0) = y
    End If

On Error GoTo 0
    Exit Sub
imgEquipo_DragOver_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgEquipo_DragOver of Formulario frmEquipoPlanoLocalizacion"
    
End Sub


Private Sub imgEquipo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo imgEquipo_MouseDown_Error
    
    Dim elemento As Long
    If (Button = vbLeftButton) Then
        difxPos = x
        difyPos = y
        elemento = 0
        mvarlngSelMaquina = elemento
        imgEquipo.Drag vbBeginDrag
    ElseIf (Button = vbRightButton) And (Index <> 0) Then
        mvarlngSelMaquina = 0
        PopupMenu mnumaquina
    End If

On Error GoTo 0
Exit Sub
imgEquipo_MouseDown_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Sub imgEquipo_MouseDown del/de la Formulario frmEquipoPlanoLocalizacion"
On Error GoTo 0
End Sub



Private Sub imgTodosEquipos_DblClick(Index As Integer)
On Error GoTo imgTodosEquipos_DblClick_Error

    mvarlngSelMaquina = Index
    Call mnuinfo_Click

On Error GoTo 0
    Exit Sub
imgTodosEquipos_DblClick_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgTodosEquipos_DblClick of Formulario frmEquipoPlanoLocalizacion"
End Sub

Private Sub imgTodosEquipos_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)

On Error GoTo imgTodosEquipos_DragDrop_Error

    Source.Left = Source.Left - (mvarColIntCX(Index) - x)
    Source.Top = Source.Top - (mvarColIntCY(Index) - y)
    
    Exit Sub
    If x < mvarColIntCX(Index) Then
        Source.Left = Source.Left - x
    Else
        Source.Left = Source.Left + x
    End If
    
    If y < mvarColIntCY(Index) Then
        Source.Top = Source.Top - y
    Else
        Source.Top = Source.Top + y
    End If
    
    cmdGuardarPosiciones.Enabled = True

On Error GoTo 0
    Exit Sub
imgTodosEquipos_DragDrop_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgTodosEquipos_DragDrop of Formulario frmEquipoPlanoLocalizacion"
End Sub


Private Sub imgTodosEquipos_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)

On Error GoTo imgTodosEquipos_DragOver_Error
    
    If State = 0 Then
        mvarColIntCX(Index) = x
        mvarColIntCY(Index) = y
    End If
    
    

On Error GoTo 0
    Exit Sub
imgTodosEquipos_DragOver_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgTodosEquipos_DragOver of Formulario frmEquipoPlanoLocalizacion"
End Sub

Private Sub imgTodosEquipos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo imgTodosEquipos_MouseDown_Error

    Dim elemento As Long
    If (Button = vbLeftButton) Then
        difxPos = x
        difyPos = y
        elemento = Index
        mvarlngSelMaquina = elemento
        imgTodosEquipos(elemento).Drag vbBeginDrag
    ElseIf (Button = vbRightButton) And (Index <> 0) Then
        mvarlngSelMaquina = Index
        PopupMenu mnumaquina
    End If
    

On Error GoTo 0
    Exit Sub
imgTodosEquipos_MouseDown_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgTodosEquipos_MouseDown of Formulario frmEquipoPlanoLocalizacion"
End Sub


Private Sub imgTodosEquipos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    imgTodosEquipos(Index).Left = imgTodosEquipos(Index).Left - (mvarColIntCX(Index) - x)
    imgTodosEquipos(Index).Top = imgTodosEquipos(Index).Top - (mvarColIntCY(Index) - y)
    
    Exit Sub
    If x < mvarColIntCX(Index) Then
        Source.Left = Source.Left - x
    Else
        Source.Left = Source.Left + x
    End If
    
    If y < mvarColIntCY(Index) Then
        Source.Top = Source.Top - y
    Else
        Source.Top = Source.Top + y
    End If
    
    cmdGuardarPosiciones.Enabled = True
End Sub




Private Sub mnuinfo_Click()
    Dim objfrm As New frmEquipoPlanoInfo
    Dim strCad As String

    If mvarlngSelMaquina = 0 Then
        Set objfrm.EQUIPO = mvarobjEquipo
    Else
        strCad = imgTodosEquipos(mvarlngSelMaquina).Tag
        strCad = Split(strCad, ";")(0)
        strCad = Split(strCad, "=")(1)
        
        Set objfrm.EQUIPO = mvarcolEquipos.Item(strCad)
    End If
    objfrm.Show vbModal
    Unload objfrm
    Set objfrm = Nothing
End Sub



Public Property Get VerTodo() As Boolean

On Error GoTo VerTodo_Error

    VerTodo = mvarblnVerTodo

On Error GoTo 0
    Exit Property
VerTodo_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure VerTodo of Formulario frmEquipoPlanoLocalizacion"

End Property

Public Property Let VerTodo(ByVal blnVerTodo As Boolean)

On Error GoTo VerTodo_Error

    mvarblnVerTodo = blnVerTodo

On Error GoTo 0
    Exit Property
VerTodo_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure VerTodo of Formulario frmEquipoPlanoLocalizacion"

End Property

Private Sub CargarImagenesTodosEquipos()
On Error GoTo CargarImagenesTodosEquipos_Error
Dim objEq As New clsEquipos
Dim objEqUnico As clsEquipos
Dim lngCont As Long
Dim strRutaImagen As String
    
    If mvarblnEnCarga Then Exit Sub
    
    If imgTodosEquipos.Count > 1 Then Exit Sub
    
    Set objEq = New clsEquipos
    lngCont = 0
    Set mvarcolEquipos = objEq.ListadoObjetos()
    
    If Not mvarobjEquipo Is Nothing Then
        Set objEqUnico = mvarobjEquipo
    Else
        Set objEqUnico = New clsEquipos
    End If
    
    For Each objEq In mvarcolEquipos.Iterator
        If objEq.getID_EQUIPO <> objEqUnico.getID_EQUIPO Then
            If objEq.getID_EQUIPO = 1515 Then
                'MsgBox ""
            End If
            lngCont = lngCont + 1
            
            If objEq.getCoordY = 0 Then objEq.setCoordY = imgPlano.Top
            If objEq.getCoordx = 0 Then objEq.setCoordX = imgPlano.Left
            
            
            ReDim Preserve mvarColIntCX(lngCont)
            ReDim Preserve mvarColIntCY(lngCont)
            
            mvarColIntCX(lngCont) = objEq.getCoordx
            mvarColIntCY(lngCont) = objEq.getCoordY
            
            Load imgTodosEquipos(lngCont)
            
            imgTodosEquipos(lngCont).Left = objEq.getCoordx
            imgTodosEquipos(lngCont).Top = objEq.getCoordY
            imgTodosEquipos(lngCont).Tag = "id=" & CStr(objEq.getID_EQUIPO)
            imgTodosEquipos(lngCont).ToolTipText = objEq.getNOMBRE
            
            strRutaImagen = ReadINI(App.Path + "\config.ini", "Documentos", "familias_equipos_img")
            If strRutaImagen <> "" Then
                strRutaImagen = strRutaImagen & "\" & objEq.getFAMILIA_ID & ".jpg"
                imgTodosEquipos(lngCont).Picture = LoadPicture(strRutaImagen)
            End If
            
            imgTodosEquipos(lngCont).Visible = True
            imgTodosEquipos(lngCont).ZOrder
            
        End If
    Next objEq
        
    Set objEq = Nothing
    
On Error GoTo 0
    Exit Sub
CargarImagenesTodosEquipos_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarImagenesTodosEquipos of Formulario frmEquipoPlanoLocalizacion"

End Sub

Private Sub guardarPosiciones(ByVal guardarEnBd As Boolean)
On Error GoTo guardarPosiciones_Error

Dim lngCont As Long, lngDif As Long
Dim strId As String, arrTag() As String
Dim objEq As clsEquipos

lngDif = -1 * imgPlano.Top


'If MsgBox("Al guardar las posiciones de los equipos, no podrá restaurar las posiciones anteriores", vbInformation + vbYesNo, "¿Está seguro de continuar?") = vbNo Then
'    Exit Sub
'End If


' Guarda la posicion de la máquina única
If Not mvarobjEquipo Is Nothing Then
    mvarobjEquipo.setCoordX = imgEquipo.Left
    mvarobjEquipo.setCoordY = imgEquipo.Top + lngDif

    If guardarEnBd Then Call mvarobjEquipo.ModificarPosicionEnPlano
    
End If

' Ahora guarda la posicion en la colecion de demás objetos
If imgTodosEquipos Is Nothing Then
    cmdGuardarPosiciones.Enabled = False
    Exit Sub
End If

If imgTodosEquipos.Count <= 1 Then
    'cmdGuardarPosiciones.Enabled = False
    Exit Sub
End If

For lngCont = 1 To imgTodosEquipos.Count - 1
    strId = imgTodosEquipos(lngCont).Tag
    If strId = "id=1516" Then
        MsgBox ""
    End If
    arrTag = Split(strId, ";")
    'If UBound(arrTag) >= 1 Then
        strId = Split(arrTag(0), "=")(1)
        Set objEq = mvarcolEquipos.Item(strId)
        objEq.setCoordX = mvarColIntCX(lngCont)
        objEq.setCoordY = mvarColIntCY(lngCont) + lngDif
        
        If guardarEnBd Then Call objEq.ModificarPosicionEnPlano
        Call mvarcolEquipos.Replace(CStr(objEq.getID_EQUIPO), objEq)
    'End If
Next
    

On Error GoTo 0
    Exit Sub
guardarPosiciones_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure guardarPosiciones of Formulario frmEquipoPlanoLocalizacion"
End Sub

Public Property Get RESULTADO() As Boolean

    RESULTADO = mvarblnResultado

End Property

Public Property Let RESULTADO(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub navigator_Change()

Dim lngDif As Long, lngCont As Long

    lngDif = navigator.value + imgPlano.Top
    imgPlano.Top = (-1) * navigator.value
    
    If Not mvarblnVerTodo Then
        imgEquipo.Top = imgEquipo.Top - lngDif
    Else
        For lngCont = 1 To imgTodosEquipos.Count
            imgTodosEquipos(lngCont).Top = imgTodosEquipos(lngCont).Top - lngDif
        Next lngCont
    End If
    
    
End Sub


