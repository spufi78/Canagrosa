Attribute VB_Name = "Funciones"
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetTempPath Lib "KERNEL32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal Hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetProfileString Lib "KERNEL32" Alias "GetProfileStringA" ( _
    ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long) As Long
Public Function lista_colorear(lista As ListView, fila As Integer, color As Long) As Boolean
    Dim i As Integer
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
End Function
Public Function lista_negrita(lista As ListView, fila As Integer) As Boolean
    Dim i As Integer
    lista.ListItems(fila).bold = True
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).bold = True
    Next
End Function

Public Function recuperaIVA() As Integer
    Dim oParametros As New clsParametros
    If oParametros.Carga(parametros.IVA, "") Then
        recuperaIVA = CInt(oParametros.getVALOR)
    Else
        recuperaIVA = 0
    End If
    Set oParametros = Nothing
End Function
Public Function recuperaANNO() As Integer
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("SELECT YEAR(CURRENT_DATE)")
    If rs.RecordCount > 0 Then
        recuperaANNO = rs(0)
    Else
        recuperaANNO = Year(Date)
    End If
End Function
Public Function formatear(s As String, ENTEROS As Integer, DECIMALES As Integer) As String
    Dim aux As String
    Dim i As Integer
    For i = 1 To ENTEROS - 1
        aux = aux & "#"
    Next
    If DECIMALES > 0 Then
        aux = aux & "0."
    Else
        aux = aux & "0"
    End If
    For i = 1 To DECIMALES
        aux = aux & "0"
    Next
    formatear = Trim(Format(s, aux))
End Function

Public Function fecha_bd(ByVal fecha As String) As String
If Trim(fecha) = "" Then fecha = "0000-00-00"
fecha = Format(fecha, "yyyy-mm-dd")

fecha_bd = fecha

End Function

Sub Espera(Segundos As Single)
  Dim ComienzoSeg As Single
  Dim FinSeg As Single
  ComienzoSeg = Timer
  FinSeg = ComienzoSeg + Segundos
  Do While FinSeg > Timer
      DoEvents
      If ComienzoSeg > Timer Then
          FinSeg = FinSeg - 24 * 60 * 60
      End If
  Loop
End Sub
Function MD5(cadena As String) As String
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("SELECT MD5(" & cadena & ")", True)
    If rs.RecordCount > 0 Then
        MD5 = rs(0)
    Else
        MD5 = ""
    End If
End Function
Function Encripta(Strg$, PassWord$)
   Dim b$, s$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, p$
   j = 1
   For i = 1 To Len(PassWord$)
     p$ = p$ & Asc(Mid$(PassWord$, i, 1))
   Next
    
   For i = 1 To Len(Strg$)
     A1 = Asc(Mid$(p$, j, 1))
     j = j + 1: If j > Len(p$) Then j = 1
     A2 = Asc(Mid$(Strg$, i, 1))
     A3 = A1 Xor A2
     b$ = Hex$(A3)
     If Len(b$) < 2 Then b$ = "0" + b$
     s$ = s$ + b$
   Next
   Encripta = s$
End Function
Function Desencripta(Strg$, PassWord$)
   Dim b$, s$, i As Long, j As Long
   Dim A1 As Long, A2 As Long, A3 As Long, p$
   j = 1
   For i = 1 To Len(PassWord$)
     p$ = p$ & Asc(Mid$(PassWord$, i, 1))
   Next
   
   For i = 1 To Len(Strg$) Step 2
     A1 = Asc(Mid$(p$, j, 1))
     j = j + 1: If j > Len(p$) Then j = 1
     b$ = Mid$(Strg$, i, 2)
     A3 = val("&H" + b$)
     A2 = A1 Xor A3
     s$ = s$ + Chr$(A2)
   Next
   Desencripta = s$
End Function
Public Sub cargar_combo(combo As DataCombo, oB As Object)
    Dim rs As ADODB.Recordset
    Set rs = oB.Listado_Combo
    Set combo.RowSource = rs
    combo.ListField = rs(1).Name
    combo.BoundColumn = rs(0).Name
    Set rs = Nothing
End Sub

Public Sub Cargar_ComboBox(combo As ComboBox, oB As Object)
    
    Dim rs As ADODB.Recordset
    Set rs = oB.Listado_Combo
        
    combo.Clear
    
    rs.MoveFirst
    While Not rs.EOF
        combo.AddItem rs(1).Value
        combo.ItemData(combo.ListCount - 1) = rs(0).Value
        rs.MoveNext
    Wend
    
End Sub


Public Function localizar_directorio_escaneo_equipo()

    Dim uso As String
    Dim oParam As New clsParametros

    ' Por defecto, lo que dice en el config
    localizar_directorio_escaneo_equipo = ReadINI(App.Path & "\config.ini", "Documentos", "Escaner")
    
    uso = USUARIO.getUSO
    
    Set oParam = New clsParametros
    
    oParam.Carga parametros.ESCANER_USO_LOCAL_RUTA_CARPETA, uso
    
    If Trim(oParam.getVALOR) <> "" Then
        localizar_directorio_escaneo_equipo = Replace(oParam.getVALOR, "/", "\")
    End If
    
    Set oParam = Nothing
    
    
End Function

Public Sub log(datos As String)
    On Error Resume Next
    If USUARIO.getUSUARIO <> "" Then
        MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date)
        MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date) & "\" & Format(Date, "mmmm")
        On Error GoTo fallo
        Open ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date) & "\" & Format(Date, "mmmm") & "\" & Format(Date, "yyyy-mm-dd") & " " & UCase(USUARIO.getUSUARIO) & ".txt" For Append As #1
        If Left(datos, 3) = "frm" Then
            Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & USUARIO.getUSO & ";" & String(75, "-")
            Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & USUARIO.getUSO & ";" & datos
            Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & USUARIO.getUSO & ";" & String(75, "-")
        Else
            If Left(datos, 13) = "Desc.Error : " Then
                Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & USUARIO.getUSO & ";" & String(80, "*")
                Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & USUARIO.getUSO & ";" & datos
                Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & USUARIO.getUSO & ";" & String(80, "*")
            Else
                Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & USUARIO.getUSO & ";" & datos
            End If
        End If
    End If
    Close #1
    Exit Sub
fallo:
    Close
    Exit Sub
End Sub
Public Sub error_grave(datos As String)
    Dim cadena As String
    cadena = "Desc.Error : " & datos
    cadena = cadena & vbNewLine & "Usuario : " & USUARIO.getUSUARIO
    cadena = cadena & vbNewLine & "Puesto : " & USUARIO.getUSO
    cadena = cadena & vbNewLine & "Fecha/Hora : " & Now
    Call Enviar_Mail_CDO(BUZON_CORREO_LOG, "[ERROR-GESLAB]", cadena, vbNullString)
    log (cadena)
    MsgBox cadena, vbCritical, App.Title
End Sub
Public Sub error_grave_jgm(datos As String)
    Dim cadena As String
    cadena = "Desc.Error : " & datos
    cadena = cadena & vbNewLine & "Usuario : " & USUARIO.getUSUARIO
    cadena = cadena & vbNewLine & "Puesto : " & USUARIO.getUSO
    cadena = cadena & vbNewLine & "Fecha/Hora : " & Now
    ' Generar captura de pantalla para envio por correo
    Dim Capture As CaptureWindow
    Dim Conversor As Class1
    Set Capture = New CaptureWindow
    Set Conversor = New Class1
    Conversor.GrabarJpg Capture.CapturarPantalla(), App.Path & "\captura.jpg", CByte(70)
    
    Set captura = Nothing
    Set Conversor = Nothing
    Call Enviar_Mail_CDO(BUZON_CORREO_LOG, "[ERROR-GESLAB]", cadena, App.Path & "\captura.jpg")
    log (cadena)
    MsgBox cadena, vbCritical, App.Title
End Sub
Public Function pc_es_tablet() As Boolean
    
    Dim tablets As String, arrTablets() As String, i As Long
    
    ' carga la lista de ordenadores que son tablets
    tablets = ReadINI(App.Path & "\config.ini", "Otros", "tablets")
    arrTablets = Split(tablets, ";")
    
    pc_es_tablet = False
        
    For i = 0 To (UBound(arrTablets) - 1)
        If LCase(NOMBRE_PC) = LCase(arrTablets(i)) Then
            ' si este pc está en la lista, es un tablet
            pc_es_tablet = True
            Exit Function
        End If
    Next i
    
        
End Function

Public Function subindices(texto As String) As Boolean
    subindices = False
    If InStr(1, UCase(texto), "SUP(", vbTextCompare) > 0 Or _
       InStr(1, UCase(texto), "SUB(", vbTextCompare) > 0 Then
        subindices = True
    End If
End Function

Public Function subindice_formateado(texto As String) As String
    If subindices(texto) = True Then
        Dim pos As Integer
        Dim pos2 As Integer
        Dim numero1 As Integer
        Dim numero2 As Integer
        Dim exponente As Integer
        Dim i As Integer
        Dim total As Long
        ' Primer numero
        pos = InStr(1, UCase(texto), "X")
        If pos > 0 Then
            numero1 = Left(texto, pos - 1)
        End If
        ' Segundo número
        pos2 = InStr(1, UCase(texto), "SUP(")
        If pos < 1 Then
         pos = 0
        End If
        numero2 = Mid(texto, pos + 1, pos2 - (pos + 1))
        ' Calculamos el exponente
        pos = InStr(1, UCase(texto), "SUP")
        pos2 = InStr(pos, UCase(texto), ")")
        exponente = Mid(texto, pos + 4, pos2 - (pos + 4))
        ' Calculamos el total
        total = 1
        For i = 1 To exponente
         total = total * numero2
        Next
        subindice_formateado = CStr(numero1 * total)
    End If
End Function
Public Sub cargar_botones(ByVal frm As Form)
    Dim ctrl As Control
'    On Error Resume Next
   On Error GoTo cargar_botones_Error

    For Each ctrl In frm.Controls
        Select Case TypeName(ctrl)
           Case "CheckBox", "OptionButton", "CommandButton", "TextBox", "ComboBox"
               Select Case UCase(ctrl.Name)
                 Case "CMDACEPTAR", "CMDOK"
                    Set ctrl.Picture = frmMenu.botones.ListImages(1).Picture
                 Case "CMDANADIR", "CMDANADIRREACTIVO", "CMDANADIREQUIPO"
                    Set ctrl.Picture = frmMenu.botones.ListImages(4).Picture
'                 Case "CMDANULAR"
'                    ctrl.Text = "Anular"
'                    Set ctrl.Picture = frmMenu.botones.ListImages(29).Picture
                 Case "CMDMODIFICAR", "CMDMODIFICAREQUIPO"
                    Set ctrl.Picture = frmMenu.botones.ListImages(6).Picture
                 Case "CMDELIMINAR", "CMDANULAR", "CMDELIMINAREQUIPO", "CMDELIMINARREACTIVO", "CMDDESVINCULAR", "CMDBORRAR"
                    Set ctrl.Picture = frmMenu.botones.ListImages(5).Picture
                 Case "CMDIMPRIMIR", "CMDLISTADO"
                    Set ctrl.Picture = frmMenu.botones.ListImages(7).Picture
                 Case "CMDCANCEL", "CMDSALIR"
                    Set ctrl.Picture = frmMenu.botones.ListImages(3).Picture
                 Case "CMDBUSCAR"
'                    ctrl.Text = "Buscar"
                    Set ctrl.Picture = frmMenu.botones.ListImages(8).Picture
                 Case "CMDDUPLICAR"
'                    ctrl.Text = "Duplicar"
                    Set ctrl.Picture = frmMenu.botones.ListImages(10).Picture
                 Case "CMDQUIEN"
'                    ctrl.Text = "¿Donde?"
                    Set ctrl.Picture = frmMenu.botones.ListImages(11).Picture
                  Case "CMDMAIL", "CMDINVITACION", "CMDCORREO", "CMDMOON", "CMDCERTIFICADO"
                    Set ctrl.Picture = frmMenu.botones.ListImages(20).Picture
                 Case "CMDFAMILIA"
'                    ctrl.Text = "¿Donde?"
                    Set ctrl.Picture = frmMenu.botones.ListImages(9).Picture
                 Case "CMDETIQUETA", "CMDETIQUETAS", "CMDETIQUETASOLUCIONES"
'                    ctrl.Text = "¿Donde?"
                    Set ctrl.Picture = frmMenu.botones.ListImages(13).Picture
                 Case "CMDLIMPIAR", "CMDRECARGA", "CMDINICIAR"
                    Set ctrl.Picture = frmMenu.botones.ListImages(14).Picture
                 Case "CMDMINIMIZAR"
                    Set ctrl.Picture = frmMenu.botones.ListImages(15).Picture
                 Case "CMDDETER", "CMDVERPNC", "CMDVERSELLANTE"
                    Set ctrl.Picture = frmMenu.botones.ListImages(16).Picture
                 Case "CMDDETERMINACIONES", "CMDPARAMETROS", "CMDSUBTIPOS"
                    Set ctrl.Picture = frmMenu.botones.ListImages(31).Picture
                 Case "CMDINFORME", "CMDFIRMAS", "CMDREVISIONES"
                    Set ctrl.Picture = frmMenu.botones.ListImages(17).Picture
                 Case "CMDVIDA", "CMDCOMIENZO"
                    Set ctrl.Picture = frmMenu.botones.ListImages(18).Picture
                 Case "CMDINFREGISTRO", "CMDPROFORMA"
                    Set ctrl.Picture = frmMenu.botones.ListImages(19).Picture
                 Case "CMDADJUNTOS", "CMDADJUNTAR", "CMDFACTURACION", "CMDGENERA"
                    Set ctrl.Picture = frmMenu.botones.ListImages(20).Picture
                 Case "CMDVERMUESTRA"
                    Set ctrl.Picture = frmMenu.botones.ListImages(21).Picture
                 Case "CMDVEREXCEL", "CMDEXCEL", "CMDEXCELFACTURAS"
                    Set ctrl.Picture = frmMenu.botones.ListImages(22).Picture
                 Case "CMDMOSTRAR", "CMDPREVISUALIZAR", "CMDCREARDOCUMENTOS", "CMDVERACCION"
                    Set ctrl.Picture = frmMenu.botones.ListImages(23).Picture
                 Case "CMDESCANER"
                    Set ctrl.Picture = frmMenu.botones.ListImages(24).Picture
                 Case "CMDUSB"
                    Set ctrl.Picture = frmMenu.botones.ListImages(25).Picture
                 Case "CMDGRAFICO", "CMDCURVAS"
                    Set ctrl.Picture = frmMenu.botones.ListImages(26).Picture
                 Case "CMDCONSOLIDAR", "CMDHISTORIALCAMBIOS"
                    Set ctrl.Picture = frmMenu.botones.ListImages(27).Picture
                 Case "CMDIMAGEN"
                    Set ctrl.Picture = frmMenu.botones.ListImages(28).Picture
                 Case "CMDOBSERVADOR", "CMDCLIENTES"
                    Set ctrl.Picture = frmMenu.botones.ListImages(30).Picture
                 Case "CMDTIPOENSAYO"
                    Set ctrl.Picture = frmMenu.botones.ListImages(31).Picture
                 Case "CMDPNT", "CMDCUALIFICAR", "CMDCERTIFICADO", "CMDFIRMACLIENTE"
                    Set ctrl.Picture = frmMenu.botones.ListImages(32).Picture
                 Case "CMDVERIFICACION"
                    Set ctrl.Picture = frmMenu.botones.ListImages(33).Picture
                 Case "CMDCALCULAR"
                    Set ctrl.Picture = frmMenu.botones.ListImages(34).Picture
                 Case "CMDOFERTAS"
                    Set ctrl.Picture = frmMenu.botones.ListImages(35).Picture
                 Case "CMDFINALIZAR", "CMDCERTIFICAR", "CMDCERTIFICAR2"
                    Set ctrl.Picture = frmMenu.botones.ListImages(32).Picture
                 Case "CMDPARAR"
                    Set ctrl.Picture = frmMenu.botones.ListImages(2).Picture
                 Case "CMDDOCCURSO"
                    Set ctrl.Picture = frmMenu.botones.ListImages(19).Picture
                 Case "CMDTRACCION"
                    Set ctrl.Picture = frmMenu.botones.ListImages(36).Picture
                 Case "CMDALBARAN"
                    Set ctrl.Picture = frmMenu.botones.ListImages(37).Picture
                 Case "CMDFACTURA"
                    Set ctrl.Picture = frmMenu.botones.ListImages(38).Picture
                 Case "CMDMATERIAL"
                    Set ctrl.Picture = frmMenu.botones.ListImages(39).Picture
                 Case "CMDDIMENSION"
                    Set ctrl.Picture = frmMenu.botones.ListImages(40).Picture
               End Select
        End Select
    Next

   On Error GoTo 0
   Exit Sub

cargar_botones_Error:
    MsgBox "Control : " & TypeName(ctrl) & " BOTON: " & UCase(ctrl.Name)
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_botones of Módulo Funciones"
End Sub
Public Function generar_clave_web() As String
    Randomize
    generar_clave_web = str(Int((100000 - 999999 + 1) * Rnd + 999999))
End Function
Function monedaNum(VALOR As String) As Single
    Dim m As String
    m = VALOR
    m = Replace(m, "€", "")
    m = Replace(m, ".", "")
'    m = Replace(m, ",", ".")
    monedaNum = CSng(m)
End Function
Function moneda(VALOR As String) As String
'    If UCase(ReadINI(App.Path + "\config.ini", "server", "tipo")) = "ACCESS" Then
'        moneda = Format(VALOR, "currency")
'    Else
        moneda = Format(Replace(VALOR, ".", ","), "currency")
'    End If
End Function
Function moneda4(VALOR As String) As String
'    If UCase(ReadINI(App.Path + "\config.ini", "server", "tipo")) = "ACCESS" Then
'        moneda = Format(VALOR, "currency")
'    Else
        moneda4 = Format(Replace(VALOR, ".", ","), "0.0000")
'    End If
End Function
Function moneda_bd(VALOR As String) As String
    VALOR = IIf(Trim(VALOR) = "", "0", VALOR)
    VALOR = Replace(VALOR, "€", "")
    If UCase(ReadINI(App.Path + "\config.ini", "server", "tipo")) = "ACCESS" Then
        moneda_bd = Format(VALOR, "currency")
    Else
        moneda_bd = Replace(Format(VALOR, "0.00"), ",", ".")
    End If
End Function
Function numerico_bd(VALOR As String) As String
    If Trim(VALOR) = "" Then
        numerico_bd = "null"
    Else
        numerico_bd = Replace(Format(VALOR, "0.00"), ",", ".")
    End If
End Function
Function numerico(VALOR As String, DECIMALES As Integer) As String
    If Trim(VALOR) = "" Then
        numerico = ""
    Else
        If DECIMALES = 0 Then
            numerico = Format(VALOR, "0")
        Else
            Dim d As String
            For i = 1 To DECIMALES
                d = d + "0"
            Next
            numerico = Format(VALOR, "0." + d)
        End If
    End If
End Function

Function numericoBD(VALOR As String) As String
    If Trim(VALOR) = "" Then
        numericoBD = "null"
    Else
        Dim s As String
        s = Replace(VALOR, ".", "")
        s = Replace(VALOR, ",", ".")
        numericoBD = s
    End If
End Function
Function moneda_bd4(VALOR As String) As String
    VALOR = IIf(Trim(VALOR) = "", "0", VALOR)
    VALOR = Replace(VALOR, "€", "")
    If UCase(ReadINI(App.Path + "\config.ini", "server", "tipo")) = "ACCESS" Then
        moneda_bd4 = Format(VALOR, "currency")
    Else
        moneda_bd4 = Replace(Format(VALOR, "0.0000"), ",", ".")
    End If
End Function
Function texto(VALOR As Object) As String
    If IsNull(VALOR) Then
        texto = ""
    Else
        texto = CStr(Trim(VALOR))
    End If
End Function
Function entero(VALOR As Object) As Integer
    If IsNull(VALOR) Then
        entero = 0
    Else
        entero = CInt(VALOR)
    End If
End Function
Function textoBD(VALOR As String) As String
    textoBD = CStr(Trim(VALOR))
End Function
Function textoList(VALOR As Object) As String
    If IsNull(VALOR) Then
        textoList = " "
    Else
        If VALOR = "" Then
            textoList = " "
        Else
            textoList = CStr(Trim(VALOR))
        End If
    End If
End Function


Public Function Establecer_Impresora(ByVal NamePrinter As String) As Boolean
On Error GoTo errSub
       
    'Variable de referencia
    Dim obj_Impresora As Object
       
    'Creamos la referencia
    Set obj_Impresora = CreateObject("WScript.Network")
        obj_Impresora.setdefaultprinter NamePrinter
    Set obj_Impresora = Nothing
           
        'La función devuelve true y se cambió con éxito
        Establecer_Impresora = True
'        MsgBox "La impresora se cambió correctamente", vbInformation
    Exit Function
       
       
'Error al cambiar la impresora
errSub:
If Err.Number = 0 Then Exit Function
   Establecer_Impresora = False
   MsgBox "error: " & Err.Number & Chr(13) & "Description: " & Err.Description
   On Error GoTo 0
End Function

Public Function Impresora_Predeterminada() As String
  
    Dim Buffer As String
    Dim ret As Integer
  
    Buffer = Space(255)
  
    ret = GetProfileString("Windows", ByVal "device", "", _
                                 Buffer, Len(Buffer))
  
    If ret Then
        Impresora_Predeterminada = UCase(Left(Buffer, _
                                   InStr(Buffer, ",") - 1))
    Else
        MsgBox "Error al recuperar la impresora predeterminada del sistema.", vbCritical, App.Title
    End If
    
End Function
'M1051-I
'Public Sub GridBooleanCell(ByRef grid As MSFlexGrid, ByVal fila As Long, ByVal columna As Long, ByVal VALOR As Boolean)
'    With grid
'        .Row = fila
'        .Col = columna
'        .CellFontName = "Wingdings"
'        .CellFontSize = 14
'        .CellAlignment = flexAlignCenterCenter
'        ' edita la celda
'
'        If VALOR Then
'            '.TextMatrix(.Rows - 1, 1) = Chr(254) ' true
'            .Text = Chr(254) ' true
'        Else
'            '.TextMatrix(.Rows - 1, 1) = Chr(168) ' false
'            .Text = Chr(168) ' false
'        End If
'    End With
'End Sub
'
'Public Sub GridBooleanCell_cambiar_a_no_booleano(ByRef grid As MSFlexGrid, ByVal fila As Long, ByVal columna As Long, ByVal fila_formato As Long, ByVal columna_formato As Long)
'Dim formato As String, font_size As Single, alineacion As Integer
'
'    With grid
'        .Row = fila_formato
'        .Col = columna_formato
'        formato = .CellFontName
'        font_size = .CellFontSize
'        alineacion = .CellAlignment
'
'        .Row = fila
'        .Col = columna
'        .CellFontName = formato
'        .CellFontSize = font_size
'        .CellAlignment = alineacion
'    End With
'End Sub
'M1051-F

Public Function KeyAscii_SoloNumerico(ByRef objTxt As TextBox, ByVal KeyAscii As Integer, Optional ByVal Negativos As Boolean = False) As Integer

Select Case KeyAscii
    Case 8
        KeyAscii_SoloNumerico = KeyAscii
    Case 45
        If Negativos Then
            If Len(Trim(strCad)) = 0 Then
                KeyAscii_SoloNumerico = KeyAscii
            End If
        End If

    Case Asc("0") To Asc("9")
        KeyAscii_SoloNumerico = KeyAscii
    Case Else
        KeyAscii_SoloNumerico = 0
End Select

End Function



Public Function KeyAscii_SoloDecimal(ByRef objTxt As TextBox, ByVal KeyAscii As Integer, Optional ByVal Negativos As Boolean = False) As Integer

Dim strCad As String
strCad = objTxt.Text

Select Case KeyAscii
    Case 8
        KeyAscii_SoloDecimal = KeyAscii
    Case 45
        If Negativos Then
            If Len(Trim(strCad)) = 0 Then
                KeyAscii_SoloDecimal = KeyAscii
            End If
        End If
    Case Asc("0") To Asc("9")
        KeyAscii_SoloDecimal = KeyAscii
    Case Asc("."), Asc(",")
        If InStr(1, strCad, ",") <= 0 Then
            ' No Existe el separador Decimal
            'Ahora comprueba que no esté en la primera posicion
            If Len(Trim(strCad)) = 0 Then
                KeyAscii_SoloDecimal = 0 ' Esta en la primera posicion
            Else
                KeyAscii_SoloDecimal = Asc(",") ' Siempre usa la coma como separador decimal
            End If
        End If
    Case Else
        KeyAscii_SoloDecimal = 0
End Select

End Function

Public Function KeyAscii_SoloDecimal_tbgrid(ByRef objTxt As String, ByVal KeyAscii As Integer, Optional ByVal Negativos As Boolean = False) As Integer

Dim strCad As String

strCad = objTxt

Select Case KeyAscii
    Case 8
        KeyAscii_SoloDecimal_tbgrid = KeyAscii
    Case 45
        If Negativos Then
            If Len(Trim(strCad)) = 0 Then
                KeyAscii_SoloDecimal_tbgrid = KeyAscii
            End If
        End If
    Case Asc("0") To Asc("9")
        KeyAscii_SoloDecimal_tbgrid = KeyAscii
    Case Asc("."), Asc(",")
        If InStr(1, strCad, ",") <= 0 Then
            ' No Existe el separador Decimal
            'Ahora comprueba que no esté en la primera posicion
            If Len(Trim(strCad)) = 0 Then
                KeyAscii_SoloDecimal_tbgrid = 0 ' Esta en la primera posicion
            Else
                KeyAscii_SoloDecimal_tbgrid = Asc(",") ' Siempre usa la coma como separador decimal
            End If
        End If
    Case vbKeyEscape
        KeyAscii_SoloDecimal_tbgrid = KeyAscii
    Case Else
        KeyAscii_SoloDecimal_tbgrid = 0
End Select

End Function
'M1051-I
'Public Function GridBooleanCell_Estado(ByRef grid As MSFlexGrid, ByVal fila As Long, ByVal columna As Long) As Boolean
'    GridBooleanCell_Estado = False
'    With grid
'        .Row = fila
'        .Col = columna
'        .CellFontName = "Wingdings"
'        .CellFontSize = 14
'        .CellAlignment = flexAlignCenterCenter
'        ' comprueba el valor de la celda
'
'        If Trim(.TextMatrix(fila, columna)) = "" Then
'            GridBooleanCell_Estado = False
'        ElseIf Asc(.TextMatrix(fila, columna)) = 254 Then ' false
'            GridBooleanCell_Estado = True
'        Else
'            GridBooleanCell_Estado = False
'        End If
'    End With
'End Function
'
'Public Sub EliminarFilaGrid(ByRef grid As MSFlexGrid, ByVal fila As Integer)
'Dim intRow As Integer, intCol As Integer
'
'With grid
'    If fila <= 0 Then Exit Sub
'
'    For intRow = fila To .Rows - 2
'        For intCol = 0 To .COLS - 1
'            .TextMatrix(intRow, intCol) = .TextMatrix(intRow + 1, intCol)
'        Next intCol
'    Next intRow
'
'    .Rows = .Rows - 1
'End With
'End Sub
'M1051-F


Public Function ClrStr(Optional ByVal cadena As String = "", Optional ByVal ConvertirAMinuscula As Boolean = True, Optional ByVal QuitarAcentos As Boolean = False, Optional ByVal QuitarParentesis As Boolean) As String

If ConvertirAMinuscula Then cadena = LCase(cadena)
    
If QuitarAcentos Then
    cadena = Replace(cadena, "á", "a")
    cadena = Replace(cadena, "Á", "A")
    cadena = Replace(cadena, "é", "e")
    cadena = Replace(cadena, "í", "i")
    cadena = Replace(cadena, "ó", "o")
    cadena = Replace(cadena, "ú", "u")
    cadena = Replace(cadena, "É", "E")
    cadena = Replace(cadena, "Í", "I")
    cadena = Replace(cadena, "Ó", "O")
    cadena = Replace(cadena, "Ú", "U")
End If

If QuitarParentesis Then
    cadena = Replace(cadena, "(", "_")
    cadena = Replace(cadena, ")", "_")
End If

cadena = Replace(cadena, "'", "")

' devuelve la cadena
ClrStr = cadena
 
End Function


Public Function getDataComboSel(obj As DataCombo, Optional ByVal default As Long = -1) As Long
    
On Error GoTo getDataComboSel_Error

    If Trim(obj.BoundText) = "" Then
        getDataComboSel = default
    ElseIf Not IsNumeric(obj.BoundText) Then
        getDataComboSel = default
    Else
        getDataComboSel = CLng(obj.BoundText)
    End If

On Error GoTo 0
    Exit Function
getDataComboSel_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getDataComboSel of Módulo Funciones"
    getDataComboSel = default
End Function


Public Function OrdernarValoresArrayInt(ByRef prmArr() As Integer, ByVal min As Integer, ByVal Max As Integer) As Integer()

Dim newArr() As Integer
Dim iCont As Integer, intTemp As Integer
Dim blnContinuar As Boolean
    
    ReDim Preserve newArr(Max)
    
    For iCont = min To Max
        newArr(iCont) = prmArr(iCont)
    Next iCont
    
    Do
        blnContinuar = False
        
        For iCont = min + 1 To Max
            If newArr(iCont - 1) > newArr(iCont) Then ' Si el valor anterior es mayor que el actual
                intTemp = newArr(iCont)
                newArr(iCont) = newArr(iCont - 1)
                newArr(iCont - 1) = intTemp
                blnContinuar = True
            End If
        Next iCont
        
    Loop While blnContinuar

    OrdernarValoresArrayInt = newArr


End Function


Public Function OrdernarValoresArrayLong(ByRef prmArr() As Long, ByVal min As Long, ByVal Max As Long) As Long()

Dim newArr() As Long
Dim iCont As Long, intTemp As Long
Dim blnContinuar As Boolean
    
    ReDim Preserve newArr(Max)
    
    For iCont = min To Max
        newArr(iCont) = prmArr(iCont)
    Next iCont
    
    Do
        blnContinuar = False
        
        For iCont = min + 1 To Max
            If newArr(iCont - 1) > newArr(iCont) Then ' Si el valor anterior es mayor que el actual
                intTemp = newArr(iCont)
                newArr(iCont) = newArr(iCont - 1)
                newArr(iCont - 1) = intTemp
                blnContinuar = True
            End If
        Next iCont
        
    Loop While blnContinuar

    OrdernarValoresArrayLong = newArr


End Function


Public Sub EliminarElementoArrayInt(ByRef prmArr() As Integer, ByVal intIndiceEliminado As Integer, ByVal Max As Integer)
        
        Dim iCont As Integer
        
        For iCont = intIndiceEliminado To Max - 1
            prmArr(iCont) = prmArr(iCont + 1)
        Next
        
        ReDim Preserve prmArr(Max - 1)
        
End Sub


Public Function JoinIntArr2String(ByRef prmIntArray() As Integer, Optional delimiter As String = "", Optional ByVal min As Long = -1)
    Dim Max As Long, cont As Long
    Dim strres As String
    
    If min = -1 Then
        min = LBound(prmIntArray)
    End If
    Max = UBound(prmIntArray)
    
    strres = ""
    
    For cont = min To Max
        strres = strres & CStr(prmIntArray(cont)) & delimiter
    Next cont

    strres = Left(strres, Len(strres) - Len(delimiter))

    JoinIntArr2String = strres

End Function

Public Function JoinLongArr2String(ByRef prmLngArray() As Long, Optional delimiter As String = "")
    Dim min As Long, Max As Long, cont As Long
    Dim strres As String
    
    min = LBound(prmLngArray)
    Max = UBound(prmLngArray)
    
    strres = ""
    
    For cont = min To Max
        strres = strres & CStr(prmLngArray(cont)) & delimiter
    Next cont

    strres = Left(strres, Len(strres) - Len(delimiter))

    JoinLongArr2String = strres

End Function

Public Function SplitString2IntArr(ByVal prmCad As String) As Integer()

    Dim x() As Integer
    Dim cont As Integer
    
        ReDim x(Len(prmCad))
    
        For cont = 1 To Len(prmCad)
            x(cont) = CInt(Mid(prmCad, cont, 1))
        Next cont
    
        SplitString2IntArr = x
    
End Function

Public Function SplitString2LongArr(ByVal prmCad As String) As Long()

    Dim x() As Long
    Dim cont As Long
    
        ReDim x(Len(prmCad))
    
        For cont = 1 To Len(prmCad)
            x(cont) = CLng(Mid(prmCad, cont, 1))
        Next cont
    
        SplitString2LongArr = x
    
End Function


Public Function getFechaServidor() As Date
    Dim rs As ADODB.Recordset
    
    Set rs = datos_bd("SELECT LOCALTIMESTAMP as FECHA")
    
    rs.MoveFirst
    
    getFechaServidor = rs(0)
    
    Set rs = Nothing
    
End Function


Public Function DirTempLocalCreate() As String
'    On Error Resume Next
    Dim strTemp As String, fso As New FileSystemObject
    Dim strRnd As String
    'Create a buffer
   On Error GoTo DirTempLocalCreate_Error

    strTemp = String(100, Chr$(0))
    'Get the temporary path
    GetTempPath 100, strTemp
    'strip the rest of the buffer
    strTemp = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
    
    Randomize
    
    strRnd = CStr(Int((100000 * Rnd) + 1))
    strRnd = String(8, "_") & strRnd
    strRnd = Right(strRnd, 8)
    
    strTemp = strTemp & strRnd
    If Not fso.FolderExists(strTemp) Then
        fso.CreateFolder strTemp
    End If
    Set fso = Nothing
    
    DIRECTORIO_TEMPORAL = strTemp & "\"
   
   On Error GoTo 0
   Exit Function
DirTempLocalCreate_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DirTempLocalCreate of Módulo Funciones"
End Function

Public Function DirTempLocalDelete() As String
'    On Error Resume Next
    Dim fso As New FileSystemObject
   On Error GoTo DirTempLocalDelete_Error

    fso.DeleteFolder Left(DIRECTORIO_TEMPORAL, Len(DIRECTORIO_TEMPORAL) - 1), True
    Set fso = Nothing

   On Error GoTo 0
   Exit Function
DirTempLocalDelete_Error:
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DirTempLocalDelete of Módulo Funciones"
   log "Error " & Err.Number & " (" & Err.Description & ") in procedure DirTempLocalDelete of Módulo Funciones"

End Function


Public Function EscanearATemp(Optional ByVal ExtensionArchivo As String = "pdf") As String

    Dim rutaEscaner As String
    Dim nombreNuevo As String
    Dim archivoEscaneado As String
    
    nombreNuevo = ""
        
    frmEscaner.Show 1
    
    If documento_escaner <> "" Then
        nombreNuevo = ""
        nombreNuevo = Replace(InputBox("Introduzca nombre para el archivo Escaneado, sin ninguna extensión. " & vbCrLf & "Si lo deja en blanco, se generará de forma automática", "Escaneando Archivo", , Screen.Width / 3, (Screen.Height / 3)), ":", " ")
    
        If Trim(nombreNuevo) = "" Then
            Randomize
            nombreNuevo = CStr(Int((100000 * Rnd) + 1))
            nombreNuevo = String(8, "_") & nombreNuevo
            nombreNuevo = Right(nombreNuevo, 8)
        End If
        
        nombreNuevo = DIRECTORIO_TEMPORAL & nombreNuevo & "." & ExtensionArchivo
        
        Dim fso As New FileSystemObject
        Call fso.CopyFile(documento_escaner, nombreNuevo, True)
        Set fso = Nothing
        
    End If
            
    EscanearATemp = nombreNuevo

End Function

Public Function Eliminar_Caracteres_Archivo(cadena As String)
    Dim nombreNuevo As String
    nombreNuevo = cadena
    nombreNuevo = Replace(nombreNuevo, ":", "")
    nombreNuevo = Replace(nombreNuevo, "/", "")
    nombreNuevo = Replace(nombreNuevo, "\", "")
    nombreNuevo = Replace(nombreNuevo, "*", "")
    nombreNuevo = Replace(nombreNuevo, "?", "")
    nombreNuevo = Replace(nombreNuevo, "<", "")
    nombreNuevo = Replace(nombreNuevo, ">", "")
    nombreNuevo = Replace(nombreNuevo, "'", "")
    Eliminar_Caracteres_Archivo = nombreNuevo
End Function

Public Sub OrdernarColumnaMsGrid(ByRef grid As MSFlexGrid, ByVal sort_column As Integer)
Static m_SortColumn As Integer
Static m_SortOrder As Integer
    ' Hide the FlexGrid.
    grid.visible = False
    grid.Refresh

    ' Sort using the clicked column.
    grid.Col = sort_column
    grid.ColSel = sort_column
    grid.Row = 0
    grid.RowSel = 0

    ' If this is a new sort column, sort ascending.
    ' Otherwise switch which sort order we use.
    If m_SortColumn <> sort_column Then
        m_SortOrder = flexSortGenericAscending
    ElseIf m_SortOrder = flexSortGenericAscending Then
        m_SortOrder = flexSortGenericDescending
    Else
        m_SortOrder = flexSortGenericAscending
    End If
    grid.Sort = m_SortOrder

    ' Restore the previous sort column's name.
'    If m_SortColumn >= 0 Then
'        grid.TextMatrix(0, m_SortColumn) = _
'            Mid$(grid.TextMatrix(0, m_SortColumn), 3)
'    End If

    ' Display the new sort column's name.
    m_SortColumn = sort_column
'    If m_SortOrder = flexSortGenericAscending Then
'        If Left(grid.TextMatrix(0, m_SortColumn), 2) = "> " Or Left(grid.TextMatrix(0, m_SortColumn), 2) = "< " Then
'            grid.TextMatrix(0, m_SortColumn) = "> " & _
'                Mid(grid.TextMatrix(0, m_SortColumn), 3)
'        Else
'            grid.TextMatrix(0, m_SortColumn) = "> " & _
'                grid.TextMatrix(0, m_SortColumn)
'        End If
'    Else
'        If Left(grid.TextMatrix(0, m_SortColumn), 2) = "> " Or Left(grid.TextMatrix(0, m_SortColumn), 2) = "< " Then
'            grid.TextMatrix(0, m_SortColumn) = "< " & _
'                Mid(grid.TextMatrix(0, m_SortColumn), 3)
'        Else
'            grid.TextMatrix(0, m_SortColumn) = "< " & _
'                grid.TextMatrix(0, m_SortColumn)
'        End If
'    End If

    ' Display the FlexGrid.
    grid.visible = True
    
End Sub

Public Function MostrarInforme(codmuestra As Long) As Boolean
   On Error GoTo Informe_Error
'M1051    Dim oTD As New clsTipos_documentos
'M1051    If oTD.esManual(codmuestra) Then
'M1051        frmPrevisualizarWord.Show
'M1051    Else
        frmPrevisualizar.PK = codmuestra
        frmPrevisualizar.Show 1
'M1051    End If
    MostrarInforme = True
    Set oTD = Nothing
   On Error GoTo 0
   Exit Function

Informe_Error:
    MostrarInforme = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MostrarInforme"
End Function

Public Function Salir()
    DirTempLocalDelete
    Dim oemp As New clsUsuarios
    oemp.deslogonear (USUARIO.getID_EMPLEADO)
    Set oemp = Nothing
    Set gFSO = Nothing
    End
End Function

'Public Sub FTP(fichero_remoto As String, fichero_local As String, Abrir As Boolean)
'    Dim oFtp As New clsFTP
'   On Error GoTo trae_fichero_Error
'    If fichero_local = "" Then
'        fichero_local = App.Path & "\DOC.PDF"
'    End If
'    With oFtp
'        .Servidor = ReadINI(App.Path + "\config.ini", "server", "FTP")
'        .usuario = "geslab"
'        .PassWord = "aer0p0lis"
'        .ConectarFtp
'        .TipoTransferencia = [ BINARIO ]
'        .ObtenerArchivo fichero_remoto, fichero_local, True
'        .Desconectar
'    End With
'    If Abrir = True Then
'        Dim iret As Long
'        iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
'    End If
'   On Error GoTo 0
'   Exit Sub
'
'trae_fichero_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure trae_fichero of Módulo globales"
'End Sub

Public Function calcularFechaFinalizacion(fechaInicio As Date, numeroDias As Integer) As Date
    ' Devuelve la fecha final quitando los sabados y domingos
    Dim fechaAux As Date
   On Error GoTo ChecarTiempo_Error

    DIAS = numeroDias
    fechaAux = fechaInicio
    While DIAS > 0
        fechaAux = DateAdd("d", 1, fechaAux)
        If Weekday(fechaAux) <> vbSaturday And Weekday(fechaAux) <> vbSunday Then
            DIAS = DIAS - 1
        End If
    Wend
    calcularFechaFinalizacion = fechaAux

   On Error GoTo 0
   Exit Function

ChecarTiempo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularFechaFinalizacion of Módulo Funciones"
End Function

Public Function Solo_Numeros(ByRef sText As String) As String
    Dim sActualChar                 As String * 1
    Dim lTotalChar                  As Long
    Dim x                           As Long
    
    lTotalChar = LenB(sText) \ 2
  
    If CBool(lTotalChar) Then
        For x = 1 To lTotalChar
            sActualChar = Mid$(sText, x, 1)
            If IsNumeric(sActualChar) Then Solo_Numeros = Solo_Numeros & sActualChar
        Next
    End If
    
End Function

'M1171-I
Public Sub envioCorreoTramite(NUMERO_PAQUETES As Long)
'-------------- ENVÍO AUTOMÁTICO DE CORREO A LISTA DE DISTRIBUCIÓN TRAS TRAMITACIÓN DE PAQUETE ------'
    If USUARIO.getPER_TRAMITACION_CONTRATA = True Then
       Exit Sub
    End If
    
    If NUMERO_PAQUETES = 0 Then
       Exit Sub
    End If
    log ("Envío correo de trámite")
    
    Dim destinatario As String
    Dim mensaje As String
    Dim ASUNTO As String
    Dim cabecera As String
    Dim oParametro As New clsParametros
    Dim oPaquete As New clsSC_Paquetes
    
    Dim RSPaquetes As ADODB.Recordset
    Dim rs As ADODB.Recordset

    oParametro.Carga parametros.PARAM_CORREO_DISTRIBUCION_TRAMITE, ""
    destinatario = oParametro.getVALOR
    
    If destinatario <> "" Then
        Set RSPaquetes = oPaquete.Listado_parcial(NUMERO_PAQUETES)
        If RSPaquetes.RecordCount > 0 Then
            Do
                'Cuerpo del correo (uno por paquete)
                ASUNTO = "Tramitación de pedido a proveedor. Código : " & RSPaquetes("CODIGO_SC")
                mensaje = "Se ha creado el siguiente pedido: " & vbNewLine & vbNewLine
        
                mensaje = mensaje & vbNewLine & " Código : " & RSPaquetes("CODIGO_SC")
                mensaje = mensaje & vbNewLine & " Presupuesto : " & RSPaquetes("PRESUPUESTO")
                mensaje = mensaje & vbNewLine & " Generada por : " & "(" & USUARIO.getUSUARIO & ") " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
                
                ' LISTADO DETALLE
                Select Case CInt(RSPaquetes("TIPO"))
                    Case TOBJETO_SC_DETERMINACIONES
                        cabecera = "LISTA DE DETERMINACIONES"
                        ASUNTO = ASUNTO & " (DETERMINACIONES)"
                        Set rs = oPaquete.Listado_muestras_determinaciones(RSPaquetes("ID_PAQUETE"), RSPaquetes("EDICION"))
                    Case TOBJETO_SC_EFICACIA
                        cabecera = "LISTA DE ENSAYOS"
                        ASUNTO = ASUNTO & " (ENSAYOS)"
                        Set rs = oPaquete.Listado_muestras_ensayos(RSPaquetes("ID_PAQUETE"), RSPaquetes("EDICION"))
                End Select
            
                If rs.RecordCount > 0 Then
                     mensaje = mensaje & vbNewLine
                     mensaje = mensaje & vbNewLine & "--------------------------------------------------------------------------------------------------------"
                     mensaje = mensaje & vbNewLine & cabecera
                     mensaje = mensaje & vbNewLine & "--------------------------------------------------------------------------------------------------------"
                     Do
                         mensaje = mensaje & vbNewLine
                         mensaje = mensaje & Format(Left(rs(1), 40), "!" & String(40, "@")) & "    "
                         mensaje = mensaje & Format(Left(rs(2), 40), "!" & String(40, "@")) & "    "
                         mensaje = mensaje & Format(Left(rs(6), 8), "!" & String(8, "@")) & "     "
                         rs.MoveNext
                     Loop Until rs.EOF
                End If
                Set rs = Nothing
                mensaje = mensaje & vbNewLine
                mensaje = mensaje & vbNewLine
                mensaje = mensaje & " Mensaje enviado automáticamente desde Geslab el " & Format(Date, "dd-mm-yyyy") & " a las " & Format(Time, "hh:mm:ss")
                
               ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
               ' ret = Enviar_Mail_CDO("daniel.gallardo@ixitec.net", ASUNTO, mensaje, vbNullString)
                
                'siguiente paquete
                RSPaquetes.MoveNext
            Loop Until RSPaquetes.EOF
        End If
    End If
    Set oParametro = Nothing
End Sub

'M1171-F
Public Function verificar_impresion(muestra As Long) As Integer
    Dim c As String
    Dim rs As ADODB.Recordset
    c = "select * from impresion where muestra_id = " & muestra
    Set rs = datos_bd(c)
    If rs.RecordCount = 0 Then
        verificar_impresion = 0
    Else
        If rs("estado") = 3 Then
            verificar_impresion = 2
        Else
            verificar_impresion = 1
        End If
    End If
End Function

'M1359-I
Public Function sDigitoControlBanco(sEntidad As String, sSucursal As String, sCuenta As String) As String
'DEVUELVE EL DÍGITO DE CONTROL
    Dim Temporal As Integer
    
   On Error GoTo sDigitoControlBanco_Error

    Temporal = 0
    Temporal = Temporal + Mid(sEntidad, 1, 1) * 4
    Temporal = Temporal + Mid(sEntidad, 2, 1) * 8
    Temporal = Temporal + Mid(sEntidad, 3, 1) * 5
    Temporal = Temporal + Mid(sEntidad, 4, 1) * 10
    Temporal = Temporal + Mid(sSucursal, 1, 1) * 9
    Temporal = Temporal + Mid(sSucursal, 2, 1) * 7
    Temporal = Temporal + Mid(sSucursal, 3, 1) * 3
    Temporal = Temporal + Mid(sSucursal, 4, 1) * 6
    Temporal = 11 - (Temporal Mod 11)
    
    If Temporal = 11 Then
        sDigitoControlBanco = "0"
    ElseIf Temporal = 10 Then sDigitoControlBanco = "1"
        Else: sDigitoControlBanco = Format(Temporal, "0")
    End If
    
    Temporal = 0
    Temporal = Temporal + Mid(sCuenta, 1, 1) * 1
    Temporal = Temporal + Mid(sCuenta, 2, 1) * 2
    Temporal = Temporal + Mid(sCuenta, 3, 1) * 4
    Temporal = Temporal + Mid(sCuenta, 4, 1) * 8
    Temporal = Temporal + Mid(sCuenta, 5, 1) * 5
    Temporal = Temporal + Mid(sCuenta, 6, 1) * 10
    Temporal = Temporal + Mid(sCuenta, 7, 1) * 9
    Temporal = Temporal + Mid(sCuenta, 8, 1) * 7
    Temporal = Temporal + Mid(sCuenta, 9, 1) * 3
    Temporal = Temporal + Mid(sCuenta, 10, 1) * 6
    Temporal = 11 - (Temporal Mod 11)
    
    If Temporal = 11 Then
        sDigitoControlBanco = sDigitoControlBanco + "0"
    ElseIf Temporal = 10 Then sDigitoControlBanco = sDigitoControlBanco + "1"
        Else: sDigitoControlBanco = sDigitoControlBanco + Format(Temporal, "0")
    End If

   On Error GoTo 0
   Exit Function

sDigitoControlBanco_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sDigitoControlBanco of Módulo Funciones"
End Function
'M1359-F
Public Function CalcularIBAN(pais As String, cuenta As String) As String
    ' Recibe el pais con 2 letras (ES para Espa&ntilde;a)
    ' Recibe el n&uacute;mero de cuenta
    Dim Letras As String * 26
    Dim iban As String
    Dim Dividendo As Integer
    Dim resto As Integer
 
    ' Quita los posibles espacios
   On Error GoTo CalcularIBAN_Error

    cuenta = Replace(Replace(cuenta, " ", ""), "-", "")
 
    ' Calcula el valor de las letras, las quita y añade el valor al final
    Letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    iban = cuenta & CStr(InStr(1, Letras, Left(pais, 1)) + 9) & CStr(InStr(1, Letras, Right(pais, 1)) + 9) & "00"
 
    For contador = 1 To Len(iban)
        Dividendo = resto & Mid(iban, contador, 1)
        resto = Dividendo Mod 97
    Next contador
 
'    CalcularIBAN = "IBAN" & pais & Format((98 - Resto), "00") & cuenta
    CalcularIBAN = pais & Format((98 - resto), "00") & cuenta

   On Error GoTo 0
   Exit Function

CalcularIBAN_Error:

    MsgBox "Error al calcular el IBAN.", vbCritical, App.Title
 
End Function
Public Function IsLoadForm(ByVal FORMULARIO As String, Optional activar As Boolean) As Boolean
    Dim rtn As Integer, i As Integer
    rtn = False
    Do Until i > Forms.Count - 1 Or rtn
        If UCase(Forms(i).Name) = UCase(FORMULARIO) Then rtn = True
        i = i + 1
    Loop
    If rtn Then
        If Not IsMissing(activar) Then
            If activar Then
                Forms(i - 1).WindowState = vbNormal
            End If
        End If
    End If
    IsLoadForm = rtn
End Function
Public Function GetDirTemp() As String
    If Environ$("temp") <> vbNullString Then
       GetDirTemp = Environ$("tmp")
    End If
End Function
Public Sub colorearLista(lista As ListView, fila As Integer, color As Long)
    Dim i As Integer
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
End Sub
Public Function abrirRegistroMuestra(idMuestra As Long)
    If idMuestra > 0 Then
        Dim oMuestra As New clsMuestra
        oMuestra.CargaMuestra (idMuestra)
        Select Case oMuestra.getANALISIS_MODIFICADO
            Case 2 ' Control de eficacia
                With frmCE_Resultados
                    .PK_ID_MUESTRA = idMuestra
                    .Show 1
                End With
            Case 3 ' Sellante
                frmSE_Resultados.Show 1
            Case 5 ' Plasma
                If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_ROCKWELL Or _
                   oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_BRINELL Or _
                   oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_VICKERS Then
                    With frmPlasma_Dureza
                        .PK = idMuestra
                        .Show 1
                    End With
                ElseIf oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_SHORE_PIEZAS Then
                    With frmPlasma_Dureza_Shore
                        .PK = idMuestra
                        .Show 1
                    End With
                ElseIf oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.ENSAYO_TRACCION Then
                    With frmPlasma_ETR
                        .PK = idMuestra
                        .Show 1
                    End With
                Else
                    With frmPlasma_Resultados
                        .PK = idMuestra
                        .Show 1
                    End With
                End If
            Case Else
                frmDeterminaciones.Show 1
        End Select
    End If
End Function
