Attribute VB_Name = "general"
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
'Public conn As ADODB.Connection
Public MODO_PRUEBA As Boolean
Public USUARIO As clsUsuarios
Public referencia_word As String
Public referencia_pdf As String
Public database As String
Public DIRECTORIO_TEMPORAL As String

Public Function datos_bd(ByVal consulta As String, Optional no_log As Boolean) As ADODB.Recordset
    On Error GoTo fallo
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        Dim rs As New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        log (consulta)
        rs.Open consulta
        Set rs.ActiveConnection = Nothing ' DESCONECTAR EL RS
        Set datos_bd = rs
        Set rs = Nothing
    End If
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el acceso a la bd: " & Err.Description
'    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function
Public Function execute_bd(ByVal consulta As String, Optional Actualizar As Boolean) As Boolean
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        log (consulta)
        conn.Execute consulta
    End If
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el execute bd: " & Err.Description
'    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function
Public Function CrearConexionGlobal(conn As ADODB.Connection, user As String, PassWord As String) As Boolean
    Dim ipRegistro As String
    On Error GoTo falloConexion
    ipRegistro = ReadINI(App.Path + "\config.ini", "server", "ip")
    database = ReadINI(App.Path + "\config.ini", "server", "bd")
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DRIVER=" & ReadINI(App.Path + "\config.ini", "SERVER", "DRIVER") & ";" _
                            & "SERVER=" & ipRegistro & ";" _
                            & "DATABASE=" & database & ";" _
                            & "UID=" & BD_USUARIO & ";" _
                            & "PWD=" & BD_PASS & ";" _
                            & "OPTION=3"
'                            & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    conn.Open
    CrearConexionGlobal = True
    Exit Function
falloConexion:
    CrearConexionGlobal = False
    MsgBox "CrearConexionGlobal:" & Err.Source & " (" & Err.Number & ") " & Err.Description, vbCritical, App.Title
End Function
Public Sub Espera(Segundos As Single)
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
Public Function insertar_actualizaciones(ByVal consulta As String, Optional registros As Integer) As Boolean
    On Error Resume Next
    Dim c As String
    c = "INSERT INTO ACTUALIZACIONES (FTIMESTP,CONSULTA,REGISTROS,ACTUALIZADA) VALUES (CURRENT_TIMESTAMP,'" & UCase(Trim(Replace(consulta, "'", "#"))) & "'," & registros & "," & USUARIO.getID_EMPLEADO & ")"
    execute_bd c
'    conn.Execute c
End Function

Public Sub log(datos As String)
    On Error Resume Next
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log"
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date)
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date) & "\pdf"
    On Error GoTo fallo
    Open ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date) & "\pdf\" & Format(Date, "yyyy-mm-dd") & " PDF.txt" For Append As #1
    If Left(datos, 36) <> "SELECT A.MUESTRA_ID,A.TIPO,E.USUARIO" Then
        Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & datos
        frmPDF.txtLog = frmPDF.txtLog & Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & datos & vbNewLine
    End If
    Close
    Exit Sub
fallo:
    Close
    Exit Sub
End Sub
Public Function subindices(texto As String) As Boolean
    subindices = False
    If InStr(1, UCase(texto), "SUP(", vbTextCompare) > 0 Or _
       InStr(1, UCase(texto), "SUB(", vbTextCompare) > 0 Then
        subindices = True
    End If
End Function
Public Sub texto_formateado(texto As String, rango As Range, parentesis As Integer)
    If subindices(texto) = True Then
       If parentesis = 1 Then
        rango.InsertAfter " ("
       End If
       Dim ACTIVO As Boolean
       ACTIVO = False
       h = 1
       While h <= Len(texto)
           If UCase(Mid(texto, h, 3)) = "SUP" Or _
              UCase(Mid(texto, h, 3)) = "SUB" Then
                pos = InStr(h, texto, ")")
                rango.InsertAfter Mid(texto, h + 4, pos - (h + 4))
                If UCase(Mid(texto, h, 3)) = "SUP" Then
                    For i = (pos - (h + 4)) To 1 Step -1
                        rango.Characters(Len(rango.Text) - (i + 1)).Font.Superscript = True
                    Next
                Else
                    For i = (pos - (h + 4)) To 1 Step -1
                        rango.Characters(Len(rango.Text) - (i + 1)).Font.Subscript = True
                    Next
                End If
                ACTIVO = True
                h = pos
           Else
'                Debug.Print Mid(texto, h, 1) & " -> " & Asc(Mid(texto, h, 1))
                If Asc(Mid(texto, h, 1)) <> 10 Then
'                    Select Case Asc(Mid(texto, h, 1))
'                    Case 10
'                        rango.InsertAfter vbNewLine
'                    Case Else
                        rango.InsertAfter Mid(texto, h, 1)
                        If ACTIVO = True Then
                            rango.Characters(Len(rango.Text) - 2).Font.Superscript = False
                            rango.Characters(Len(rango.Text) - 2).Font.Subscript = False
                            ACTIVO = False
                        End If
'                    End Select
                End If
           End If
           h = h + 1
       Wend
       If parentesis = 1 Then
        rango.InsertAfter ")"
        rango.Characters(Len(rango.Text) - 2).Font.Superscript = False
        rango.Characters(Len(rango.Text) - 2).Font.Subscript = False
       End If
    Else
       If parentesis = 1 And Trim(texto) <> "" Then
           rango.InsertAfter " (" & texto & ")"
       Else
        rango.InsertAfter texto
       End If
    End If
End Sub

Public Function fecha_bd(ByVal fecha As String) As String
If Trim(fecha) = "" Then fecha = "0000-00-00"
fecha = Format(fecha, "yyyy-mm-dd")

fecha_bd = fecha

End Function
Function moneda_bd(VALOR As String) As String
    VALOR = IIf(Trim(VALOR) = "", "0", VALOR)
    
    If UCase(ReadINI(App.Path + "\config.ini", "server", "tipo")) = "ACCESS" Then
        moneda_bd = Format(VALOR, "currency")
    Else
        moneda_bd = Replace(Format(VALOR, "0.00"), ",", ".")
    End If
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

Public Sub FTP(fichero_remoto As String, fichero_local As String, Abrir As Boolean)
    Dim oFtp As New clsFTP
   On Error GoTo trae_fichero_Error
    If fichero_local = "" Then
        fichero_local = App.Path & "\DOC.PDF"
    End If
    With oFtp
        .Servidor = ReadINI(App.Path + "\config.ini", "server", "FTP")
        .USUARIO = "geslab"
        .PassWord = "aer0p0lis"
        .ConectarFtp
        .TipoTransferencia = [ BINARIO ]
        .ObtenerArchivo fichero_remoto, fichero_local, True
        .Desconectar
    End With
    If Abrir = True Then
        Dim iret As Long
        iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
    End If
   On Error GoTo 0
   Exit Sub

trae_fichero_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure trae_fichero of Módulo globales"
End Sub
Function moneda(VALOR As String) As String
    If UCase(ReadINI(App.Path + "\config.ini", "server", "tipo")) = "ACCESS" Then
        moneda = Format(VALOR, "currency")
    Else
        moneda = Format(Replace(VALOR, ".", ","), "currency")
    End If
End Function

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
Public Function insertarInformeBD(idMuestra As Long, tipo As Integer) As Boolean
    Dim oMuestra As New clsMuestra
    Dim oDocumentacion As New clsDocumentacion
   On Error GoTo insertarInformeBD_Error
    Dim MTQM As Boolean
    If tipo = C_TIPOS_IMPRESION.VB_AIRBUS Then
        MTQM = True
    Else
        MTQM = False
    End If
    
    insertarInformeBD = True
    oMuestra.CargaMuestra (idMuestra)
    Dim destino As String
    ' Si es una revision de MTQM pero el informe el manual, no se regenera
    If oMuestra.getINFORME_MANUAL = 1 And tipo = C_TIPOS_IMPRESION.VB_AIRBUS Then
        tipo = 1
    End If
    destino = NOMBRE_DOCUMENTO(idMuestra, True, tipo, oMuestra.getULT_EDICION_IMP) & ".pdf"
    oDocumentacion.SubirInforme idMuestra, oMuestra.getULT_EDICION_IMP, destino, referencia_pdf, MTQM, oMuestra.getANNO
    Set oDocumentacion = Nothing
   On Error GoTo 0
   Exit Function

insertarInformeBD_Error:
    insertarInformeBD = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertarInformeBD of Módulo general"

End Function

Public Function insertarInformeEquipoBD(ID As Long, tipo As Integer, EDICION As Integer) As Boolean
    insertarInformeEquipoBD = False
    Dim fichero As String
    Dim nombre As String
    nombre = CStr(ID) & ".pdf"
    fichero = App.Path & "\certificados\" & nombre
    SubirInformeEquipo ID, tipo, EDICION, fichero, nombre, Year(Date)
    insertarInformeEquipoBD = True
   On Error GoTo 0
   Exit Function

insertarInformeBD_Error:
    insertarInformeEquipoBD = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertarInformeBD of Módulo general"

End Function
Public Sub cargar_botones(ByVal frm As Form)
End Sub

Public Function SubirInformeEquipo(ID As Long, tipo As Integer, EDICION As Integer, fichero As String, nombre As String, ANNO As Integer) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    Dim tabla As String
    tabla = "equipos_informes"
    ' Borramos si existe previamente
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE ID = " & ID & " AND EDICION = " & EDICION
    ' Lo insertamos
    rs.Open "SELECT * FROM " & tabla & " WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!ID = ID
    rs!tipo = tipo
    rs!EDICION = EDICION
    rs!ANNO = ANNO
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    rs.Update
    rs.Close
    mystream.Close
    SubirInformeEquipo = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirInformeEquipo = "Error " & Err.Number & " (" & Err.Description & ") in : " & MUESTRA_ID & " -> " & fichero & " -> " & nombre
End Function

Public Sub OrdernarColumnaMsGrid(ByRef grid As MSFlexGrid, ByVal sort_column As Integer)
Static m_SortColumn As Integer
Static m_SortOrder As Integer
    ' Hide the FlexGrid.
    grid.Visible = False
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
    grid.Visible = True
    
End Sub
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

Public Function recuperaIVA() As Integer
    Dim oParametros As New clsParametros
    If oParametros.Carga(parametros.IVA, "") Then
        recuperaIVA = CInt(oParametros.getVALOR)
    Else
        recuperaIVA = 0
    End If
    Set oParametros = Nothing
End Function
