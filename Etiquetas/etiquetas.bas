Attribute VB_Name = "etiquetas"
Public Enum ETIQUETAS_TIPOS
    ETIQUETAS_TIPOS_REX = 1
    ETIQUETAS_TIPOS_MUESTRAS = 2
    ETIQUETAS_TIPOS_EQUIPOS_CAL = 3
    ETIQUETAS_TIPOS_EQUIPOS_VER = 4
    ETIQUETAS_TIPOS_EQUIPOS = 5
    ETIQUETAS_TIPOS_RPR = 6
End Enum

Private Declare Function GetProfileString Lib "KERNEL32" Alias "GetProfileStringA" ( _
    ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long) As Long
Public Function generaFirma(ID_EMPLEADO As Long) As String
    Dim rs As ADODB.Recordset
    Dim c As String
    c = "select * from usuarios_firmas where ID_EMPLEADO = " & ID_EMPLEADO
    Set rs = datos_bd(c)
    Dim firma As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs("FIRMA")
        On Error Resume Next
        Dim ruta As String
        ruta = App.Path & "\firmas"
        MkDir ruta
        Dim fichero
        fichero = ruta & "\" & ID_EMPLEADO & ".jpg"
        mystream.SaveToFile fichero, adSaveCreateOverWrite
        mystream.Close
        firma = fichero
    End If
    generaFirma = firma
End Function

Public Function etiqueta_REACTIVO(LISTA_REACTIVOS As String, USUARIO_ID As Long, carpeta As String, informe As String) As Boolean
    Dim P1() As String
    Dim P2() As String
   On Error GoTo etiqueta_REACTIVO_Error

    ReDim P1(1) As String
    ReDim P2(1) As String
    P1(1) = "FIRMA"
'    Dim oUsuario As New clsUsuarios
'    oUsuario.CARGAR USUARIO_ID
'    P2(1) = Replace(oUsuario.getFIRMA, "/", "\")
    P2(1) = generaFirma(USUARIO_ID)
    With frmReport
        .iniciar
        .informe = carpeta & "\" & Replace(informe, ".rpt", "")
        .criterio = "{botes_ex.ID_BOTE_EX} in [" & LISTA_REACTIVOS & "]"
        .ParametrosNombre = P1
        .ParametrosValores = P2
        .imprimir = True
        .generar
    End With
    Set frmReport = Nothing

   On Error GoTo 0
   Exit Function

etiqueta_REACTIVO_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure etiqueta_REACTIVO of Módulo etiquetas"
End Function
Public Function etiqueta_RPR(LISTA_REACTIVOS As String, USUARIO_ID As Long, carpeta As String, informe As String) As Boolean
    Dim P1() As String
    Dim P2() As String
   On Error GoTo etiqueta_REACTIVO_Error

    ReDim P1(1) As String
    ReDim P2(1) As String
    P1(1) = "FIRMA"
'    Dim oUsuario As New clsUsuarios
'    oUsuario.CARGAR USUARIO_ID
'    P2(1) = Replace(oUsuario.getFIRMA, "/", "\")
    P2(1) = generaFirma(USUARIO_ID)
    With frmReport
        .iniciar
        .informe = carpeta & "\" & Replace(informe, ".rpt", "")
        .criterio = "{rpr_botes.ID_BOTE_PR} in [" & LISTA_REACTIVOS & "]"
        .ParametrosNombre = P1
        .ParametrosValores = P2
        .imprimir = True
        .generar
    End With
    Set frmReport = Nothing
   On Error GoTo 0
   Exit Function

etiqueta_REACTIVO_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure etiqueta_REACTIVO of Módulo etiquetas"
End Function

Public Function etiqueta_CALIBRACION(ID_CALIBRACION As Long, USUARIO_ID As Long, carpeta As String, informe As String) As Boolean
   On Error GoTo etiqueta_CALIBRACION_Error
    ReDim P1(1) As String
    ReDim P2(1) As String
    P1(1) = "FIRMA"
    P2(1) = generaFirma(USUARIO_ID)
    
    With frmReport
        .iniciar
        .informe = carpeta & "\" & Replace(informe, ".rpt", "")
        .criterio = "{eq_calibracion_equipos.ID_CALIBRACION} =" & ID_CALIBRACION
        .ParametrosNombre = P1
        .ParametrosValores = P2
        .imprimir = True
        .generar
    End With
    Set frmReport = Nothing

   On Error GoTo 0
   Exit Function

etiqueta_CALIBRACION_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure etiqueta_CALIBRACION of Módulo etiquetas"
End Function
Public Function etiqueta_VERIFICACION(ID_VERIFICACION As Long, USUARIO_ID As Long, carpeta As String, informe As String) As Boolean
   On Error GoTo etiqueta_CALIBRACION_LIMITACIONES_Error
    ReDim P1(1) As String
    ReDim P2(1) As String
    P1(1) = "FIRMA"
    P2(1) = generaFirma(USUARIO_ID)
    With frmReport
        .iniciar
        .informe = carpeta & "\" & Replace(informe, ".rpt", "")
        .criterio = "{eq_verificacion_equipos.ID_VERIFICACION} =" & ID_VERIFICACION
        .ParametrosNombre = P1
        .ParametrosValores = P2
        .imprimir = True
        .generar
    End With
    Set frmReport = Nothing

   On Error GoTo 0
   Exit Function

etiqueta_CALIBRACION_LIMITACIONES_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure etiqueta_CALIBRACION_LIMITACIONES of Módulo etiquetas"
End Function
Public Function etiqueta_EQUIPO(ID_EQUIPO As Long, USUARIO_ID As Long, carpeta As String, informe As String) As Boolean
   On Error GoTo etiqueta_EQUIPO_Error
    ReDim P1(1) As String
    ReDim P2(1) As String
    P1(1) = "FIRMA"
    P2(1) = generaFirma(USUARIO_ID)

    With frmReport
        .iniciar
        .informe = carpeta & "\" & Replace(informe, ".rpt", "")
        .criterio = "{equipos.ID_EQUIPO} =" & ID_EQUIPO
        .ParametrosNombre = P1
        .ParametrosValores = P2
        .imprimir = True
        .generar
    End With
    Set frmReport = Nothing

   On Error GoTo 0
   Exit Function

etiqueta_EQUIPO_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure etiqueta_EQUIPO of Módulo etiquetas"

End Function
Public Function etiqueta_MUESTRA(ID_MUESTRA As Long, USUARIO_ID As Long, carpeta As String, informe As String) As Boolean

   On Error GoTo etiqueta_MUESTRA_Error

    With frmReport
        .iniciar
        .informe = carpeta & "\" & Replace(informe, ".rpt", "")
        .criterio = "{muestras.ID_MUESTRA} =" & ID_MUESTRA
        .imprimir = True
        .generar
    End With
    Set frmReport = Nothing

   On Error GoTo 0
   Exit Function

etiqueta_MUESTRA_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure etiqueta_MUESTRA of Módulo etiquetas"

End Function
