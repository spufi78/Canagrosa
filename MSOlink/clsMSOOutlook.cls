VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMSOOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Const PrinterIdle = 3
Const PrinterPrinting = 4
Const PrinterWarmingUp = 5

Const recDepth = 3
Private Const EM_FMTLINES = &HC8

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
 (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public nombreGenerado As String
Public asuntoGenerado As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function Guarda_mensaje_outlook(ByRef prmConexionBD As ADODB.Connection, ByRef prmUsuario As Object, ByVal prmDirOrigen As String, ByVal prmDirDestino As String, Optional prmNombre As String = "") As Boolean
Attribute Guarda_mensaje_outlook.VB_Description = "Devuelve la ruta final del archivo en formato .msg de Outlook"
Attribute Guarda_mensaje_outlook.VB_UserMemId = 0

    Dim oO As New OUTLOOK.Application
    Dim oM As MailItem
    Dim strNombre As String, strDir As String, strDestino As String
    Dim randomCont As Integer, blnFormatoHTML As Boolean, strPrinter_default As String
    Dim oFrm As frmSelMsg
    Dim oImp As clsImpresion
    Dim strNombreArchivo As String
    Dim intIntentos As Integer, lngID As Long
    
    ' Aplica la conexion
    Set conn = prmConexionBD
    
    Set oFrm = New frmSelMsg
    oFrm.Show vbModal
    
    If Not oFrm.Aceptar Then
        Guarda_mensaje_outlook = False
        Set oO = Nothing
        Set oM = Nothing
        Set oFrm = Nothing
        Set oImp = Nothing
        Exit Function
    End If
        
    
    Set oM = oFrm.correo
    
    Unload oFrm
    Set oFrm = Nothing
    
    strDir = prmDirOrigen
    strDestino = prmDirDestino
    
    If Right(Trim(strDir), 1) <> "\" Then
        strDir = strDir & "\"
    End If
    
    If Right(Trim(strDestino), 1) <> "\" Then
        strDestino = strDestino & "\"
    End If
    
    
    ' Averiguar el nombre
    If Trim(prmNombre) <> "" Then
        strNombre = prmNombre
    Else
        strNombre = Format(oM.ReceivedTime, "yyyymmddHhNnSs") & "-" & Eliminar_Caracteres_Archivo(oM.SenderEmailAddress)
    End If
    nombreGenerado = strNombre
    asuntoGenerado = oM.Subject
    
    If Trim(strNombre) <> "" Then
        
        If oM.BodyFormat = olFormatPlain Or oM.BodyFormat = olFormatUnspecified Then
            strNombre = strDir & strNombre & ".txt"
            oM.SaveAs strNombre, olTXT
        Else
            strNombre = strDir & strNombre & ".htm"
            oM.SaveAs strNombre, olHTML
        End If
        
        Set oImp = New clsImpresion
        
        With oImp
            .setEMPLEADO_ID = prmUsuario.getID_EMPLEADO
            .setTIPO = 41
            .setMUESTRA_ID = 0
            .setRUTA_ORIGEN = Replace(strNombre, "\", "/")
            .setRUTA_DESTINO = Replace(strDestino, "\", "/")
            .setPUESTO = prmUsuario.getUSO
            lngID = .Insertar
            
            intIntentos = 0
            While .CARGAR(lngID)
                DoEvents
                Sleep 1000
                DoEvents
                intIntentos = intIntentos + 1
                If intIntentos > 10 Then
                    If MsgBox("El proceso para guardar el correo electrónico seleccionado se está demorando. ¿Desea continuar esperando?", vbInformation + vbYesNo, "Guardar Correo a PDF") = vbNo Then
                        Guarda_mensaje_outlook = False
                        Set oO = Nothing
                        Set oM = Nothing
'                        Set oFrm = Nothing
                        Set oImp = Nothing

                        Exit Function
                    Else
                        intIntentos = 0
                    End If
                End If
                
            Wend
            
        End With
        ' una vez terminada la conversion, elimina el origina en htm/txt
        On Error Resume Next
        Kill strNombre
        On Error GoTo Error_Guarda_mensaje_outlook
    End If
    
    Guarda_mensaje_outlook = True
    Set oO = Nothing
    Set oM = Nothing
'    Set oFrm = Nothing
    Set oImp = Nothing
Exit Function
Error_Guarda_mensaje_outlook:

    Guarda_mensaje_outlook = False
    Set oO = Nothing
    Set oM = Nothing
    Set oFrm = Nothing
    Set oImp = Nothing
End Function






