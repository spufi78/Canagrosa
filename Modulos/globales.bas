Attribute VB_Name = "globales"
Option Explicit
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public MODO_PRUEBA As Boolean
Public USUARIO As clsUsuarios
Public NOMBRE_PC As String
' Variables Globales para la ventana de Clientes
Public opCliente As String
Public id_cliente As Integer

Public id_tipoMuestra As Integer
Public id_analisis As Integer

' Determinaciones
Public id_determinaciones() As Integer ' dinamico
Public auxiliar As Variant 'para enviar datos de un formotro form
Public gmuestra As Long
Public gdeterminacion As Long

' Baños
Public num_banos As Integer
Public cliente_banos As Integer
Public analisis_banos As Integer
Public plantilla_bano() As Integer

' Muestras
Public muestras() As Long
Public ETIQUETAS() As Long


' Plantillas
Public PLANTILLA As Integer
Public nueva_plantilla As Integer
Public RUTA As String
Public referencia_word As String
Public referencia_pdf As String

' Documentos de pago
Public numero_documentos_pago As Integer
Public documentos_pago() As Long

' Mantenimiento de tablas
 Public glogin As Integer
'Public gtipo_analisis As Integer
'Public gtipo_determinacion As Integer
'public gtipo_muestra As Integer
'Public gformula As Integer
'Public gbano As Integer
Public gespecifico As Integer
'Public gcliente As Integer

'E0061-I
'Se elimina esta variable global para hacer uso de PK
'Public gproveedor As Integer
'E0061-F

Public gempleado As Integer
Public ganomalia As Long
Public gbsm As Long
Public grecarga As Long
'Public greactivoex As Long
'Public gbotereactivoex As Long
Public greactivopr As Long
Public gpedido As Long
Public gTipo_Bote As Long
'Public gdependencia_deter As Integer
'Public gdependencia_campo As Integer
Public gdoc As Long
Public pegatina As Integer
Public gAgenda As Long
Public gid_concepto As Long
'Public gOferta As Long
'Public gOperario As Long
'Motivo
Public MOTIVO As String
'Indicadores
Public gindicadores_campos As Long
Public gindicadores As Long
'Alodine
Public gAlodine As Long
Public glote As Long
Public gCE_Ficha As Long
' Se
Public gSE_Sellante As Long
'CA
'Public gCA_documento As Long

Public gID As Long

Public documento_escaner As String
Public documento_escaner_nombre As String
Public documento_escaner_eliminar As Boolean
Public DIRECTORIO_TEMPORAL As String

Public G_TRAZABILIDAD_ERROR As String
Public gFSO As New Scripting.FileSystemObject

Public Sub enviar_informe_error(MUESTRA As Long, ERROR As String)
    On Error Resume Next
    Dim sPara As String
    Dim sAsunto As String
    Dim sMensaje As String
    Dim sFichero_Log As String
    sPara = "informatica@canagrosa.com"
    If MUESTRA = 0 Then
        sAsunto = "Error al generar informe"
    Else
        sAsunto = "Error al generar la muestra (ID : " & MUESTRA & ")"
    End If
    sMensaje = sMensaje & vbNewLine & "*****************************"
    If MUESTRA <> 0 Then
        sMensaje = sMensaje & vbNewLine & " ID : " & MUESTRA
    End If
    sMensaje = sMensaje & vbNewLine & " FECHA : " & Date
    sMensaje = sMensaje & vbNewLine & " HORA : " & Time
    sMensaje = sMensaje & vbNewLine & " ERROR : " & ERROR
    sMensaje = sMensaje & vbNewLine & "*****************************"
    sFichero_Log = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date) & "\pdf\" & Format(Date, "yyyy-mm-dd") & " PDF.txt"
    Enviar_Mail_CDO sPara, sAsunto, sMensaje, sFichero_Log
End Sub

