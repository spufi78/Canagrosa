VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmPrevisualizarPDF 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14805
   Icon            =   "frmPrevisualizarPDF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   14805
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin AcroPDFLibCtl.AcroPDF pdf1 
      Height          =   7800
      Left            =   495
      TabIndex        =   3
      Top             =   135
      Width           =   13650
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.Frame frmop 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   12645
      TabIndex        =   0
      Top             =   8190
      Width           =   2085
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ESC-Salir"
         Height          =   915
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   45
         Width           =   960
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   915
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir ultima edición generada"
         Top             =   45
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmPrevisualizarPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Public tipo As Integer
Public ruta As String
Public NOMBRE As String

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
   On Error GoTo cmdImprimir_Click_Error
' Modificacion para abrir el cuadro de dialogo de impresion en lugar de envio al servidor

    MOTIVO = ""
    frmMotivo.Show 1
    If Trim(MOTIVO) = "" Then
        MsgBox "Para imprimir el documento es necesario introducir un motivo.", vbInformation, App.Title
        Exit Sub
    End If
    Me.MousePointer = 11
'    Dim oimp As New clsImpresion
'    With oimp
'        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
'        .setMUESTRA_ID = PK
'        .setPUESTO = USUARIO.getUSO
'        If TIPO = CALIDAD_VIDA_TIPOS.CALIDAD_VIDA_TIPOS_DOCUMENTO Then
'            .setTIPO = 55
'        Else
'            .setTIPO = 56
'        End If
'        .Insertar
'    End With
'   set oimp = nothing
    Dim oVida As New clsCa_documentos_vida
    With oVida
        .setIDENTIFICADOR = PK
        .setMOTIVO = MOTIVO
        .setTIPO_ID = tipo
        .setSUBTIPO_ID = CALIDAD_VIDA_SUBTIPOS.CALIDAD_VIDA_SUBTIPOS_IMPRESION
        .setUSUARIO = USUARIO.getID_EMPLEADO
        .Insertar
    End With
    ' Correo Documento calidad
    Dim asunto As String
    Dim mensaje As String
    If tipo = CALIDAD_VIDA_TIPOS.CALIDAD_VIDA_TIPOS_DOCUMENTO Then
        Dim oDoc As New clsCa_documentos
        oDoc.Carga PK
        asunto = "Impresión de documento controlado : " & oDoc.getCODIGO
        mensaje = vbNewLine & "Codigo : " & oDoc.getCODIGO
        mensaje = mensaje & vbNewLine & "Documento : " & oDoc.getNOMBRE
        Set oDoc = Nothing
    Else
        Dim oNorma As New clsCa_normas
        oNorma.Carga PK
        asunto = "Impresión de Norma : " & oNorma.getCODIGO
        mensaje = vbNewLine & "Codigo : " & oNorma.getCODIGO
        mensaje = mensaje & vbNewLine & "Documento : " & oNorma.getNOMBRE
        Set oNorma = Nothing
    End If
    mensaje = mensaje & vbNewLine & "Usuario : " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
    mensaje = mensaje & vbNewLine & "Puesto : " & USUARIO.getUSO
    mensaje = mensaje & vbNewLine & "Fecha : " & Date
    mensaje = mensaje & vbNewLine & "Hora : " & Time
    mensaje = mensaje & vbNewLine & "Motivo : " & MOTIVO
    Enviar_Mail_CDO "informatica@canagrosa.com", asunto, mensaje, ""
    Enviar_Mail_CDO "calidad@canagrosa.com", asunto, mensaje, ""
'    Enviar_Mail_CDO "salvador.alarcon@canagrosa.com", ASUNTO, mensaje, ""
    Me.MousePointer = 0
    ' Si es un documento de calidad, muestro el word o excel de trabajo, sino saco el cuadro de dialogo
    If tipo = CALIDAD_VIDA_TIPOS.CALIDAD_VIDA_TIPOS_DOCUMENTO Then
        Dim documento As String
        documento = calidad_ruta_documento_trabajo(PK)
        If Dir(documento) <> "" Then
            If UCase(Right(documento, 3)) <> "DOC" Then
                r = Shell("rundll32.exe url.dll,FileProtocolHandler " & documento, vbMaximizedFocus)
            Else
                ver_documento_word documento
            End If
        Else
            MsgBox "Error al buscar el documento con ese código.", vbExclamation, App.Title
        End If
'        MsgBox "Se ha enviado el documento a la impresora principal.", vbInformation, App.Title
    Else
        pdf1.printWithDialog
    End If
    Set oVida = Nothing
    Set oDoc = Nothing

   On Error GoTo 0
   Exit Sub

cmdImprimir_Click_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmPrevisualizarPDF"
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Caption = NOMBRE
    mostrar_pdf ruta
End Sub
Private Sub mostrar_pdf(DOC As String)
'    If Dir(DOC) <> "" Then
        Dim fso As New FileSystemObject
   On Error GoTo mostrar_pdf_Error

        If fso.FileExists(DOC) Then
'            Dim cad() As String
'            Dim temp As String
'            cad = Split(DOC, "\")
'            temp = DIRECTORIO_TEMPORAL & "\" & cad(UBound(cad))
'            fso.CopyFile DOC, temp
            If USUARIO.getPER_IMPRESION_PNT = True Then
                cmdImprimir.Enabled = True
                cmdImprimir.Visible = True
                pdf1.setShowToolbar True
            Else
                pdf1.setShowToolbar False
                cmdImprimir.Visible = False
            End If
            pdf1.LoadFile DOC
        Else
            MsgBox "Error al mostrar el documento.", vbCritical, App.Title
        End If
'    End If

   On Error GoTo 0
   Exit Sub

mostrar_pdf_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mostrar_pdf of Formulario frmPrevisualizarPDF"
End Sub
Private Sub Form_Resize()
   On Error GoTo Form_Resize_Error

    pdf1.Left = 0
    pdf1.top = 0
    pdf1.Width = Me.Width - 170
    pdf1.Height = Me.Height - 1420
    frmop.Left = Me.Width - frmop.Width - 150
    frmop.top = Me.Height - frmop.Height - 405

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Formulario frmPrevisualizarPDF"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    pdf1.LoadFile vbNullString
End Sub

