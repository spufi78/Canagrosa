VERSION 5.00
Begin VB.Form frmVidaMuestra 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Vida de la Muestra"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   Icon            =   "frmVidaMuestra.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   8580
      Picture         =   "frmVidaMuestra.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1050
   End
   Begin VB.TextBox texto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6270
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   405
      Width           =   9570
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informe de vida de la Muestra"
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
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   9570
   End
End
Attribute VB_Name = "frmVidaMuestra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo Form_Load_Error
    log (Me.Name)
    If gmuestra <> 0 Then
        Dim omuestra As New clsMuestra
        omuestra.CargaMuestra (gmuestra)
        ' Recepción de la muestra
        Dim ousu As New clsUsuarios
        ousu.cargar (omuestra.getEMPLEADO_ID)
        texto = texto & "------------------------------------------------------------------" & vbNewLine
        texto = texto & "               Datos de la recepción de la muestra      " & vbNewLine
        texto = texto & "------------------------------------------------------------------" & vbNewLine
        texto = texto & "Recepcionada por : " & ousu.getNOMBRE & " el " & Format(omuestra.getFECHA_RECEPCION, "dd-mm-yyyy") & " a las " & omuestra.getHORA_RECEPCION & vbNewLine
        texto = texto & vbNewLine
        texto = texto & "------------------------------------------------------------------" & vbNewLine
        texto = texto & "                   Análisis de determinaciones        " & vbNewLine
        texto = texto & "------------------------------------------------------------------" & vbNewLine
        Dim odeter As New clsDeterminaciones
        Dim oTD As New clsTipos_determinacion
        Dim rs As New ADODB.Recordset
        Set rs = odeter.lista_determinaciones(gmuestra)
        If rs.RecordCount <> 0 Then
            Do
                odeter.CargarDeterminacion (rs("id_Determinacion"))
                oTD.CargarTipoDeterminacion (rs("id_tipo_determinacion"))
                If Trim(odeter.getRESULTADO) <> "" Then
                    ousu.cargar (odeter.getEMPLEADO_ID)
                    texto = texto & oTD.getNOMBRE & " (RESULTADO " & Trim(odeter.getRESULTADO) & ")  analizada por " & ousu.getNOMBRE & " el " & Format(odeter.getFECHA, "dd-mm-yyyy") & " a las " & Format(odeter.getHORA, "hh:mm") & vbNewLine
                Else
                    texto = texto & oTD.getNOMBRE & " sin analizar." & vbNewLine
                End If
                rs.MoveNext
            Loop Until rs.EOF
        End If
        texto = texto & vbNewLine & vbNewLine
        If omuestra.getFECHA_CIERRE = "" Then
            texto = texto & "MUESTRA ABIERTA."
        Else
            texto = texto & "CERRADA EL : " & Format(omuestra.getFECHA_CIERRE, "dd-mm-yyyy")
        End If
        texto = texto & vbNewLine & vbNewLine
        texto = texto & "------------------------------------------------------------------" & vbNewLine
        texto = texto & "                             INFORMES                  " & vbNewLine
        texto = texto & "------------------------------------------------------------------" & vbNewLine
        If omuestra.getULT_EDICION_IMP = 0 Then
            texto = texto & "La muestra no tiene ediciones impresas."
        Else
            texto = texto & "Ultima edición impresa : " & omuestra.getULT_EDICION_IMP & vbNewLine
            If omuestra.getULT_EDICION_IMP > 1 Then
                Dim consulta As String
                consulta = "select EDICION,OBSERVACIONES,USUARIO,FECHA from muestras_nueva_edicion where muestra_id = " & gmuestra & " order by edicion"
'                Dim rs As ADODB.Recordset
                Set rs = datos_bd(consulta)
                If rs.RecordCount > 0 Then
                    Do
                        texto = texto & "Edicion nº " & rs(0) & " generada el " & Format(rs(3), "dd-mm-yyyy") & " por el usuario " & rs(2) & vbNewLine
                        texto = texto & "   MOTIVO : " & rs(1) & vbNewLine
                        rs.MoveNext
                    Loop Until rs.EOF
                End If
            End If
        End If
    End If
   On Error GoTo 0
   Exit Sub
Form_Load_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmVidaMuestra"
End Sub

