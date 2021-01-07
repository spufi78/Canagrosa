VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDecodificadora 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decodificadora"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   Icon            =   "frmDecodificadora.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   11895
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   1050
      TabIndex        =   0
      Top             =   6075
      Width           =   2895
   End
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   870
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   7110
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   1050
      TabIndex        =   2
      Top             =   6705
      Width           =   10770
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7110
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7110
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7110
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10755
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7110
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   1050
      TabIndex        =   1
      Top             =   6390
      Width           =   10770
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5340
      Left            =   15
      TabIndex        =   8
      Top             =   720
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   9419
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
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Valor"
      Height          =   240
      Left            =   90
      TabIndex        =   13
      Top             =   6120
      Width           =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Parámetros"
      Height          =   240
      Left            =   90
      TabIndex        =   12
      Top             =   6750
      Width           =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción"
      Height          =   240
      Left            =   90
      TabIndex        =   11
      Top             =   6435
      Width           =   960
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Descripción de campos decodificados"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   360
      Width           =   2700
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11250
      Picture         =   "frmDecodificadora.frx":08CA
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Decodificadora"
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
      Left            =   90
      TabIndex        =   9
      Top             =   45
      Width           =   1620
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmDecodificadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CODIGO As Long
Private Sub cmdAdjuntos_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    'M0499-I
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_DECODIFICADORA
        .CODIGO_DECODIFICADORA = CODIGO
        .COBJETO = lista.ListItems(lista.selectedItem.Index).Text
        .Show 1
    End With
    Set frmAdjuntos = Nothing
'    frmMuestras_Adjuntos.Inicializar
'    frmMuestras_Adjuntos.tipo = ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_DECODIFICADORA
'    frmMuestras_Adjuntos.PK_DECODIFICADORA = CODIGO
'    frmMuestras_Adjuntos.PK_DECODIFICADORA_VALOR = lista.ListItems(lista.selectedItem.Index).Text
'    frmMuestras_Adjuntos.Show 1
    'M0499-F
End Sub
Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a insertar el registro. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oDecodificadora As New clsDecodificadora
            With oDecodificadora
                .setCODIGO = CODIGO
                .setIDIOMA = "ES"
                If txtDatos(2) = "" Then
                    .setVALOR = 0
                Else
                    .setVALOR = txtDatos(2)
                End If
                .setDESCRIPCION = txtDatos(0)
                .setPARAMETROS = txtDatos(1)
                .insertar
            End With
 'M1089-I
 'En caso de alta de un parámetro de periodicidad de equipos debemos registrar los valores
 ' en una tabla paralela llamada eq_periodicidad
 
            If oDecodificadora.getCODIGO = DECODIFICADORA.EQ_periodicidad Then
                Dim oPeriodicidad As New clsEquiposPeriodicidad
                oPeriodicidad.setID_PERIODICIDAD = oDecodificadora.getVALOR
                oPeriodicidad.setDESCRIPCION = txtDatos(0)
 
                'Obtención de los días y días de preaviso
                                
                Dim strParametros() As String
                Dim intCount As Integer
                Dim VALOR As Integer
                               
                strParametros = Split(txtDatos(1), ";")
                
                For intCount = LBound(strParametros) To UBound(strParametros)
                    
                    If strParametros(intCount) <> "" Then 'Para prevenir el caso de encontrar un ; al final de la línea de parámetros
                    
                      VALOR = CInt(Solo_Numeros(strParametros(intCount)))
                      
                      'intcount: número de parámetros
                      Select Case intCount
                      Case 0 'Primer parámetro: dias
                         oPeriodicidad.setDIAS = VALOR
                      Case 1 'Segundo parámetro: Dias preaviso
                         oPeriodicidad.setDIAS_PREAVISO = VALOR
                      End Select
                      
                    End If
                    
                Next intCount
                oPeriodicidad.setMESES = 0
                oPeriodicidad.insertar
                
                Set oPeriodicidad = Nothing
            End If
 'M1089-F
            
            cargar_lista
        End If
    End If
    txtDatos(0).SetFocus

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmDecodificadora"
End Sub

Private Sub cmdcancel_Click()
    If CODIGO = DECODIFICADORA.EMPLEADOS_DEPARTAMENTOS Then
        frmEmpleados_Categorias.cargar_categorias
    End If
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR el registro. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oDecodificadora As New clsDecodificadora
            If oDecodificadora.Eliminar(CODIGO, lista.ListItems(lista.selectedItem.Index).Text) = True Then
                'M1089-I
                'En caso de borrado de un parámetro de periodicidad de equipos debemos borrar los mismos valores
                ' en una tabla paralela llamada eq_periodicidad
 
                If CODIGO = DECODIFICADORA.EQ_periodicidad Then
                    Dim oPeriodicidad As New clsEquiposPeriodicidad
                    oPeriodicidad.Eliminar CLng(lista.ListItems(lista.selectedItem.Index).Text)

                    Set oPeriodicidad = Nothing
                End If
                'M1089-F
                
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_Error

    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a modificar el registro. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oDecodificadora As New clsDecodificadora
            With oDecodificadora
                .setIDIOMA = "ES"
                .setDESCRIPCION = txtDatos(0)
                .setPARAMETROS = txtDatos(1)
                .modificar CODIGO, lista.ListItems(lista.selectedItem.Index).Text
            End With
            
            'M1089-I
            'En caso de modificación de un parámetro de periodicidad de equipos debemos actualizar los valores
            ' en una tabla paralela llamada eq_periodicidad
 
            If CODIGO = DECODIFICADORA.EQ_periodicidad Then
                Dim oPeriodicidad As New clsEquiposPeriodicidad
                
                oPeriodicidad.setDESCRIPCION = txtDatos(0)
 
                'Obtención de los días y días de preaviso
                                
                Dim strParametros() As String
                Dim intCount As Integer
                Dim VALOR As Integer
                               
                strParametros = Split(txtDatos(1), ";")
                
                For intCount = LBound(strParametros) To UBound(strParametros)
                    
                    If strParametros(intCount) <> "" Then 'Para prevenir el caso de encontrar un ; al final de la línea de parámetros
                    
                      VALOR = CInt(Solo_Numeros(strParametros(intCount)))
                      
                      'intcount: número de parámetros
                      Select Case intCount
                      Case 0 'Primer parámetro: dias
                         oPeriodicidad.setDIAS = VALOR
                      Case 1 'Segundo parámetro: Dias preaviso
                         oPeriodicidad.setDIAS_PREAVISO = VALOR
                      End Select
                      
                    End If
                    
                Next intCount
                oPeriodicidad.setMESES = 0
                oPeriodicidad.modificar CLng(lista.ListItems(lista.selectedItem.Index).Text)
                
                Set oPeriodicidad = Nothing
            End If
            'M1089-F
            
            cargar_lista
            txtDatos(0) = ""
            txtDatos(1) = ""
        End If
    End If
    txtDatos(0).SetFocus

   On Error GoTo 0
   Exit Sub

cmdModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmDecodificadora"
End Sub

Private Sub Form_Activate()
'    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 300
    Me.top = 300
    With lista.ColumnHeaders
        .Add , , "Valor", 800, lvwColumnLeft
        .Add , , "Descripción", 5400, lvwColumnLeft
        .Add , , "Parámetros", 5400, lvwColumnLeft
    End With
    If CODIGO <> 0 Then
        Dim oDecodificadora As New clsDecodificadora
        oDecodificadora.Carga CODIGO
        lbltitulo = oDecodificadora.getDESCRIPCION
        Me.Caption = lbltitulo
    End If
    If CODIGO = DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS Or CODIGO = DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_PAGOS Then
        txtDatos(2).BackColor = vbWhite
        txtDatos(2).Enabled = True
    End If
    If CODIGO = DECODIFICADORA.CA_NORMAS_SUBTIPOS Then
        lista.ColumnHeaders.Item(3).Text = "Correo destino normas"
        Label2.Caption = "Correo"
    End If
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oDecodificadora As New clsDecodificadora
    Set rs = oDecodificadora.Listado(CODIGO)
    txtDatos(0) = ""
    txtDatos(1) = ""
    txtDatos(2) = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("valor"))
            .SubItems(1) = rs("descripcion")
            .SubItems(2) = rs("parametros")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oDecodificadora = Nothing
    Set rs = Nothing
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

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtDatos(0).Text = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        txtDatos(1).Text = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        txtDatos(2).Text = lista.ListItems(lista.selectedItem.Index).Text
    End If
End Sub

'M1089-I
'Función para obtener los números de izquierda a derecha de una cadena de texto
'Ejemplo: cadena="dias_preaviso=15" -> Solo_Numeros(cadena) = "15"

'Public Function Solo_Numeros(ByRef sText As String) As String
'    Dim sActualChar                 As String * 1
'    Dim lTotalChar                  As Long
'    Dim x                           As Long
'
'    lTotalChar = LenB(sText) \ 2
'
'    If CBool(lTotalChar) Then
'        For x = 1 To lTotalChar
'            sActualChar = Mid$(sText, x, 1)
'            If IsNumeric(sActualChar) Then Solo_Numeros = Solo_Numeros & sActualChar
'        Next
'    End If
'
'End Function

'M1089-F
