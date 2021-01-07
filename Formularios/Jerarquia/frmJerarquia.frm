VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmJerarquia 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jerarquia por Tipos de Muestra"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "frmJerarquia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   11760
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Generar Jerarquia para :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   4815
      TabIndex        =   4
      Top             =   7785
      Width           =   2985
      Begin VB.OptionButton opOpcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baños"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   540
         Width           =   2580
      End
      Begin VB.OptionButton opOpcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipos de Muestra"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   2580
      End
   End
   Begin VB.CommandButton cmdVer 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Objeto Seleccionado"
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
      Left            =   90
      Picture         =   "frmJerarquia.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7785
      Width           =   2805
   End
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
      Left            =   10665
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7785
      Width           =   1050
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3150
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJerarquia.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJerarquia.frx":2416
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJerarquia.frx":2EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJerarquia.frx":37BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJerarquia.frx":3AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJerarquia.frx":43AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJerarquia.frx":4C88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   7410
      Left            =   60
      TabIndex        =   0
      Top             =   345
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13070
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Jerarquia por Tipos de Muestra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   11625
   End
End
Attribute VB_Name = "frmJerarquia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdVer_Click()
    If Tree.Nodes.Count > 0 Then
       Dim clave As Variant
       Dim nodX As Node
       clave = Split(Tree.Nodes(Tree.SelectedItem.Index).Key, ".")
       Select Case clave(0)
        Case "TM"
            frmTM_Detalle.PK = clave(1)
            frmTM_Detalle.Show 1
        Case "TA"
            frmTA_Detalle.PK = clave(1)
            frmTA_Detalle.Show 1
        Case "TD"
            frmTD_Detalle.PK = clave(1)
            frmTD_Detalle.Show 1
        Case "FO"
            frmFORMULA_Detalle.PK = clave(1)
            frmFORMULA_Detalle.Show 1
        Case "BA"
            frmBANO_Detalle.PK = clave(1)
            frmBANO_Detalle.Show 1
       End Select
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 150
    Me.Top = 50
    cargar_tipos_muestra
End Sub

Private Sub opOpcion_Click(Index As Integer)
    Select Case Index
        Case 0
            cargar_tipos_muestra
        Case 1
            cargar_banos
    End Select
End Sub

Private Sub Tree_DblClick()
   On Error GoTo Tree_DblClick_Error

    If Tree.Nodes.Count > 0 Then
        If Tree.Nodes(Tree.SelectedItem.Index).Children = 0 Then
            ' clave (0) = "TM"
            ' clave (1) = ID
            Dim clave As Variant
            Dim nodX As Node
            clave = Split(Tree.Nodes(Tree.SelectedItem.Index).Key, ".")
            Dim rs As ADODB.RecordSet
            Select Case clave(0)
                Case "TM"
                    Dim oTipos_analisis As New clsTipos_analisis
                    Set rs = oTipos_analisis.lista_tipo_muestra(CInt(clave(1)), "", False)
                    If rs.RecordCount > 0 Then
                        Do
                            Set nodX = Tree.Nodes.Add(Tree.Nodes(Tree.SelectedItem.Index).Key, tvwChild, "TA." & rs(4), rs(0) & " (" & rs(4) & ")", 7)
                            rs.MoveNext
                        Loop Until rs.EOF
                        nodX.EnsureVisible
                    End If
                Case "TA"
                    Dim oTipos_Determinacion As New clsTipos_determinacion
                    Set rs = oTipos_Determinacion.DeterminacionesPorDefecto(CLng(clave(1)))
                    'Debug.Print "Comienzo"
                    If rs.RecordCount > 0 Then
                        Do
                            Set nodX = Tree.Nodes.Add(Tree.Nodes(Tree.SelectedItem.Index).Key, tvwChild, "TD." & rs(0) & "." & clave(1) & "." & rs(2), rs(2) & " (" & rs(0) & ")", 3)
                            rs.MoveNext
                        Loop Until rs.EOF
                        nodX.EnsureVisible
                    End If
                Case "TD" ' Determinacion, carga formula y campos
                    Dim oTipos_Determinacion2 As New clsTipos_determinacion
                    oTipos_Determinacion2.CargarTipoDeterminacion (CInt(clave(1)))
                    Dim oFormula As New clsFormulas
                    oFormula.CARGAR (oTipos_Determinacion2.getFORMULA_ID)
                    Dim ClaveFormula As String
                    ClaveFormula = "FO." & oFormula.getID_FORMULA & "." & clave(1) & "." & Format(Time, "hh:mm:ss")
                    Set nodX = Tree.Nodes.Add(Tree.Nodes(Tree.SelectedItem.Index).Key, tvwChild, ClaveFormula, "Formula : " & oFormula.getNOMBRE & " (" & oFormula.getID_FORMULA & ")", 4)
                    Dim oformulas_campos As New clsFormulas_campos
                    Set rs = oformulas_campos.Lista_Formulas_Unidades(oTipos_Determinacion2.getFORMULA_ID)
                    If rs.RecordCount > 0 Then
                        Do
                            If rs(2) = 0 Then
                                Set nodX = Tree.Nodes.Add(ClaveFormula, tvwChild, "CA." & rs(0) & "." & rs(2) & "." & Format(Time, "hh:mm:ss"), "Campo : " & rs(1) & " " & rs(3), 5)
                            Else
                                Set nodX = Tree.Nodes.Add(ClaveFormula, tvwChild, "CA." & rs(0) & "." & rs(2) & "." & Format(Time, "hh:mm:ss"), "Campo Solución : " & rs(1) & " " & rs(3), 5)
                            End If
                            rs.MoveNext
                        Loop Until rs.EOF
                    End If
                    nodX.EnsureVisible
                Case "BA"
                    Dim oTipos_Determinacion3 As New clsTipos_determinacion
                    Set rs = oTipos_Determinacion3.DeterminacionesBano(CLng(clave(1)))
                    If rs.RecordCount > 0 Then
                        Do
                            Set nodX = Tree.Nodes.Add(Tree.Nodes(Tree.SelectedItem.Index).Key, tvwChild, "TD." & rs(0) & "." & clave(1) & "." & rs(2), rs(2) & " (" & rs(0) & ")", 3)
                            rs.MoveNext
                        Loop Until rs.EOF
                        nodX.EnsureVisible
                    End If
                    
            End Select
        End If
    End If

   On Error GoTo 0
   Exit Sub

Tree_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Tree_DblClick of Formulario frmJerarquia"
End Sub
Public Sub cargar_tipos_muestra()
    Dim nodX As Node
   On Error GoTo cargar_tipos_muestra_Error
    lbltitulo = "Jerarquia por Tipos de Muestras"
    Tree.Nodes.Clear
    Dim oTipos_muestra As New clsTipos_muestra
    Dim rs As ADODB.RecordSet
    Set rs = oTipos_muestra.Listado_sin_anular
    If rs.RecordCount > 0 Then
        Do
            Set nodX = Tree.Nodes.Add(, , "TM" & "." & rs(0) & "." & rs(1), rs(1) & " (" & rs(0) & ")", 6)
            rs.MoveNext
        Loop Until rs.EOF
    End If

   On Error GoTo 0
   Exit Sub

cargar_tipos_muestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_tipos_muestra of Formulario frmJerarquia"
End Sub
Public Sub cargar_banos()
    Dim nodX As Node
   On Error GoTo cargar_banos_Error
    lbltitulo = "Jerarquia por Baños"
    Tree.Nodes.Clear
    Dim obanos As New clsBanos
    Dim rs As ADODB.RecordSet
    Set rs = obanos.Listado
    If rs.RecordCount > 0 Then
        Do
            Set nodX = Tree.Nodes.Add(, , "BA" & "." & rs(0) & "." & rs(1), rs(1) & "   --> Cliente : " & rs(2) & " (" & rs(0) & ")", 2)
            rs.MoveNext
        Loop Until rs.EOF
    End If

   On Error GoTo 0
   Exit Sub

cargar_banos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_tipos_muestra of Formulario frmJerarquia"
End Sub

