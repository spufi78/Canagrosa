VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDocumento_Cobro 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del cobro"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frmDocumento_Cobro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdVer 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Factura"
      Height          =   765
      Left            =   3210
      Picture         =   "frmDocumento_Cobro.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2250
      Width           =   1605
   End
   Begin VB.CommandButton cmdObra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos de la Obra"
      Height          =   765
      Left            =   30
      Picture         =   "frmDocumento_Cobro.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2250
      Width           =   1545
   End
   Begin VB.CommandButton cmdCliente 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos del Cliente"
      Height          =   765
      Left            =   1590
      Picture         =   "frmDocumento_Cobro.frx":1B7E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2250
      Width           =   1605
   End
   Begin VB.CommandButton cmdRecibo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Recibo"
      Height          =   915
      Left            =   90
      Picture         =   "frmDocumento_Cobro.frx":2448
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6480
      Width           =   1770
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cobrar"
      Height          =   870
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6495
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   5625
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6495
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1425
      Index           =   3
      Left            =   90
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4980
      Width           =   6570
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   1020
      MaxLength       =   100
      TabIndex        =   5
      Top             =   4350
      Width           =   5655
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   4800
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   3
      Top             =   3510
      Width           =   1845
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   2970
      MaxLength       =   100
      TabIndex        =   2
      Top             =   3510
      Width           =   900
   End
   Begin MSComCtl2.DTPicker fecha 
      Height          =   330
      Left            =   1020
      TabIndex        =   1
      Top             =   3480
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   12632256
      Format          =   51838977
      CurrentDate     =   38002
   End
   Begin MSDataListLib.DataCombo cmbfp 
      Height          =   315
      Left            =   1020
      TabIndex        =   4
      Top             =   3960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lista 
      Height          =   1860
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   3281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14609914
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Datos del cobro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   0
      TabIndex        =   16
      Top             =   3060
      Width           =   6765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   15
      Top             =   4410
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   14
      Top             =   4740
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Forma Pago"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   13
      Top             =   4020
      Width           =   855
   End
   Begin VB.Label lblCampos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hora"
      Height          =   240
      Index           =   0
      Left            =   2460
      TabIndex        =   12
      Top             =   3540
      Width           =   465
   End
   Begin VB.Label lblCampos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
      Height          =   240
      Index           =   2
      Left            =   75
      TabIndex        =   11
      Top             =   3540
      Width           =   825
   End
   Begin VB.Label lblCampos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Empleado"
      Height          =   240
      Index           =   4
      Left            =   3990
      TabIndex        =   10
      Top             =   3540
      Width           =   825
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Vencimientos e importes de la factura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   -30
      TabIndex        =   9
      Top             =   30
      Width           =   6765
   End
End
Attribute VB_Name = "frmDocumento_Cobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pk As Long
Public FP_ID As Long
Private Sub cmdCliente_Click()
    Dim oDOC As New clsDocumentos
    Dim oObra As New clsObras
    oDOC.Carga pk
    oObra.Carga oDOC.getOBRA_ID

    frmClientes.pk = oObra.getCLIENTE_ID
    frmClientes.Show 1
    Set oDOC = Nothing
End Sub

Private Sub cmdObra_Click()
    Dim oDOC As New clsDocumentos
    oDOC.Carga pk
    frmObras.pk = oDOC.getOBRA_ID
    frmObras.Show 1
    Set oDOC = Nothing
End Sub
Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = False Then
        Exit Sub
    End If
    Dim ocobro As New clsDocumentos_cobros
    With ocobro
        .Eliminar_Vencimiento pk, lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        .setDOCUMENTO_ID = pk
        .setVENCIMIENTO = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        .setFECHA = Format(fecha, "yyyy-mm-dd")
        .setHORA = Format(txtDatos(0), "hh:mm:ss")
        .setFP_ID = cmbfp.BoundText
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setDATOS = txtDatos(2)
        .setOBSERVACIONES = txtDatos(3)
        If .Insertar <> 0 Then
            ' Si es recibo, marcarlo como cobrado
            Dim oDR As New clsDocumentos_Recibos
            If oDR.Carga(pk, lista.ListItems(lista.SelectedItem.Index).SubItems(1)) Then
                oDR.ESTADO pk, lista.ListItems(lista.SelectedItem.Index).SubItems(1), ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_COBRADO
            Else
                oDR.setDOCUMENTO_ID = pk
                oDR.setFECHA = Format(lista.ListItems(lista.SelectedItem.Index).SubItems(2), "yyyy-mm-dd")
                oDR.setIMPORTE = moneda_bd(lista.ListItems(lista.SelectedItem.Index).SubItems(3))
                oDR.setVENCIMIENTO = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
                oDR.setCOBRADO = ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_COBRADO
                oDR.Insertar
            End If
            
            If Not oDR.Recibos_Pendientes(pk) Or lista.ListItems.Count = 1 Then
                Dim oDOC As New clsDocumentos
                oDOC.COBRADO pk, fecha
                oDOC.modificar_estado pk, ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_COBRADA
                Set oDOC = Nothing
            End If
            
            Set oDR = Nothing
        End If
    End With
    MsgBox "Los datos del cobro se han insertado correctamente.", vbInformation, App.Title
    Unload Me

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmDocumento_Cobro"
End Sub

Private Sub cmdRecibo_Click()
   On Error GoTo cmdRecibo_Click_Error

    If lista.ListItems.Count > 0 Then
        Dim datos As String
        
        Dim oDOC As New clsDocumentos
        Dim oObra As New clsObras
        Dim ocliente As New clsCliente
        Dim oProvincia As New clsProvincias
        Dim oMunicipio As New clsMunicipios
        Dim tNum2Text As New cNum2Text
        
        oDOC.Carga pk
        oObra.Carga oDOC.getOBRA_ID
        ocliente.CargaCliente oObra.getCLIENTE_ID
        oProvincia.Carga ocliente.getPROVINCIA_ID
        oMunicipio.Cargar ocliente.getMUNICIPIO_ID
       
        datos = datos & "<RECIBO>"
        datos = datos & "  <NUMERO_RECIBO>" & Format(oDOC.getNUMERO, "0000") & "/" & oDOC.getANNO & "</NUMERO_RECIBO>"
        datos = datos & "  <LOCALIDAD>ARCOS DE LA FRONTERA</LOCALIDAD>"
        datos = datos & "  <IMPORTE>" & lista.ListItems(lista.SelectedItem.Index).SubItems(3) & "</IMPORTE>"
        datos = datos & "  <FECHA>" & Format(oDOC.getFECHA, "dd-mm-yyyy") & "</FECHA>"
        datos = datos & "  <VENCIMIENTO>" & Format(lista.ListItems(lista.SelectedItem.Index).SubItems(2), "dd-mm-yyyy") & "</VENCIMIENTO>"
        datos = datos & "  <IMPORTE_LETRAS>" & UCase(tNum2Text.Numero2Letra(lista.ListItems(lista.SelectedItem.Index).SubItems(3), , 2, "euro", "céntimo", Masculino, Masculino)) & "</IMPORTE_LETRAS>"
        datos = datos & "  <BANCO>" & oObra.getBANCO & "</BANCO>"
        datos = datos & "  <OFICINA>" & oObra.getBANCO_DIRECCION & "</OFICINA>"
        datos = datos & "  <CCC>" & oObra.getCCC & "</CCC>"
        datos = datos & "  <CLIENTE_NOMBRE>" & ocliente.getNOMBRE & "</CLIENTE_NOMBRE>"
        datos = datos & "  <CLIENTE_DIRECCION>" & ocliente.getDIRECCION & "</CLIENTE_DIRECCION>"
        datos = datos & "  <CLIENTE_LOCALIDAD>" & ocliente.getCP & " " & oMunicipio.getNOMBRE & "</CLIENTE_LOCALIDAD>"
        datos = datos & "  <CLIENTE_PROVINCIA>" & oProvincia.getNOMBRE & "</CLIENTE_PROVINCIA>"
        datos = datos & "</RECIBO>"
        
        
        Dim tx As TextStream
        Dim gFSO As New Scripting.FileSystemObject
        
        Set tx = gFSO.CreateTextFile(ReadINI(App.Path + "\config.ini", "documentos", "informes") & "\recibo.xml", True, True)
        
        tx.WriteLine Replace("<?xml version='1.0' encoding='UTF-16'?>", "'", Chr(34))
        tx.WriteLine datos
        tx.Close
        
        With frmReport
            .iniciar
            .informe = "rptrecibo"
            .consulta = ""
            .imprimir = False
            .pdf = ""
            .xml = ReadINI(App.Path + "\config.ini", "documentos", "informes") & "\recibo.xml"
            .generar
            .Show 1
        End With
        Unload frmReport
    End If

   On Error GoTo 0
   Exit Sub

cmdRecibo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRecibo_Click of Formulario frmDocumento_Cobro"

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVer_Click()
    frmDocumento.PK_DOCUMENTO = pk
    frmDocumento.Show 1
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    cabecera_lista
            
    If pk > 0 Then
        cargar_datos
    End If
End Sub

Private Sub cargar_combos()
    Cargar_Combo cmbfp, New clsForma_pago
End Sub
Private Sub cabecera_lista()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Vencimiento", 1200, lvwColumnCenter
        .Add , , "Fecha", 1400, lvwColumnCenter
        .Add , , "Importe", 1600, lvwColumnRight
        .Add , , "Estado", 2200, lvwColumnCenter
    End With
End Sub

Public Sub cargar_datos()
    Dim oDR As New clsDocumentos_Recibos
    Dim oDOC As New clsDocumentos
   On Error GoTo cargar_datos_Error

    oDOC.Carga pk
    lbltitulo = "Vencimientos e importes de la factura: " & Format(oDOC.getNUMERO, "0000") & "/" & oDOC.getANNO
    FP_ID = oDOC.getFP_ID
    Dim rs As ADODB.Recordset
    Set rs = oDR.Listado(pk)
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = Format(rs(2), "dd-mm-yyyy")
                If oDOC.getIVA = 0 Then
                    .SubItems(3) = moneda(rs(3))
                Else
                    .SubItems(3) = moneda(rs(3) + ((rs(3) * oDOC.getIVA) / 100))
                End If
                If rs(4) = 0 Then
                    .SubItems(4) = "Pendiente"
                ElseIf (rs(4)) = 1 Then
                    .SubItems(4) = "Cobrado"
                ElseIf (rs(4)) = 2 Then
                    .SubItems(4) = "Remesa"
                
                ElseIf (rs(4)) = 3 Then
                    .SubItems(4) = "Descuento"

                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    Else
        With lista.ListItems.Add(, , pk)
            .SubItems(1) = 1
            .SubItems(2) = Format(oDOC.getFECHA, "dd-mm-yyyy")
            .SubItems(3) = moneda(oDOC.getTOTAL + (oDOC.getTOTAL * oDOC.getIVA / 100))
            .SubItems(4) = "Pendiente"
        End With
    End If
        Dim ocobro As New clsDocumentos_cobros
        If ocobro.Carga(pk, 1) = True Then
            With ocobro
                fecha = .getFECHA
                txtDatos(0) = Format(.getHORA, "hh:mm:ss")
                Dim oEmpleado As New ClsUsuario
                oEmpleado.Cargar (.getEMPLEADO_ID)
                txtDatos(1) = oEmpleado.getNOMBRE
                txtDatos(2) = .getDATOS
                txtDatos(3) = .getOBSERVACIONES
                cmbfp.BoundText = .getFP_ID
            End With
'            cmdok.Visible = False
'            txtDatos(0).Locked = True
'            txtDatos(3).Locked = True
'            txtDatos(2).Locked = True
'            fecha.Enabled = False
'            cmbfp.Locked = True
        End If
'    End If
    lista_Click

   On Error GoTo 0
   Exit Sub

cargar_datos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_datos of Formulario frmDocumento_Cobro"
End Sub

Public Function validar() As Boolean
    validar = True
    If cmbfp.Text = "" Then
        validar = False
        MsgBox "Introduzca la forma de pago.", vbExclamation, App.Title
    End If
End Function

Private Sub lista_Click()
    Dim oDC As New clsDocumentos_cobros
    If oDC.Carga(pk, lista.ListItems(lista.SelectedItem.Index).SubItems(1)) = True Then
        fecha = oDC.getFECHA
        txtDatos(0) = Format(oDC.getHORA, "hh:mm:ss")
        Dim oEmpleado As New ClsUsuario
        oEmpleado.Cargar oDC.getEMPLEADO_ID
        txtDatos(1) = oEmpleado.getUSUARIO
        txtDatos(2) = oDC.getDATOS
        txtDatos(3) = oDC.getOBSERVACIONES
    Else
        fecha = Date
        cmbfp.BoundText = FP_ID
        txtDatos(0) = Format(Time, "hh:mm:ss")
        txtDatos(1) = USUARIO.getUSUARIO
        txtDatos(2) = ""
        txtDatos(3) = ""
    End If
    Set oDC = Nothing
End Sub
Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &HC0FFFF
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
