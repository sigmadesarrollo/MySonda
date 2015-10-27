VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmPagoAjustes 
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8565
   Icon            =   "frmPagoAjustes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   8565
   Begin MSCommLib.MSComm Com 
      Left            =   600
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox txtPago 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "&Salir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmPagoAjustes.frx":030A
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
   End
   Begin vbAcceleratorGrid6.vbalGrid grdUsuarios 
      Height          =   2895
      Left            =   3120
      TabIndex        =   12
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5106
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
      DefaultRowHeight=   17
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Total a Pagar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "Abonos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Ultimo Saldo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Pago:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "Saldo a Pagar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1545
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<Saldo>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   2640
      Width           =   945
   End
   Begin VB.Label lblUltimoSaldo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<Saldo>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label lblAbonos 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<Abonos>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<Total>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "frmPagoAjustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Grabar_Datos
Abrir_Cajon
limpiar
Cargar_Usuarios
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim ctrl As Control
Me.Height = 4320
Me.Width = 8685
CentrarForm frmPagoAjustes, frmMDI
Checar_Saldos
For Each ctrl In Controls
    If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = ""
Next
Crear_Encabezados
Cargar_Usuarios
End Sub

Private Sub Checar_Saldos()
dbDatos.Execute "UPDATE Usuarios SET Cargo=0,Abono=0 WHERE Cargo=Abono"
dbDatos.Execute "UPDATE Usuarios SET Prestamo=0,Abonoprestamo=0 WHERE prestamo=Abonoprestamo"
End Sub

Private Sub grdUsuarios_Click(ByVal lRow As Long, ByVal lCol As Long)
Dim rcUsuarios As New ADODB.Recordset

rcUsuarios.Open "SELECT * FROM Usuarios WHERE id=" & grdUsuarios.CellItemData(grdUsuarios.SelectedRow, 1) & "", dbDatos, adOpenDynamic, adLockOptimistic

With rcUsuarios
   If frmPagoAjustes.Tag = 1 Then
      lblTotal.Caption = str(Format(!Cargo, "###,###,##0.00"))
      lblAbonos.Caption = str(Format(!abono, "###,###,##0.00"))
      lblUltimoSaldo.Caption = str(Format(!Cargo - !abono, "###,###,##0.00"))
   Else
      lblTotal.Caption = str(Format(!prestamo, "###,###,##0.00"))
      lblAbonos.Caption = str(Format(!abonoprestamo, "###,###,##0.00"))
      lblUltimoSaldo.Caption = str(Format(!prestamo - !abonoprestamo, "###,###,##0.00"))
   End If
    .Close
End With
If Val(lblUltimoSaldo.Caption) > 0 Then
    cmdAceptar.Enabled = True
End If

End Sub

Private Sub txtPago_Change()
lblSaldo.Caption = Format(CCur(lblUltimoSaldo.Caption) - CCur(IIf(txtPago.Text = "", 0, txtPago.Text)), "###,###,##0.00")
End Sub

Private Sub txtPago_KeyPress(KeyAscii As Integer)
KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub Grabar_Datos()
Dim rcUsuarios As New ADODB.Recordset

rcUsuarios.Open "SELECT * FROM Usuarios WHERE ID=" & grdUsuarios.CellItemData(grdUsuarios.SelectedRow, 1) & "", dbDatos, adOpenDynamic, adLockOptimistic

With rcUsuarios
   If frmPagoAjustes.Tag = 1 Then
      !abono = !abono + CCur(txtPago.Text)
   Else
      !abonoprestamo = !abonoprestamo + CCur(txtPago.Text)
   End If
    .Update
End With

dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC) VALUES " & _
                  "(#" & Format(Date, "MM/DD/YY") & "#,0,0,'" & IIf(frmPagoAjustes.Tag = 1, "AJ01", "PR01") & "','110101'," & CCur(txtPago.Text) & "," & TIPO_CARGO & ",0,'" & IIf(frmPagoAjustes.Tag = 1, "Pago Ajuste ", "Pago Prestamo ") & "'+'" & grdUsuarios.CellText(grdUsuarios.SelectedRow, 1) & "','" & NombrePc & "')"


dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC) VALUES " & _
                  "(#" & Format(Date, "MM/DD/YY") & "#,0,0,'" & IIf(frmPagoAjustes.Tag = 1, "AJ01", "PR01") & "','199401'," & CCur(txtPago.Text) & "," & TIPO_CARGO & ",0,'" & IIf(frmPagoAjustes.Tag = 1, "Pago Ajuste ", "Pago Prestamo ") & "'+'" & grdUsuarios.CellText(grdUsuarios.SelectedRow, 1) & "','" & NombrePc & "')"


dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC) VALUES " & _
                  "(#" & Format(Date, "MM/DD/YY") & "#,0,0,'" & IIf(frmPagoAjustes.Tag = 1, "AJ50", "PR50") & "','" & IIf(frmPagoAjustes.Tag = 1, "151250", "151350") & "'," & CCur(txtPago.Text) & "," & TIPO_ABONO & ",0,'" & IIf(frmPagoAjustes.Tag = 1, "Pago Ajuste ", "Pago Prestamo ") & "'+'" & grdUsuarios.CellText(grdUsuarios.SelectedRow, 1) & "','" & NombrePc & "')"

End Sub

'Funcion para abrir el cajon mediante el puerto serial(COM1)
Private Sub Abrir_Cajon()
On Error GoTo error

    'Cierra el puerto para permitir nuevos parametros
   If Com.PortOpen Then
        Com.PortOpen = False
    End If
    
    'Puerto que sera usado
    Com.CommPort = 1
    'Baudios, paridad, datos, detener
    Com.Settings = "9600,N,8,1"
    
    'Activa el puerto COM
    Com.PortOpen = True
    
    'Texto de salida para el puerto
    Com.Output = "U"
        
error:
   Maneja_Error Err

End Sub

Private Sub limpiar()
Dim ctrl As Control

For Each ctrl In Controls
    If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = ""
    If TypeOf ctrl Is TextBox Then ctrl.Text = ""
    If TypeOf ctrl Is vbalGrid Then ctrl.Clear
Next
cmdAceptar.Enabled = False
End Sub

'Creamos los encabezados para la lista
Private Sub Crear_Encabezados()
   With grdUsuarios
      .ImageList = frmMDI.img
      .AddColumn "K1", "Usuarios", ecgHdrTextALignLeft, , 165, , , , , , , CCLSortString
      .AddColumn "K2", "Prestamo", ecgHdrTextALignLeft, , 60, , , , , , , CCLSortString
      .AddColumn "K3", "Abono", ecgHdrTextALignLeft, , 60, , , , , , , CCLSortString
      .AddColumn "K4", "Saldo", ecgHdrTextALignLeft, , 60, , , , , , , CCLSortString
   End With
End Sub

'Cargamos los usuarios en el grid
Private Sub Cargar_Usuarios()
   On Error GoTo error
   'Dim dbCatalogos As New ADODB.Connection
   Dim rcUsuarios As New ADODB.Recordset
   
   'dbCatalogos.Open CONEXION & Path & "\Base De Datos\Datos.mdb" & USUARIO
   rcUsuarios.Open "SELECT * FROM Usuarios", dbDatos, adOpenDynamic, adLockOptimistic
   grdUsuarios.Redraw = False
   With rcUsuarios
      While Not .EOF
         grdUsuarios.AddRow
         grdUsuarios.CellText(grdUsuarios.Rows, 1) = !Nombre
         grdUsuarios.CellItemData(grdUsuarios.Rows, 1) = !ID
         grdUsuarios.CellText(grdUsuarios.Rows, 2) = !prestamo
         grdUsuarios.CellText(grdUsuarios.Rows, 3) = !abonoprestamo
         grdUsuarios.CellText(grdUsuarios.Rows, 4) = !prestamo - !abonoprestamo
         .MoveNext
      Wend
   End With
   grdUsuarios.Redraw = True
   rcUsuarios.Close
   'dbCatalogos.Close
   
error:
   'Verificamos si hay error
   Maneja_Error Err
   
   Set rcUsuarios = Nothing
   'Set dbCatalogos = Nothing
End Sub

