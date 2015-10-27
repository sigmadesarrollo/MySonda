VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmCambios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambios"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCambios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   8925
   Begin VB.TextBox txtNombrePago 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.TextBox txtCodigo2 
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
      Left            =   4560
      MaxLength       =   8
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtCodigo1 
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
      Left            =   120
      MaxLength       =   7
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   6240
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
      Picture         =   "frmCambios.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   6240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "&Aceptar"
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
   Begin DevPowerFlatBttn.FlatBttn cmdMosClave 
      Height          =   285
      Left            =   5400
      TabIndex        =   11
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      AutoSize        =   0   'False
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
   End
   Begin vbAcceleratorGrid6.vbalGrid grdCambiar 
      Height          =   2055
      Left            =   120
      TabIndex        =   26
      Top             =   2880
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3625
      RowMode         =   -1  'True
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
   Begin vbAcceleratorGrid6.vbalGrid grdNueva 
      Height          =   2055
      Left            =   4560
      TabIndex        =   27
      Top             =   2880
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3625
      RowMode         =   -1  'True
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
   Begin MSCommLib.MSComm Com 
      Left            =   360
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Folio:"
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
      Left            =   6915
      TabIndex        =   32
      Top             =   0
      Width           =   600
   End
   Begin VB.Label lblVencimiento 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<Ven>"
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
      Left            =   7935
      TabIndex        =   31
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento:"
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
      Left            =   6105
      TabIndex        =   30
      Top             =   480
      Width           =   1395
   End
   Begin VB.Label lblFolio 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<Folio>"
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
      Left            =   7830
      TabIndex        =   29
      Top             =   0
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
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
      TabIndex        =   28
      Top             =   5160
      Width           =   630
   End
   Begin VB.Label lblTotal2 
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
      Left            =   6360
      TabIndex        =   25
      Top             =   5160
      Width           =   900
   End
   Begin VB.Label Label4 
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
      Left            =   4560
      TabIndex        =   24
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      Caption         =   "<Codigo>"
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
      TabIndex        =   23
      Top             =   2880
      Width           =   1110
   End
   Begin VB.Label lblArticulo1 
      AutoSize        =   -1  'True
      Caption         =   "<Articulo>"
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
      TabIndex        =   22
      Top             =   3360
      Width           =   1185
   End
   Begin VB.Label lblPrecio1 
      AutoSize        =   -1  'True
      Caption         =   "<Precio>"
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
      Left            =   960
      TabIndex        =   21
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Label lblNvoSaldo 
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
      Left            =   6240
      TabIndex        =   20
      Top             =   5640
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nvo. Saldo:"
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
      Left            =   4560
      TabIndex        =   19
      Top             =   5640
      Width           =   1245
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
      Left            =   6000
      TabIndex        =   18
      Top             =   960
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
      Left            =   6600
      TabIndex        =   17
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Saldo:"
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
      Left            =   6840
      TabIndex        =   16
      Top             =   1920
      Width           =   675
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
      Left            =   7800
      TabIndex        =   15
      Top             =   1920
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
      Left            =   7560
      TabIndex        =   14
      Top             =   1440
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
      Left            =   7800
      TabIndex        =   13
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
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
      TabIndex        =   12
      Top             =   360
      Width           =   945
   End
   Begin VB.Label lblPrecio2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<Precio>"
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
      Left            =   6240
      TabIndex        =   10
      Top             =   4680
      Width           =   1005
   End
   Begin VB.Label lblArticulo2 
      AutoSize        =   -1  'True
      Caption         =   "<Articulo>"
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
      Left            =   5520
      TabIndex        =   9
      Top             =   2880
      Width           =   1185
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Precio Articulo:"
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
      Left            =   4560
      TabIndex        =   8
      Top             =   4680
      Width           =   1635
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Articulo:"
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
      Left            =   4560
      TabIndex        =   7
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nueva Prenda:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Prenda a cambiar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   2235
   End
End
Attribute VB_Name = "frmCambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 30/07/02
' Modulo frmCambios - frmCambios.frm
' Ultima Modificacion - 30/07/02
'
'////////////////////////////////////////////////////////////////

Option Explicit
Dim fl() As cFlatControl

Private Sub cmdAceptar_Click()
   If Validar Then Grabar_Datos
   limpiar
End Sub

'Validamos si tan correctos los datos
Private Function Validar() As Boolean
   Validar = True
     
   If txtCodigo2.Tag = "" Then
      MsgBox "Favor de poner el nuevo articulo", vbOKOnly + vbInformation
      txtCodigo2.SetFocus
      Validar = False
      Exit Function
   End If
   
End Function

Private Sub cmdMosClave_Click()
   frmMostrarClientesVentas.Ver Me, txtNombrePago
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   inicializar
End Sub

'Inicializamos la forma
Private Sub inicializar()
Dim ctrl As Control
   Screen.MousePointer = vbHourglass
   CentrarForm Me, frmMDI
   Me.Top = 0
   Me.Left = 0
   Poner_Flat fl, Me.Controls, Me
   For Each ctrl In Controls
      If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = ""
   Next
   Crear_Encabezados
   Screen.MousePointer = vbDefault
End Sub

Private Sub Crear_Encabezados()
   With grdCambiar
      .AddColumn "K1", "Codigo", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
      .AddColumn "K2", "Articulo", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
      .AddColumn "K3", "Precio", ecgHdrTextALignRight, , 80, , , , , , , CCLSortNumeric
   End With
   
   With grdNueva
      .AddColumn "K1", "Codigo", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
      .AddColumn "K2", "Articulo", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
      .AddColumn "K3", "Precio", ecgHdrTextALignRight, , 80, , , , , , , CCLSortNumeric
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Quitar_Flat fl
End Sub

Private Sub txtCodigo1_GotFocus()
   Seleccionar_Texto txtCodigo1
   Cambiar_Color True, txtCodigo1
End Sub

Private Sub txtCodigo1_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Buscar_Clave
      Buscar_Codigo1 lblCodigo
   End If
End Sub

Private Sub txtCodigo1_LostFocus()
   Cambiar_Color False, txtCodigo1
End Sub


Private Sub txtCodigo2_GotFocus()
   Seleccionar_Texto txtCodigo2
   Cambiar_Color True, txtCodigo2
End Sub

Private Sub txtCodigo2_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
         Buscar_Codigo txtCodigo2
         If CCur(IIf(lblNvoSaldo.Caption = "", 0, lblNvoSaldo.Caption)) < 0 Then
            lblNvoSaldo.Caption = "0.00"
         End If
   End If
End Sub

Private Sub txtCodigo2_LostFocus()
   Cambiar_Color False, txtCodigo2
End Sub

'Buscamos el codigo del articulo
Private Sub Buscar_Codigo(txt As TextBox)
   On Error GoTo error
   Dim rcArticulos As New ADODB.Recordset
   
   rcArticulos.Open "SELECT * FROM Inventario WHERE Codigo='" & txt & "'", dbDatos, adOpenKeyset, adLockOptimistic
   
   With rcArticulos
      If .RecordCount = 0 Then
         MsgBox "El Articulo no se encuentra, Favor de poner correctamente el codigo", vbOKOnly + vbInformation
         txt.Tag = ""
         txt.SetFocus
      Else
         txt.Tag = !ID
         grdNueva.AddRow
         grdNueva.CellItemData(grdNueva.Rows, 1) = !ID
         grdNueva.CellText(grdNueva.Rows, 1) = !codigo
         grdNueva.CellTextAlign(grdNueva.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
         grdNueva.CellText(grdNueva.Rows, 2) = !descripcion
         grdNueva.CellTextAlign(grdNueva.Rows, 2) = DT_LEFT Or DT_WORD_ELLIPSIS
         grdNueva.CellText(grdNueva.Rows, 3) = !precio
         grdNueva.CellTextAlign(grdNueva.Rows, 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
         lblPrecio2.Caption = Format(!precio, "###,###,##0.00")
         lblTotal2.Caption = Format(CCur(IIf(lblTotal2.Caption = "", 0, lblTotal2.Caption)) + !precio, "###,###,##0.00")
         lblNvoSaldo.Caption = Format(((CCur(lblTotal.Caption) - CCur(lblPrecio1.Caption)) + CCur(lblTotal2.Caption) - CCur(lblAbonos.Caption)), "###,###,##0.00")
      End If
   End With
   
   rcArticulos.Close
   
   txtCodigo2.Text = ""
   
error:
      Maneja_Error Err
      
      Set rcArticulos = Nothing
End Sub


'Grabamos los datos del cambio
Private Sub Grabar_Datos()
   On Error GoTo error
   Dim Movimiento As Long
   Dim Renglon As Integer
   
   'Aumentamos el inventario
   For Renglon = 1 To grdCambiar.Rows
      dbDatos.Execute "UPDATE Inventario SET Existencia=Existencia+1 WHERE Codigo='" & grdCambiar.CellText(Renglon, 1) & "'"
   Next Renglon
   'Grabar_Inventario
   
   
   Movimiento = Regresa_Movimiento(False)
   Regresa_Movimiento True
   
   'Grabamos el cargo
   dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC) VALUES " & _
                  "(#" & Format(Date, "MM/DD/YY") & "#,'Cambios'," & Movimiento & ",0,'CA01','620301'," & CCur(lblPrecio1.Caption) & "," & TIPO_CARGO & ",0,'" & NombrePc & "')"
  'Grabamos el abono
  dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC) VALUES " & _
                  "(#" & Format(Date, "MM/DD/YY") & "#,'Cambios Apartados'," & Movimiento & ",0,'CA50','620750'," & CCur(lblPrecio1.Caption) & "," & TIPO_ABONO & ",0,'" & NombrePc & "')"
   
   Movimiento = Regresa_Movimiento(False)
   Regresa_Movimiento True
   
   'Disminuimos el inventario
   For Renglon = 1 To grdNueva.Rows
      dbDatos.Execute "UPDATE Inventario SET Existencia=Existencia-1 WHERE Codigo='" & grdNueva.CellText(Renglon, 1) & "'"
   Next Renglon
   
   'Grabamos el cargo
   dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC) VALUES " & _
                  "(#" & Format(Date, "MM/DD/YY") & "#,'Cambios'," & Movimiento & ",0,'CA50','620350'," & CCur(lblPrecio2.Caption) & "," & TIPO_CARGO & ",0,'" & NombrePc & "')"
   'Grabamos el abono
   dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC) VALUES " & _
                  "(#" & Format(Date, "MM/DD/YY") & "#,'Cambios Apartados'," & Movimiento & ",0,'CA01','620701'," & CCur(lblPrecio2.Caption) & "," & TIPO_ABONO & ",0,'" & NombrePc & "')"


   dbDatos.Execute "UPDATE VentasCA SET Total=" & CCur(lblTotal2.Caption) & " WHERE ID=" & txtNombrePago.Tag & ""
   
   If CCur(lblNvoSaldo.Caption) <= 0 Then
    dbDatos.Execute "UPDATE VentasCA SET Pagado=True WHERE ID=" & txtNombrePago.Tag & ""
   End If
   
   If MsgBox("¿Desea Continuar con el apartado?", vbYesNo + vbQuestion) = vbNo Then
    Grabar_Abonos
   End If

error:
   Maneja_Error Err

End Sub

'Grabamos los abonos
Private Sub Grabar_Abonos()
   On Error GoTo error
   Dim Movimiento As Integer
   Dim crImporte As Currency
   
   crImporte = Val(lblNvoSaldo.Caption)
   
   dbDatos.Execute "INSERT INTO Abonos (IDVenta,Fecha,Abono,PC) VALUES (" & _
                                txtNombrePago.Tag & ",#" & Format(Date, "MM/DD/YY") & "#," & lblNvoSaldo.Caption & ",'" & NombrePc & "')"
                                
    dbDatos.Execute "UPDATE VentasCA SET Pagado=True WHERE ID=" & txtNombrePago.Tag
   
    Movimiento = Regresa_Movimiento(False)
   
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC) VALUES " & _
                              "(#" & Format(Date, "MM/DD/YY") & "#," & Movimiento & ",0,'AB05','110101'," & crImporte & "," & TIPO_CARGO & ",0,'Abonos','" & NombrePc & "')"
    'Grabamos el abono
    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC) VALUES " & _
                              "(#" & Format(Date, "MM/DD/YY") & "#," & Movimiento & ",0,'AB05','620550'," & crImporte & "," & TIPO_ABONO & ",0,'Abonos','" & NombrePc & "')"
                              
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC) VALUES " & _
                              "(#" & Format(Date, "MM/DD/YY") & "#," & Movimiento & ",0,'AB05','199401'," & crImporte & "," & TIPO_CARGO & ",0,'Abonos','" & NombrePc & "')"

                  
    Imprimir_Recibo_Abono
    
    Regresa_Movimiento True
    
    Abrir_Cajon

error:
   Maneja_Error Err

End Sub

Private Sub txtNombrePago_GotFocus()
Cambiar_Color True, txtNombrePago
Seleccionar_Texto txtNombrePago
End Sub

Private Sub txtNombrePago_LostFocus()
Cambiar_Color False, txtNombrePago
End Sub


'Buscamos el cliente para los pagos
Public Sub Buscar_Cliente(ID As Long)
   On Error GoTo error
   Dim rcCliente As New ADODB.Recordset
   Dim rcAbonos As New ADODB.Recordset
   Dim crTotal As Currency
   Dim crAbonos As Currency
   'Dim Fecha
      
   'Fecha = Null
   rcCliente.Open "SELECT *  FROM VentasCA WHERE ID=" & ID, dbDatos, adOpenDynamic, adLockOptimistic
   
   'If Date > rcCliente!Vencimiento Then
   ' MsgBox ("El cliente ha sobrepasado su fecha de vencimiento, imposible realizar cambios")
   ' txtNombrePago.Text = ""
   ' Exit Sub
   'End If

   
   With rcCliente
      txtNombrePago.Tag = !ID
      lblTotal.Caption = Format(Format(!total, "Currency"), "###,###,###,##0.00")
      lblFolio.Caption = !folio
      lblVencimiento.Caption = !Vencimiento
      crTotal = !total
   End With
   
   rcAbonos.Open "SELECT * FROM Abonos WHERE IDVenta=" & rcCliente!ID & " ORDER BY Fecha", dbDatos, adOpenDynamic, adLockOptimistic
   
   
   With rcAbonos
      While Not .EOF
         crAbonos = crAbonos + !abono
         .MoveNext
      Wend
   End With
   
   
   lblAbonos.Caption = Format(Format(crAbonos, "Currency"), "###,###,###,##0.00")
   lblUltimoSaldo.Caption = Format(Format(crTotal - crAbonos, "Currency"), "###,###,###,##0.00")
   lblUltimoSaldo.Tag = crTotal - crAbonos
   
      
   rcCliente.Close
   rcAbonos.Close
   
error:
   Maneja_Error Err

   Set rcCliente = Nothing
   Set rcAbonos = Nothing
      
End Sub

Private Sub Buscar_Clave()
Dim rcInventario As New ADODB.Recordset

rcInventario.Open "SELECT Codigo FROM Inventario", dbDatos, adOpenDynamic, adLockOptimistic

With rcInventario
   Do While Not .EOF
      If Mid(!codigo, 1, 7) = Mid(txtCodigo1.Text, 1, 7) Then
         lblCodigo.Caption = !codigo
         Exit Do
      End If
      .MoveNext
   Loop
End With

rcInventario.Close

End Sub

Private Sub Buscar_Codigo1(lbl As Label)
   On Error GoTo error
   Dim rcArticulos As New ADODB.Recordset
   
   rcArticulos.Open "SELECT * FROM Inventario WHERE Codigo='" & lbl.Caption & "'", dbDatos, adOpenKeyset, adLockOptimistic
   
   With rcArticulos
      If .RecordCount = 0 Then
         MsgBox "El Articulo no se encuentra, Favor de poner correctamente el codigo", vbOKOnly + vbInformation
         lbl.Tag = ""
         txtCodigo1.SetFocus
      Else
         lbl.Tag = !ID
         grdCambiar.AddRow
         grdCambiar.CellItemData(grdCambiar.Rows, 1) = !ID
         grdCambiar.CellText(grdCambiar.Rows, 1) = !codigo
         grdCambiar.CellTextAlign(grdCambiar.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
         grdCambiar.CellText(grdCambiar.Rows, 2) = !descripcion
         grdCambiar.CellTextAlign(grdCambiar.Rows, 2) = DT_LEFT Or DT_WORD_ELLIPSIS
         grdCambiar.CellText(grdCambiar.Rows, 3) = !precio
         grdCambiar.CellTextAlign(grdCambiar.Rows, 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
         lblPrecio1.Caption = Format(CCur(IIf(lblPrecio1.Caption = "", "0", lblPrecio1.Caption)) + !precio, "###,###,##0.00")
      End If
   End With
   
   rcArticulos.Close
   
   txtCodigo1.Text = ""
   
error:
      Maneja_Error Err
      
      Set rcArticulos = Nothing
End Sub


Private Sub limpiar()
Dim ctrl As Control

For Each ctrl In Controls
   If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = ""
   If TypeOf ctrl Is TextBox And Mid(ctrl.Name, 1, 3) = "txt" Then ctrl.Text = ""
   If TypeOf ctrl Is vbalGrid Then ctrl.Clear
Next
End Sub

Private Sub Imprimir_Recibo_Abono()
   Dim Impresora As Printer
   
   'If Not Mostrar_Imprimir(Me, , vbPRORLandscape) Then Exit Sub
   
   Set Impresora = Printer
   
   With Impresora
    'On Error goto error
    '.PaperSize = vbPRPSUser
    .ScaleMode = vbMillimeters
    '.Height = 6000
    '.Width = 12240
    .Font = "Times New Roman"
    .FontSize = 8
    .FontBold = False
    
    '***************************************
    'Imprimimos los datos de arriba
    
    'imprimimos el no de folio
    .FontBold = True
    .Font = "Times New Roman"
    .FontSize = 10
    .CurrentX = 27
    .CurrentY = 0
    Impresora.Print "MONTEPIO"
    .FontSize = 16
    .CurrentX = 15
    .CurrentY = 5
    Impresora.Print "CASA OCAMPO"
    
    
    .Font = "Times New Roman"
    .FontBold = False
    .FontSize = 10
      
    'Imprimimos nombre cliente
    .CurrentX = 10
    .CurrentY = 15
    Impresora.Print UCase(Trim(txtNombrePago.Text))
    
    .FontBold = True
    'imprimimos encabezado
    .CurrentX = 25
    .CurrentY = 25
    Impresora.Print "Estado de Cuenta"
    .FontBold = False
    
    'imprimimos el no de folio
    .CurrentX = 10
    .CurrentY = 35
    Impresora.Print "Folio: " + lblFolio.Caption
    
    'imprimimos el total a pagar
    .CurrentX = 10
    .CurrentY = 45
    Impresora.Print "Total del Apartado: " + lblTotal.Caption
        
    'abonos anteriores
    .CurrentX = 10
    .CurrentY = 50
    Impresora.Print "Abonos anteriores: " + lblAbonos.Caption
        
    'Imprimimos la fecha de pago
    .CurrentX = 10
    .CurrentY = 60
    Impresora.Print "Fecha: " + Format(Date, "DD/MMMM/YY")
        
    'abono realizado
    .CurrentX = 10
    .CurrentY = 65
    Impresora.Print "Abono Realizado: " + lblNvoSaldo.Caption
    
    'saldo a pagar
    .CurrentX = 10
    .CurrentY = 70
    Impresora.Print "Saldo a Pagar: " + "0.00"
    
    'fecha de vencimiento
    .CurrentX = 10
    .CurrentY = 80
    Impresora.Print "Vence el dia: " + lblVencimiento.Caption
    
    
    .NewPage
    .EndDoc
    
    
   End With

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

