VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   8550
   Begin VB.CheckBox chkValuador 
      Appearance      =   0  'Flat
      Caption         =   "Valudador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtNombre 
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
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CheckBox chkGerente 
      Appearance      =   0  'Flat
      Caption         =   "Gerente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
   End
   Begin vbAcceleratorGrid6.vbalGrid grdUsuarios 
      Height          =   3375
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5953
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
   Begin VB.TextBox txtPass1 
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
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox txtPass 
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
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txtUsuario 
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
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   3720
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Agregar"
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
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmUsuarios.frx":1CFA
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   3720
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
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmUsuarios.frx":1D8A
   End
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   3720
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Limpiar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   255
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmUsuarios.frx":1E1B
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pemisos"
      Height          =   1095
      Left            =   3240
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton opEscritura 
         Appearance      =   0  'Flat
         Caption         =   "Lectura y &Escritura"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton opLectura 
         Appearance      =   0  'Flat
         Caption         =   "&Lectura"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBorrar 
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   3720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Eliminar"
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
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmUsuarios.frx":1F1F
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
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
      TabIndex        =   2
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Verificar Contraseña:"
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
      TabIndex        =   6
      Top             =   2640
      Width           =   2595
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña:"
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
      TabIndex        =   4
      Top             =   1800
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
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
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim bNuevo As Boolean

Private Sub cmdAgregar_Click()
   If Validar Then Grabar_Datos
End Sub

'Validamos k los datos sean correctos
Private Function Validar() As Boolean
   On Error GoTo Error
   'Dim dbCatalogos As New ADODB.Connection
   Dim rcUsuarios As New ADODB.Recordset
   Validar = True
   
   
   If txtPass.Text <> txtPass1.Text Then
      MsgBox "Las contraseñas no concuerdan, favor de volver a introducirlas", vbCritical + vbOKOnly
      Validar = False
      txtPass.Text = ""
      txtPass1.Text = ""
      txtPass.SetFocus
      GoTo Error
   End If
   
   If Trim(txtUsuario.Text) = "" Then
      MsgBox "Favor de introducir el usuario", vbCritical + vbOKOnly
      txtUsuario.SetFocus
      Validar = False
      GoTo Error
   End If
   
   If Trim(txtNombre.Text) = "" Then
      MsgBox "Favor de introducir un nombre", vbCritical + vbOKOnly
      txtNombre.SetFocus
      Validar = False
      GoTo Error
    End If
   
   'dbCatalogos.Open CONEXION & Path & "\Base De Datos\Datos.mdb" & USUARIO
   rcUsuarios.Open "SELECT Usuario FROM Usuarios WHERE Usuario='" & txtUsuario.Text & "'", dbDatos, adOpenKeyset, adLockOptimistic
   
   If rcUsuarios.RecordCount <> 0 And bNuevo Then
      MsgBox "El Usuario ya se encuentra dado de alta, favor de introducir otro nombre de usuario", vbCritical + vbOKOnly
      txtUsuario.SetFocus
      Validar = False
   End If
   
   rcUsuarios.Close
   'dbCatalogos.Close
      
Error:
   'verificamos si hay error
   Maneja_Error Err
   
   'Set dbCatalogos = Nothing
   Set rcUsuarios = Nothing
End Function

Private Sub cmdBorrar_Click()
   If grdUsuarios.SelectedRow > 0 Then
         dbDatos.Execute "DELETE * FROM Usuarios WHERE ID=" & grdUsuarios.CellItemData(grdUsuarios.SelectedRow, 1)
         grdUsuarios.RemoveRow grdUsuarios.SelectedRow
         limpiar
   End If
End Sub

Private Sub cmdLimpiar_Click()
   bNuevo = True
   limpiar
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   inicializar
End Sub

Private Sub inicializar()
   'cmdAgregar.Enabled = Not bLectura
   Screen.MousePointer = vbHourglass
   bNuevo = True
   Me.Top = 0
   Me.Left = 0
   Crear_Encabezados
   Ordenar_Grid 1, grdUsuarios, 5, 6
   Poner_Flat
   Cargar_Usuarios
   Screen.MousePointer = vbDefault
End Sub

'Cargamos los usuarios en el grid
Private Sub Cargar_Usuarios()
   On Error GoTo Error
   'Dim dbCatalogos As New ADODB.Connection
   Dim rcUsuarios As New ADODB.Recordset
   
   'dbCatalogos.Open CONEXION & Path & "\Base De Datos\Datos.mdb" & USUARIO
   rcUsuarios.Open "SELECT * FROM Usuarios", dbDatos, adOpenDynamic, adLockOptimistic
   grdUsuarios.Redraw = False
   With rcUsuarios
      While Not .EOF
         grdUsuarios.AddRow
         grdUsuarios.CellText(grdUsuarios.Rows, 1) = !Usuario
         grdUsuarios.CellItemData(grdUsuarios.Rows, 1) = !id
         grdUsuarios.CellText(grdUsuarios.Rows, 2) = !Lectura
         grdUsuarios.CellText(grdUsuarios.Rows, 3) = !password
         grdUsuarios.CellText(grdUsuarios.Rows, 4) = !GERENTE
         .MoveNext
      Wend
   End With
   grdUsuarios.Redraw = True
   rcUsuarios.Close
   'dbCatalogos.Close
   
Error:
   'Verificamos si hay error
   Maneja_Error Err
   
   Set rcUsuarios = Nothing
   'Set dbCatalogos = Nothing
End Sub


'Creamos los encabezados para la lista
Private Sub Crear_Encabezados()
   With grdUsuarios
      .ImageList = frmMDI.img
      .AddColumn "K1", "Usuarios", ecgHdrTextALignLeft, , 330, , , , , , , CCLSortString
      .AddColumn "K2", "Lectura", ecgHdrTextALignCentre, , , False
      .AddColumn "K3", "Password", ecgHdrTextALignLeft, , , False
      .AddColumn "K4", "Gerente", ecgHdrTextALignLeft, , , False
      .AddColumn "K5", "Valuador", ecgHdrTextALignLeft, , , False
   End With
End Sub


Private Sub Grabar_Datos()
   On Error GoTo Error
   'Dim dbCatalogos As New ADODB.Connection
   Dim rcUsuarios As New ADODB.Recordset
   Dim rcID As New ADODB.Recordset
   Dim Renglon As Long
   
   'dbCatalogos.Open CONEXION & Path & "\Base De Datos\Datos.mdb" & USUARIO
   
   If bNuevo Then
      rcUsuarios.Open "SELECT * FROM Usuarios WHERE ID=0", dbDatos, adOpenDynamic, adLockOptimistic
      rcUsuarios.AddNew
   Else
      rcUsuarios.Open "SELECT * FROM Usuarios WHERE ID =" & grdUsuarios.CellItemData(grdUsuarios.SelectedRow, 1), dbDatos, adOpenDynamic, adLockOptimistic
   End If
   
   With rcUsuarios
      !Usuario = txtUsuario.Text
      !password = txtPass.Text
      !Lectura = opLectura.Value
      !GERENTE = CBool(chkGerente.Value)
      !Nombre = UCase(txtNombre.Text)
      !Valuador = CBool(chkValuador.Value)
      .Update
   End With
   
   If bNuevo Then rcID.Open "SELECT MAX(ID) AS IDD FROM Usuarios", dbDatos, adOpenDynamic, adLockOptimistic
      
   With grdUsuarios
      If bNuevo Then .AddRow
      Renglon = IIf(bNuevo, .Rows, .SelectedRow)
      
      .CellText(Renglon, 1) = txtUsuario.Text
      '.CellIcon(Renglon, 1) = 15
      .CellText(Renglon, 2) = opLectura.Value
      .CellText(Renglon, 3) = txtPass.Text
      .CellText(Renglon, 4) = CBool(chkGerente.Value)
      .CellText(Renglon, 5) = CBool(chkValuador.Value)
       If bNuevo Then
         .CellItemData(Renglon, 1) = rcID!idd
         rcID.Close
      End If
   End With
   
   rcUsuarios.Close
   'dbCatalogos.Close
   limpiar
   
Error:
   'verificamos si hay error
   Maneja_Error Err
   
   Set rcUsuarios = Nothing
   'Set dbCatalogos = Nothing
   
End Sub

'Limpiamos las cajas de texto
Private Sub limpiar()
   txtUsuario.Text = ""
   txtPass.Text = ""
   txtPass1.Text = ""
   txtNombre.Text = ""
   opLectura.Value = True
   opEscritura.Value = False
   chkGerente.Value = 0
   chkValuador.Value = 0
   txtUsuario.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   'Descargamos de memoria el flat
   For i = LBound(Fl) To UBound(Fl)
      Set Fl(i) = Nothing
   Next i
End Sub

Private Sub grdUsuarios_Click(ByVal lRow As Long, ByVal lCol As Long)
   If lRow > 0 Then
      Poner_Datos lRow
      bNuevo = False
   End If
End Sub

Private Sub grdusuarios_ColumnClick(ByVal lCol As Long)
   Ordenar_Grid lCol, grdUsuarios, 5, 6
End Sub

'Ponemos en modo flat los textbox
Private Sub Poner_Flat()
   Dim Contador As Integer
   Dim Control As Object
   
   For Each Control In Controls
      If TypeOf Control Is TextBox Then
         ReDim Preserve Fl(0 To Contador)
         Set Fl(Contador) = New cFlatControl
         Fl(Contador).hWndAttach Control.hwnd, Me.hwnd, False
         Contador = Contador + 1
      ElseIf TypeOf Control Is ComboBox Then
         ReDim Preserve Fl(0 To Contador)
         Set Fl(Contador) = New cFlatControl
         Fl(Contador).hWndAttach Control.hwnd, Me.hwnd, True
         Contador = Contador + 1
      End If
   Next
   
End Sub

Private Sub grdUsuarios_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
   If grdUsuarios.SelectedRow > 0 Then
      If KeyCode = vbKeyDelete Then
         dbDatos.Execute "DELETE * FROM Usuarios WHERE ID=" & grdUsuarios.CellItemData(grdUsuarios.SelectedRow, 1)
         grdUsuarios.RemoveRow grdUsuarios.SelectedRow
         limpiar
      End If
   End If
End Sub

Private Sub txtNombre_GotFocus()
   Seleccionar_Texto txtNombre
   Cambiar_Color True, txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtNombre_LostFocus()
   Cambiar_Color False, txtNombre
End Sub

Private Sub txtPass_GotFocus()
   Seleccionar_Texto txtPass
   Cambiar_Color True, txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtPass_LostFocus()
   Cambiar_Color False, txtPass
End Sub

Private Sub txtPass1_GotFocus()
   Seleccionar_Texto txtPass1
   Cambiar_Color True, txtPass1
End Sub

Private Sub txtPass1_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtPass1_LostFocus()
   Cambiar_Color False, txtPass1
End Sub

Private Sub txtUsuario_GotFocus()
   Cambiar_Color True, txtUsuario
   Seleccionar_Texto txtUsuario
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtUsuario_LostFocus()
   Cambiar_Color False, txtUsuario
End Sub

'Ponemos la info en las caja de texto
Private Sub Poner_Datos(Renglon As Long)
Dim rcUsuarios As New ADODB.Recordset
    rcUsuarios.Open "SELECT Nombre FROM Usuarios WHERE ID=" & grdUsuarios.CellItemData(grdUsuarios.SelectedRow, 1) & "", dbDatos, adOpenDynamic, adLockOptimistic
    txtNombre.Text = IIf(IsNull(rcUsuarios!Nombre), "", rcUsuarios!Nombre)
    txtUsuario.Text = grdUsuarios.CellText(Renglon, 1)
    opLectura.Value = grdUsuarios.CellText(Renglon, 2)
    opEscritura.Value = Not grdUsuarios.CellText(Renglon, 2)
    txtPass.Text = grdUsuarios.CellText(Renglon, 3)
    txtPass1.Text = grdUsuarios.CellText(Renglon, 3)
    chkGerente.Value = IIf(CBool(grdUsuarios.CellText(Renglon, 4)), 1, 0)
    chkValuador.Value = IIf(CBool(grdUsuarios.CellText(Renglon, 5)), 1, 0)
rcUsuarios.Close
End Sub
