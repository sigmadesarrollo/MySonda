VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Begin VB.Form frmMostrarclienteventas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBuscar 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   4620
      Width           =   5055
   End
   Begin vbAcceleratorGrid6.vbalGrid grdClientes 
      Height          =   4290
      Left            =   30
      TabIndex        =   1
      Top             =   45
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7567
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
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
      DefaultRowHeight=   17
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Buscar:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4380
      Width           =   540
   End
End
Attribute VB_Name = "frmMostrarclienteventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Opcion As Boolean
Dim obj As Object
Dim frm As Form
Dim Fl() As cFlatControl
Dim Carga As Boolean
Dim Ban As Boolean

Public Sub Ver(frmOBJ As Form, Ob As Object, Optional BD As Boolean = False, Optional x As Integer)
   Set obj = Ob
   Set frm = frmOBJ
   Position Me, Ob
   Carga = BD
   If x = 0 Then Ban = False
   grdClientes.Clear True
   Inicializar
   'Me.Show , frmMDI
End Sub

'Inicializamos la forma
Private Sub Inicializar()
   Screen.MousePointer = vbHourglass
   Poner_Flat Fl, Me.Controls, Me
   Crear_Encabezados
   Cargar_Datos
   Screen.MousePointer = vbDefault
End Sub

Private Sub Crear_Encabezados()

If Ban = False Then grdClientes.Clear True
With grdClientes
   .ImageList = frmMDI.img
   .AddColumn "K1", "Nombre", ecgHdrTextALignLeft, , 500, , , , , , , CCLSortString
    Ban = True
End With
End Sub

'Cargamos los nombres de los clientes
Private Sub Cargar_Datos()
  On Error GoTo error
  Dim rcClientes As New ADODB.Recordset
   
  rcClientes.Open "select ID,concat(apellido,' ',nombre) as cliente from clientes order by concat(apellido,'',nombre)", dbDatos, adOpenForwardOnly, adLockOptimistic
  grdClientes.Redraw = False
  With rcClientes
    While Not .EOF
      grdClientes.AddRow
      grdClientes.CellText(grdClientes.Rows, 1) = !Cliente
      grdClientes.CellTextAlign(grdClientes.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
      grdClientes.CellItemData(grdClientes.Rows, 1) = !ID
      .MoveNext
    Wend
  End With
  grdClientes.Redraw = True
  
  rcClientes.Close
  
error:
  Maneja_Error Err
  
  Set rcClientes = Nothing
   
End Sub

'Buscamos al cliente
Private Sub Buscar(Codigo As String)
   Dim Indice As Integer
   Dim Cadena As String
   
   For Indice = 1 To grdClientes.Rows
      Cadena = grdClientes.CellText(Indice, 1)
      If UCase(Mid(Cadena, 1, Len(Codigo))) = UCase(Codigo) Then
         grdClientes.SelectedRow = Indice
         grdClientes.EnsureVisible Indice, 1
         Exit For
      End If
   Next Indice
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Unload Me
   If KeyCode = vbKeyReturn And grdClientes.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Quitar_Flat Fl
End Sub

Private Sub grdClientes_ColumnClick(ByVal lCol As Long)
   Ordenar_Grid lCol, grdClientes, 5, 6
End Sub

Private Sub grdClientes_DblClick(ByVal lRow As Long, ByVal lCol As Long)
   If grdClientes.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub txtBuscar_Change()
   Buscar txtBuscar.text
End Sub

Public Function Seleccionar() As Integer

'Dim clientes As ADODB.Recordset
'
'Set clientes = New ADODB.Recordset
'
'Set clientes = dbDatos.Execute("select Nombre,Direccion,Apellido from clientes where ID=" & grdClientes.CellItemData(grdClientes.SelectedRow, 1) & "")
'If Not clientes.BOF And Not clientes.EOF Then
'    If frmVentas.tTab.SelectedTab = 1 Then
'        frmVentas.txtNombre2.Tag = grdClientes.CellItemData(grdClientes.SelectedRow, 1)
'        frmVentas.txtNombre2.text = clientes!Nombre
'        frmVentas.txtDireccion2.text = IIf(IsNull(clientes!Direccion), "", clientes!Direccion)
'        frmVentas.txtApellidos2.text = clientes!Apellido
'    Else
'        frmVentas.txtNombre.Tag = grdClientes.CellItemData(grdClientes.SelectedRow, 1)
'        frmVentas.txtNombre.text = clientes!Nombre
'        frmVentas.txtDireccion.text = IIf(IsNull(clientes!Direccion), "", clientes!Direccion)
'        frmVentas.txtApellidos.text = clientes!Apellido
'    End If
'End If
'
'Set clientes = Nothing
'Unload Me

    frmVentas.Buscar_Cliente_Ventas grdClientes.CellItemData(grdClientes.SelectedRow, 1)
    Unload Me
End Function

Private Sub txtBuscar_GotFocus()
  Seleccionar_Texto txtBuscar
  Cambiar_Color True, txtBuscar
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtBuscar_LostFocus()
  Cambiar_Color False, txtBuscar
End Sub


