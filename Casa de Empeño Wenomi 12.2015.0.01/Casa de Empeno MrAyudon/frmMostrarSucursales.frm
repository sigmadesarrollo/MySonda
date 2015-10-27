VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Begin VB.Form frmMostrarSucursales 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vbAcceleratorGrid6.vbalGrid grdSucursales 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6588
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
      DisableIcons    =   -1  'True
   End
End
Attribute VB_Name = "frmMostrarSucursales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Opcion As Boolean
Dim obj As Object
Dim frm As Form
Dim Op As Boolean
Dim Fl() As cFlatControl
Dim MuestraTodas As Boolean

Public Sub ver(frmOBJ As Form, Ob As Object, Optional Opcion As Boolean = True, Optional Muestra As Boolean = False)
   Set obj = Ob
   Set frm = frmOBJ
   Position Me, Ob
   Op = Opcion
   MuestraTodas = Muestra
   Inicializar
   Me.Show , frmMDI
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

On Error GoTo error
    
    With grdSucursales
        .Clear True
        .ImageList = frmMDI.img
        .AddColumn "K1", "Razón Social", ecgHdrTextALignLeft, , 250, , , , , , , CCLSortString
        .AddColumn "K2", "Nombre Comercial", ecgHdrTextALignLeft, , 150, , , , , , , CCLSortString
        .AddColumn "K3", "Rfc", ecgHdrTextALignLeft, , 120, , , , , , , CCLSortString
    End With
    Exit Sub
    
error:
  Maneja_Error Err
End Sub

'Cargamos las sucursales
Private Sub Cargar_Datos()
  On Error GoTo error
  Dim rcClientes As New ADODB.Recordset
  
  
  rcClientes.Open "SELECT * FROM sucursales " & IIf(MuestraTodas = True, "", "where Activa=0") & " ORDER BY Clave", dbDatos, adOpenForwardOnly, adLockReadOnly
     
  grdSucursales.Redraw = False
  With rcClientes
    While Not .EOF
      grdSucursales.AddRow
      grdSucursales.CellItemData(grdSucursales.Rows, 1) = !ID
      grdSucursales.CellText(grdSucursales.Rows, 1) = !RazonSocial
      grdSucursales.CellTextAlign(grdSucursales.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
      grdSucursales.CellText(grdSucursales.Rows, 2) = !NombreComercial
      grdSucursales.CellTextAlign(grdSucursales.Rows, 2) = DT_LEFT Or DT_WORD_ELLIPSIS
      grdSucursales.CellText(grdSucursales.Rows, 3) = !RFC
      grdSucursales.CellTextAlign(grdSucursales.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
      .MoveNext
    Wend
  End With
  grdSucursales.Redraw = True
  
  rcClientes.Close
  
error:
  Maneja_Error Err
  
  Set rcClientes = Nothing
   
End Sub

'Buscamos al cliente
Private Sub Buscar(Codigo As String)
   Dim Indice As Integer
   Dim Cadena As String
   
   For Indice = 1 To grdSucursales.Rows
      Cadena = grdSucursales.CellText(Indice, 1)
      If UCase(Mid(Cadena, 1, Len(Codigo))) = UCase(Codigo) Then
         grdSucursales.SelectedRow = Indice
         grdSucursales.EnsureVisible Indice, 1
         Exit For
      End If
   Next Indice
   
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Unload Me
   If KeyCode = vbKeyReturn And grdSucursales.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub grdsucursales_ColumnClick(ByVal lCol As Long)
   Ordenar_Grid lCol, grdSucursales, 5, 6
End Sub

Private Sub grdSucursales_DblClick(ByVal lRow As Long, ByVal lCol As Long)
   If grdSucursales.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub Seleccionar()
   On Error GoTo error
   frm.BuscarSucursal grdSucursales.CellItemData(grdSucursales.SelectedRow, 1)
   Unload Me
error:
    Maneja_Error Err
End Sub
