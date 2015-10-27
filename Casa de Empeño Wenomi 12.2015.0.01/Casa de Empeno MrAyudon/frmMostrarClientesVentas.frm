VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Begin VB.Form frmMostrarClientesVentas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8070
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBuscar 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   5055
   End
   Begin vbAcceleratorGrid6.vbalGrid grdClientes 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
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
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
      DefaultRowHeight=   17
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Buscar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   540
   End
End
Attribute VB_Name = "frmMostrarClientesVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 13/06/02
' Modulo frmMostrarClientesVentas - frmMostrarClientesVentas.frm
' Ultima Modificacion - 13/06/02
'
'////////////////////////////////////////////////////////////////

Option Explicit

Dim Opcion As Boolean
Dim obj As Object
Dim frm As Form
Dim Op As Boolean
Dim Fl() As cFlatControl

Public Sub ver(frmOBJ As Form, Ob As Object, Optional Opcion As Boolean = True)
   Set obj = Ob
   Set frm = frmOBJ
   Position Me, Ob
   Op = Opcion
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
    
    With grdClientes
        .ImageList = frmMDI.img
        .AddColumn "K1", "Nombre", ecgHdrTextALignLeft, , 260, , , , , , , CCLSortString
        .AddColumn "K2", "Folio", ecgHdrTextALignRight, , 70, , , , , , , CCLSortNumeric
        .AddColumn "K3", "Fecha", ecgHdrTextALignCentre, , 90, , , , , "DD/MMM/YY", , CCLSortDate
        .AddColumn "K4", "Vencimiento", ecgHdrTextALignCentre, , 90, , , , , "DD/MMM/YY", , CCLSortDate
    End With
    Exit Sub
    
error:
    Maneja_Error Err
   
End Sub

'Cargamos los nombres de los clientes
Private Sub Cargar_Datos()
Dim rcClientes As New ADODB.Recordset

On Error GoTo error

    If Op Then
        rcClientes.Open "SELECT concat(clientes.Apellido,' ',clientes.Nombre) AS Cliente,ventas.Folio,ventas.Fecha,ventas.Vencimiento,ventas.ID FROM ventas Left Join clientes on ventas.IDCliente=clientes.ID WHERE ventas.Apartado=1 AND ventas.Pagado=0 AND ventas.Cancelado=0 ORDER BY ventas.Fecha,ventas.Folio", dbDatos, adOpenForwardOnly, adLockOptimistic
    Else
        rcClientes.Open "SELECT concat(clientes.Apellido,' ',clientes.Nombre) AS Cliente,ventas.Folio,ventas.Fecha,ventas.Vencimiento,ventas.ID FROM ventas Left Join clientes on ventas.IDCliente=clientes.ID WHERE ventas.Cancelado=0 ORDER BY ventas.Fecha,ventas.Folio", dbDatos, adOpenForwardOnly, adLockOptimistic
    End If

    grdClientes.Redraw = False
    With rcClientes
        
        While Not .EOF
            grdClientes.AddRow
            grdClientes.CellItemData(grdClientes.Rows, 1) = !ID
            grdClientes.CellText(grdClientes.Rows, 1) = !Cliente
            grdClientes.CellTextAlign(grdClientes.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdClientes.CellText(grdClientes.Rows, 2) = !Folio
            grdClientes.CellTextAlign(grdClientes.Rows, 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdClientes.CellText(grdClientes.Rows, 3) = !Fecha
            grdClientes.CellTextAlign(grdClientes.Rows, 3) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdClientes.CellText(grdClientes.Rows, 4) = !Vencimiento
            grdClientes.CellTextAlign(grdClientes.Rows, 4) = DT_CENTER Or DT_WORD_ELLIPSIS
        .MoveNext
        Wend
    
    End With
    rcClientes.Close
    Set rcClientes = Nothing
    grdClientes.Redraw = True
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcClientes = Nothing
End Sub

'Buscamos al cliente
Private Sub Buscar(Codigo As String)
Dim Indice As Integer, Cadena As String
   
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

Private Sub Seleccionar()
    obj.text = IIf(IsNull(grdClientes.CellText(grdClientes.SelectedRow, 1)), "", grdClientes.CellText(grdClientes.SelectedRow, 1))
    On Error Resume Next
    frm.Buscar_Cliente grdClientes.CellItemData(grdClientes.SelectedRow, 1)
    Unload Me
End Sub

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
