VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Begin VB.Form frmMuestraarticulos 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9435
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
   ScaleHeight     =   4920
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBuscar 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   5055
   End
   Begin vbAcceleratorGrid6.vbalGrid grdArticulos 
      Height          =   4305
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   7594
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
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Buscar:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   540
   End
End
Attribute VB_Name = "frmMuestraarticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim obj As Object
Dim frm As Form
Dim Fl() As cFlatControl
Dim Carga As Boolean, Ban As Boolean, Seccion As Integer

Public Sub Ver(frmOBJ As Form, Ob As Object, Optional BD As Boolean = False, Optional x As Integer, Optional Pestaña As Integer = 0)
    Set obj = Ob
    Set frm = frmOBJ
    Seccion = Pestaña
    Position Me, Ob
    Carga = BD
    If x = 0 Then Ban = False
    Inicializar
    Me.Show , frmMDI
End Sub

'Inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Crear_Encabezados
    Cargar_Datos
    Poner_Flat Fl, Me.Controls, Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub Crear_Encabezados()
    
    grdArticulos.Clear True
    With grdArticulos
        .ImageList = frmMDI.img
        .AddColumn "K1", "Código", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "K2", "Tipo", ecgHdrTextALignLeft, , 95, False, , , , , , CCLSortString
        .AddColumn "K3", "Descripción", ecgHdrTextALignLeft, , 210, , , , , , , CCLSortString
        .AddColumn "K4", "Existencia", ecgHdrTextALignRight, , 60, , , , , , , CCLSortNumeric
        .AddColumn "K5", "Precio", ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Marca", ecgHdrTextALignLeft, , 80, , , , , , , CCLSortString
        .AddColumn "K7", "Modelo", ecgHdrTextALignLeft, , 80, , , , , , , CCLSortString
        Ban = True
    End With

End Sub

'Cargamos los nombres de los clientes
Private Sub Cargar_Datos()
Dim rcArticulos As New ADODB.Recordset
Dim Iva As Double

On Error GoTo Error
    
'    If Carga Then
'
'        rcArticulos.Open "SELECT d.ID,d.Codigo,d.TipoPrenda,d.Descripcion,d.IDEntrada,d.Cantidad,d.TipoEntrada,d.PrecioVitrina,d.Observaciones,d.Marca,d.Modelo,tipo.Descripcion AS TipoPrenda FROM detallesentradainventario d LEFT JOIN tipo ON d.Tipo=tipo.ID WHERE d.Cantidad>0 AND (d.TipoEntrada=" & ENTRADAALMONEDA & " OR d.TipoEntrada=" & ENTRADACOMPRA & " OR d.TipoEntrada=" & ENTRADADOTACION & " OR d.TipoEntrada=" & ENTRADATRASPASO & ") ORDER BY d.Codigo", dbDatos, adOpenForwardOnly, adLockOptimistic
'    Else
'
'        rcArticulos.Open "SELECT DISTINCT (Apellidos + ' ' + Nombre) AS cliente FROM empeno", dbDatos, adOpenDynamic, adLockOptimistic
'    End If
    If Carga Then
        
        rcArticulos.Open "SELECT d.ID,d.Codigo,d.TipoPrenda,d.Descripcion,d.IDEntrada,d.Cantidad,d.TipoEntrada,d.PrecioVitrina,d.Observaciones,d.Marca,d.Modelo,tipo.Descripcion AS TipoPrenda FROM detallesentradainventario d LEFT JOIN tipo ON d.Tipo=tipo.ID WHERE d.Cantidad>0 AND (d.TipoEntrada=" & D_VENTA & " OR d.TipoEntrada=" & ENTRADAALMONEDA & " OR d.TipoEntrada=" & ENTRADACOMPRA & " OR d.TipoEntrada=" & ENTRADADOTACION & ") ORDER BY d.Codigo", dbDatos, adOpenForwardOnly, adLockOptimistic
    Else
        
        rcArticulos.Open "SELECT DISTINCT (Apellidos + ' ' + Nombre) AS cliente FROM empeno", dbDatos, adOpenDynamic, adLockOptimistic
    End If
    
    Iva = Regresa_Valor_BD("IVAVentas") / 100
    grdArticulos.Redraw = False
    With rcArticulos
        
        While Not .EOF
            grdArticulos.AddRow
            grdArticulos.CellItemData(grdArticulos.Rows, 1) = !ID
            grdArticulos.CellText(grdArticulos.Rows, 1) = !Codigo
            grdArticulos.CellTextAlign(grdArticulos.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdArticulos.CellText(grdArticulos.Rows, 2) = !TipoPrenda
            grdArticulos.CellTextAlign(grdArticulos.Rows, 2) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdArticulos.CellText(grdArticulos.Rows, 3) = !Descripcion '''''& " " & !Observaciones
            grdArticulos.CellItemData(grdArticulos.Rows, 3) = !IDEntrada
            grdArticulos.CellTextAlign(grdArticulos.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdArticulos.CellText(grdArticulos.Rows, 4) = !Cantidad
            grdArticulos.CellTextAlign(grdArticulos.Rows, 4) = DT_RIGHT
            grdArticulos.CellText(grdArticulos.Rows, 5) = !PrecioVitrina * (1 + Iva)
            grdArticulos.CellTextAlign(grdArticulos.Rows, 5) = DT_RIGHT
            grdArticulos.CellText(grdArticulos.Rows, 6) = !Marca
            grdArticulos.CellText(grdArticulos.Rows, 7) = !Modelo
        .MoveNext
        Wend
    
    End With
    rcArticulos.Close
    Set rcArticulos = Nothing
    grdArticulos.Redraw = True
    Exit Sub
  
Error:
    Maneja_Error Err
    Set rcArticulos = Nothing
End Sub

'Buscamos al articulo
Private Sub Buscar(Codigo As String)
Dim Indice As Integer, Cadena As String
   
    For Indice = 1 To grdArticulos.Rows
      
      Cadena = grdArticulos.CellText(Indice, 1)
      If UCase(Mid(Cadena, 1, Len(Codigo))) = UCase(Codigo) Then
         grdArticulos.SelectedRow = Indice
         grdArticulos.EnsureVisible Indice, 1
         Exit For
      End If
   
   Next Indice
   
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyReturn And grdArticulos.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdArticulos_ColumnClick(ByVal lCol As Long)
    Ordenar_Grid lCol, grdArticulos, 5, 6
End Sub

Private Sub grdArticulos_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    If grdArticulos.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub txtBuscar_Change()
    Buscar txtBuscar.text
End Sub

Private Sub Seleccionar()

    Select Case Seccion
    Case 0
        
        frm.MuestraDatos grdArticulos.CellItemData(grdArticulos.SelectedRow, 1), frm.grdArticulos, frm.txtCodigo, Seccion
    Case 1
        
        frm.MuestraDatos grdArticulos.CellItemData(grdArticulos.SelectedRow, 1), frm.grdArticulosApa, frm.txtCodigoApa, Seccion
    Case 2
        
        frm.MuestraDatos grdArticulos.CellItemData(grdArticulos.SelectedRow, 1)
    
    End Select
    
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
