VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~2.OCX"
Begin VB.Form frmMostrarUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBuscar 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   4575
   End
   Begin vbAcceleratorGrid6.vbalGrid grdUsuarios 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5318
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
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   660
   End
End
Attribute VB_Name = "frmMostrarUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Opcion As Boolean
Dim obj As Object
Dim frm As Form
Dim Ban As Integer
Dim Mostrar As Boolean

Public Sub Ver(frmOBJ As Form, Ob As Object, x As Boolean, Optional Ban As Boolean = True)
    Set obj = Ob
    Set frm = frmOBJ
    Mostrar = Ban
    Position Me, Ob
    Inicializar x
    Me.Show , frmMDI
End Sub

'Inicializamos la forma
Private Sub Inicializar(x As Boolean)
    Screen.MousePointer = vbHourglass
    If x = False Then grdUsuarios.Clear True
    Crear_Encabezados
    Cargar_Datos
    Screen.MousePointer = vbDefault
End Sub

Private Sub Crear_Encabezados()

On Error Resume Next

    With grdUsuarios
        .Clear True
        .AddColumn "K1", "Nombre de Sesión", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
        .AddColumn "K2", "Nombre", ecgHdrTextALignLeft, , 200, , , , , , , CCLSortString
    End With
End Sub

'Cargamos los nombres de los clientes
Private Sub Cargar_Datos()
Dim rcClientes As New ADODB.Recordset

On Error GoTo Error
    
    rcClientes.Open "SELECT * FROM usuarios WHERE Estatus=1", dbDatos, adOpenForwardOnly, adLockReadOnly
   
    grdUsuarios.Redraw = False
    With rcClientes
        While Not .EOF
            
            grdUsuarios.AddRow
            grdUsuarios.CellText(grdUsuarios.Rows, 1) = !Usuario
            grdUsuarios.CellItemData(grdUsuarios.Rows, 1) = !ID
            grdUsuarios.CellTextAlign(grdUsuarios.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdUsuarios.CellText(grdUsuarios.Rows, 2) = !Nombre
            grdUsuarios.CellTextAlign(grdUsuarios.Rows, 2) = DT_LEFT Or DT_WORD_ELLIPSIS
        
        .MoveNext
        Wend
    End With
    rcClientes.Close
    Set rcClientes = Nothing
    grdUsuarios.Redraw = True
    Exit Sub
  
Error:
    Maneja_Error Err
    Set rcClientes = Nothing
End Sub

'Buscamos al cliente
Private Sub Buscar(Codigo As String)
Dim Indice As Integer, Cadena As String
   
    For Indice = 1 To grdUsuarios.Rows
        
        Cadena = grdUsuarios.CellText(Indice, 1)
        If UCase(Mid(Cadena, 1, Len(Codigo))) = UCase(Codigo) Then
            grdUsuarios.SelectedRow = Indice
            grdUsuarios.EnsureVisible Indice, 1
            Exit For
        End If
    
    Next Indice
   
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyReturn And grdUsuarios.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub grdusuarios_ColumnClick(ByVal lCol As Long)
    Ordenar_Grid lCol, grdUsuarios, 5, 6
End Sub

Private Sub grdUsuarios_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    If grdUsuarios.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub txtBuscar_Change()
    Buscar txtBuscar.text
End Sub

Public Function Seleccionar() As Integer
    obj.text = grdUsuarios.CellText(grdUsuarios.SelectedRow, 1)
    obj.Tag = grdUsuarios.CellItemData(grdUsuarios.SelectedRow, 1)
    If Mostrar Then frm.Buscar grdUsuarios.CellItemData(grdUsuarios.SelectedRow, 1)
    Unload Me
End Function

Private Sub txtBuscar_GotFocus()
    Seleccionar_Texto txtBuscar
    Cambiar_Color True, txtBuscar
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBuscar_LostFocus()
    Cambiar_Color False, txtBuscar
End Sub
