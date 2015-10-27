VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Begin VB.Form frmMostrarCliente 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4455
   ClientLeft      =   2505
   ClientTop       =   900
   ClientWidth     =   7605
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
   ScaleHeight     =   4455
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin vbAcceleratorGrid6.vbalGrid grdClientes 
      Height          =   3735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
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
   Begin VB.TextBox txtBuscar 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   5055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Buscar:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   540
   End
End
Attribute VB_Name = "frmMostrarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 05/04/02
' Modulo frmMostrarCliente - frmMostrarCliente.frm
' Ultima Modificacion - 05/04/02
'
'////////////////////////////////////////////////////////////////

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
    Ban = False
    Inicializar
    Me.Show , IIf(frmOBJ.Name = "frmClientes", frmOBJ, frmMDI)
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

    grdClientes.Clear True
    With grdClientes
        .ImageList = frmMDI.img
        .AddColumn "K1", "Nombre", ecgHdrTextALignLeft, , 470, , , , , , , CCLSortString
    End With

End Sub

'Cargamos los nombres de los clientes
Private Sub Cargar_Datos()
Dim rcClientes As New ADODB.Recordset
    
On Error GoTo Error

    rcClientes.Open "SELECT ID,CONCAT(Apellido,' ',Nombre) AS Cliente FROM clientes ORDER BY CONCAT(Apellido,' ',Nombre)", dbDatos, adOpenForwardOnly, adLockReadOnly
    grdClientes.Redraw = False
  
    With rcClientes
        
        While Not .EOF
            grdClientes.AddRow
            grdClientes.CellText(grdClientes.Rows, 1) = !Cliente
            grdClientes.CellItemData(grdClientes.Rows, 1) = !ID
            grdClientes.CellTextAlign(grdClientes.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
        .MoveNext
        Wend
    
    End With

    grdClientes.Redraw = True
    rcClientes.Close
    
Error:
    Maneja_Error Err
    Set rcClientes = Nothing
End Sub

'Buscamos al cliente
Private Sub Buscar(Codigo As String)
'''''Dim Indice As Integer, Cadena As String
'''''
'''''    For Indice = 1 To grdClientes.Rows
'''''        Cadena = grdClientes.CellText(Indice, 1)
'''''
'''''        If UCase(Mid(Cadena, 1, Len(Codigo))) = UCase(Codigo) Then
'''''            grdClientes.SelectedRow = Indice
'''''            grdClientes.EnsureVisible Indice, 1
'''''            Exit For
'''''        End If
'''''
'''''    Next Indice
'''''
    Dim Indice As Integer, Cadena As String
   
    grdClientes.Redraw = False
    For Indice = 1 To grdClientes.Rows
        
        Cadena = grdClientes.CellText(Indice, 1)
        
        If UCase(Mid(Cadena, 1, Len(Codigo))) <> UCase(Codigo) Then
            
            grdClientes.RowVisible(Indice) = False
            grdClientes.EnsureVisible Indice, 1
        Else
            
            grdClientes.RowVisible(Indice) = True
        End If

    Next Indice
    grdClientes.Redraw = True
    
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
    
    obj.text = grdClientes.CellText(grdClientes.SelectedRow, 1) 'lsvFolios.SelectedItem.Text

    If Carga Then
    
        frm.Buscar grdClientes.CellItemData(grdClientes.SelectedRow, 1)
    Else
        frm.Buscar True, grdClientes.CellText(grdClientes.SelectedRow, 1)
    End If

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
