VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~2.OCX"
Begin VB.Form frmMostrarCuentas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7605
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBuscar 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin vbAcceleratorGrid6.vbalGrid grdCuentas 
      Height          =   3735
      Left            =   0
      TabIndex        =   2
      Top             =   720
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Buscar:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmMostrarCuentas"
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
Dim fl() As cFlatControl
Dim Carga As Boolean
Dim Ban As Boolean

Public Sub Ver(frmOBJ As Form, Ob As Object, Optional BD As Boolean = False, Optional x As Integer)
    Set obj = Ob
    Set frm = frmOBJ
    Position Me, Ob
    Carga = BD
    Ban = False
    Inicializar
    Me.Show , frmMDI
End Sub

'Inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Poner_Flat fl, Me.Controls, Me
    Crear_Encabezados
    Cargar_Datos
    Screen.MousePointer = vbDefault
End Sub

Private Sub Crear_Encabezados()

    grdCuentas.Clear True
    With grdCuentas
        .ImageList = frmMDI.img
        .AddColumn "K1", "Cuenta", ecgHdrTextALignLeft, , 120, , , , , , , CCLSortString
        .AddColumn "K2", "Descripcion", ecgHdrTextALignLeft, , 350, , , , , , , CCLSortString
    End With

End Sub

'Cargamos los nombres de los clientes
Private Sub Cargar_Datos()
   On Error GoTo error
   'Dim rcClientes As New ADODB.Recordset
   Dim rc As New ADODB.Recordset
    
   rc.Open "SELECT * FROM Cuentas", dbDatos, adOpenDynamic, adLockOptimistic
   
   grdCuentas.Clear
   grdCuentas.Redraw = False
   With grdCuentas
      While Not rc.EOF
         If (rc!CuentaContpaq & "") <> "" Then
           .AddRow
           .CellDetails .Rows, 1, rc!CuentaContpaq, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
           .CellDetails .Rows, 2, rc!DescripcionContpaq, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
         End If
         rc.MoveNext
      Wend
   End With
   On Error Resume Next
   grdCuentas.SelectedRow = 1
   grdCuentas.SetFocus
   txtBuscar.SetFocus
   grdCuentas.Redraw = True
   Err.Clear


'    rcClientes.Open "SELECT ID,CONCAT(Apellido,' ',Nombre) AS Cliente FROM clientes ORDER BY CONCAT(Apellido,' ',Nombre)", dbDatos, adOpenForwardOnly, adLockReadOnly
'    grdCuentas.Redraw = False
'
'    With rcClientes
'
'        While Not .EOF
'            grdCuentas.AddRow
'            grdCuentas.CellText(grdCuentas.Rows, 1) = !Cliente
'            grdCuentas.CellItemData(grdCuentas.Rows, 1) = !ID
'            grdCuentas.CellTextAlign(grdCuentas.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
'        .MoveNext
'        Wend
'
'    End With
'
'    grdCuentas.Redraw = True
'    rcClientes.Close
    
    rc.Close
    
error:
    Maneja_Error Err
    'Set rcClientes = Nothing
    Set rc = Nothing
End Sub

'Buscamos al cliente
Private Sub Buscar(Codigo As String)
   Dim Indice As Integer, Cadena As String
   
    For Indice = 1 To grdCuentas.Rows
        Cadena = Replace(grdCuentas.CellText(Indice, 1), "-", "")

        If UCase(Mid(Cadena, 1, Len(Codigo))) = UCase(Codigo) Then
            grdCuentas.SelectedRow = Indice
            grdCuentas.EnsureVisible Indice, 1
            Exit For
        End If

    Next Indice
   
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyReturn And grdCuentas.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat fl
End Sub

Private Sub grdCuentas_ColumnClick(ByVal lCol As Long)
    'Ordenar_Grid lCol, grdCuentas, 5, 6
End Sub

Private Sub grdCuentas_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    If grdCuentas.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub grdCuentas_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
   If KeyCode = vbKeyReturn Then If grdCuentas.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub txtBuscar_Change()
    Buscar Replace(txtBuscar.text, "-", "")
End Sub

Private Sub Seleccionar()
    
    obj.text = grdCuentas.CellText(grdCuentas.SelectedRow, 1)  'lsvFolios.SelectedItem.Text
    frm.txtEdit_KeyDown vbKeyReturn, 0
      
      
    'If Carga Then
    
    '    frm.Buscar_Cliente grdCuentas.CellItemData(grdCuentas.SelectedRow, 1)
    'Else
    '    frm.Buscar True, grdCuentas.CellText(grdCuentas.SelectedRow, 1)
    'End If

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

