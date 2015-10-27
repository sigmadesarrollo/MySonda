VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Begin VB.Form frmMostrarGrupo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4455
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
   Icon            =   "frmMostrarGrupo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBuscar 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   5055
   End
   Begin vbAcceleratorGrid6.vbalGrid grdGrupos 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Buscar:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   540
   End
End
Attribute VB_Name = "frmMostrarGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 03/06/2002
' Modulo frmMostrarGrupo - frmMostrarGrupo.frm
' Ultima Modificacion - 03/06/2002
'
'////////////////////////////////////////////////////////////////

Option Explicit
Dim opcion As Boolean
Dim obj As Object
Dim frm As Form
Dim forma As Integer
Dim Fl() As cFlatControl

Dim ban As Boolean

Public Sub Ver(frmOBJ As Form, Ob As Object, Optional x As Integer = 0, Optional opcion As Integer = 1)
   Set obj = Ob
   Set frm = frmOBJ
   Position Me, Ob
   forma = opcion
   If x = 0 Then ban = False Else ban = True
   If ban = True Then grdGrupos.Clear True
   inicializar
   Me.Show , frmMDI
End Sub

'Inicializamos la forma
Private Sub inicializar()
   Screen.MousePointer = vbHourglass
   Poner_Flat Fl, Me.Controls, Me
   Crear_Encabezados
   cargar_datos
   Screen.MousePointer = vbDefault
End Sub

Private Sub Crear_Encabezados()
   With grdGrupos
      .ImageList = frmMDI.img
      .AddColumn "K1", "Clave", ecgHdrTextALignLeft, , 81, , , , , , , CCLSortString
      .AddColumn "K2", "Descripción", ecgHdrTextALignLeft, , 289, , , , , , , CCLSortString
   End With
End Sub

'Cargamos los Grupos
Private Sub cargar_datos()
  On Error GoTo error
  Dim rcGrupos As New ADODB.Recordset
   
  rcGrupos.Open "SELECT * FROM Grupos order by clave", dbDatos, adOpenForwardOnly, adLockReadOnly
   
  grdGrupos.Redraw = False
  With rcGrupos
    While Not .EOF
      grdGrupos.AddRow
      grdGrupos.CellText(grdGrupos.Rows, 1) = !clave
      grdGrupos.CellItemData(grdGrupos.Rows, 1) = !ID
      grdGrupos.CellTextAlign(grdGrupos.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
      grdGrupos.CellText(grdGrupos.Rows, 2) = !Descripcion
      grdGrupos.CellTextAlign(grdGrupos.Rows, 2) = DT_LEFT Or DT_WORD_ELLIPSIS
      .MoveNext
    Wend
  End With
  grdGrupos.Redraw = True
  
  rcGrupos.Close
  
error:
  Maneja_Error Err
  
  Set rcGrupos = Nothing
   
End Sub

'Buscamos al cliente
Private Sub buscar(Codigo As String)
   Dim Indice As Integer
   Dim Cadena As String
   
   For Indice = 1 To grdGrupos.Rows
      Cadena = grdGrupos.CellText(Indice, 2)
      If UCase(Mid(Cadena, 1, Len(Codigo))) = UCase(Codigo) Then
         grdGrupos.SelectedRow = Indice
         grdGrupos.EnsureVisible Indice, 1
         Exit For
      End If
   Next Indice
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Unload Me
   If KeyCode = vbKeyReturn And grdGrupos.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Quitar_Flat Fl
End Sub

Private Sub grdGrupos_ColumnClick(ByVal lCol As Long)
   Ordenar_Grid lCol, grdGrupos, 5, 6
End Sub

Private Sub grdGrupos_DblClick(ByVal lRow As Long, ByVal lCol As Long)
   If grdGrupos.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub txtBuscar_Change()
   buscar txtBuscar.Text
End Sub

Private Sub Seleccionar()
   obj.Text = grdGrupos.CellText(grdGrupos.SelectedRow, 1) 'lsvFolios.SelectedItem.Text
   obj.Tag = grdGrupos.CellItemData(grdGrupos.SelectedRow, 1)
   On Error Resume Next
   frm.Buscar_Grupo grdGrupos.CellText(grdGrupos.SelectedRow, 1)
   frm.lblDescripcion.Caption = grdGrupos.CellText(grdGrupos.SelectedRow, 2)
   If forma = 1 Then
    frmInventario.txtGrupo_KeyPress (13)
   End If
   Unload Me
End Sub

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
