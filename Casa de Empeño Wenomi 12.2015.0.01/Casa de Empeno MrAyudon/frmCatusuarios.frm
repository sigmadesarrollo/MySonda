VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCatusuarios 
   Caption         =   "Catálogo de Usuarios"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   Icon            =   "frmCatusuarios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin vbAcceleratorGrid6.vbalGrid grdUsuarios 
      Height          =   2625
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   4630
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
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   2760
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
      Picture         =   "frmCatusuarios.frx":000C
   End
End
Attribute VB_Name = "frmCatusuarios"
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

Public Sub Ver(frmOBJ As Form, Ob As Object, Optional Opcion As Boolean = True)
   Set obj = Ob
   Set frm = frmOBJ
   Position Me, Ob
   Op = Opcion
   inicializar
   Me.Show , frmMDI
End Sub

'Inicializamos la forma
Private Sub inicializar()
   Screen.MousePointer = vbHourglass
   Poner_Flat Fl, Me.Controls, Me
   Crear_Encabezados
   Cargar_Datos
   Screen.MousePointer = vbDefault
End Sub

Private Sub Crear_Encabezados()
On Error GoTo Error
   With grdUsuarios
      .ImageList = frmMDI.img
      .AddColumn "K1", "Nombre de Sesión", ecgHdrTextALignLeft, , 200, , , , , , , CCLSortString
      .AddColumn "K2", "Nombre", ecgHdrTextALignLeft, , 280, , , , , , , CCLSortString
    End With

Error:
  Maneja_Error Err
   
End Sub

'Cargamos los nombres de los clientes
Private Sub Cargar_Datos()
  On Error GoTo Error
  Dim rcUsuarios As New ADODB.Recordset
  
  
 rcUsuarios.Open "select * from Usuarios order by usuario", dbDatos, adOpenDynamic, adLockOptimistic
     
  grdUsuarios.Redraw = False
  With rcUsuarios
    While Not .EOF
      grdUsuarios.AddRow
      grdUsuarios.CellItemData(grdUsuarios.Rows, 1) = !ID
      grdUsuarios.CellText(grdUsuarios.Rows, 1) = !Usuario
      grdUsuarios.CellTextAlign(grdUsuarios.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
      
        grdUsuarios.CellText(grdUsuarios.Rows, 2) = !Nombre
      grdUsuarios.CellTextAlign(grdUsuarios.Rows, 2) = DT_LEFT
      .MoveNext
    Wend
  End With
  grdUsuarios.Redraw = True
  
  rcUsuarios.Close
  
Error:
  Maneja_Error Err
  
  Set rcUsuarios = Nothing
   
End Sub

'Buscamos al cliente
Private Sub Buscar(Codigo As String)
   Dim Indice As Integer
   Dim Cadena As String
   
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

Private Sub Form_Unload(Cancel As Integer)
  Quitar_Flat Fl
End Sub

Private Sub grdUsuarios_DblClick(ByVal lRow As Long, ByVal lCol As Long)
   If grdUsuarios.SelectedRow > 0 Then Seleccionar
End Sub

Private Sub Seleccionar()
   obj.Text = grdUsuarios.CellText(grdUsuarios.SelectedRow, 1) 'lsvFolios.SelectedItem.Text
   On Error Resume Next
   frm.Buscar_Cliente grdUsuarios.CellItemData(grdUsuarios.SelectedRow, 1)
   Unload Me
End Sub

