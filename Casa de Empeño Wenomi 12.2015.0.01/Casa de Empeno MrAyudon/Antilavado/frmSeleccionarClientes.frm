VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.2#0"; "vbalSGrid6.ocx"
Begin VB.Form frmSeleccionarClientes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccionar el Cliente"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12495
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
   ScaleHeight     =   5400
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   4920
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdClientes 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   8493
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
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   11340
      TabIndex        =   1
      Top             =   4920
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
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmSeleccionarClientes.frx":0000
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   10200
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "      &Aceptar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      TextColor       =   8537065
      Object.ToolTipText     =   ""
      Picture         =   "frmSeleccionarClientes.frx":0552
   End
End
Attribute VB_Name = "frmSeleccionarClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_Nombre As String
Dim m_Apellido As String
Dim m_ID As Long

Public Property Let Nombre(Valor As String)
   m_Nombre = Valor
End Property

Public Property Let Apellido(Valor As String)
   m_Apellido = Valor
End Property

Public Property Get IDCliente() As Long
   IDCliente = m_ID
End Property

Private Sub cmdAceptar_Click()
   Seleccionar_Cliente
End Sub

Private Sub cmdSalir_Click()
   m_ID = 0
   Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then Seleccionar_Cliente
   If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
   Inicializar
End Sub

Private Sub Inicializar()
   Screen.MousePointer = vbHourglass
   Crear_Encabezados
   Cargar_Clientes
   Timer1.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Crear_Encabezados()
   With grdClientes
      .AddColumn "K1", "Numero Identificacion", ecgHdrTextALignLeft, , 141
      .AddColumn "K2", "Cliente", ecgHdrTextALignLeft, , 209
      .AddColumn "K3", "Domicilio", ecgHdrTextALignLeft, , 163
      .AddColumn "K4", "Colonia", ecgHdrTextALignLeft, , 153
      .AddColumn "K5", "Ciudad", ecgHdrTextALignLeft, , 158
      .AddColumn "K6", "Telefono", ecgHdrTextALignLeft, , 103
      .AddColumn "K7", "Fecha Nacimiento", ecgHdrTextALignCentre, , 107, , , , , "DD/MM/YYYY"
   End With
End Sub

Private Sub Cargar_Clientes()
   On Error GoTo Error
   Dim rc As New ADODB.Recordset
   Dim Sql As String
   
   Sql = "SELECT ID,CONCAT(Nombre,' ',Apellido) AS Cliente,Direccion,Colonia,Municipio,Tel,FecNac, NumeroIdentificacion AS NumIdentificacion FROM Clientes WHERE Nombre LIKE '%" & m_Nombre & "%' AND Apellido LIKE '%" & m_Apellido & "%'"
   rc.Open Sql, dbDatos, adOpenForwardOnly, adLockOptimistic
   
   With rc
      While Not rc.EOF
         grdClientes.AddRow
         grdClientes.CellDetails grdClientes.Rows, 1, !NumIdentificacion, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , , , !ID
         grdClientes.CellDetails grdClientes.Rows, 2, !Cliente, DT_LEFT Or DT_WORD_ELLIPSIS
         grdClientes.CellDetails grdClientes.Rows, 3, !Direccion, DT_LEFT Or DT_WORD_ELLIPSIS
         grdClientes.CellDetails grdClientes.Rows, 4, !Colonia, DT_LEFT Or DT_WORD_ELLIPSIS
         grdClientes.CellDetails grdClientes.Rows, 5, !Municipio, DT_LEFT Or DT_WORD_ELLIPSIS
         grdClientes.CellDetails grdClientes.Rows, 6, !Tel, DT_LEFT Or DT_WORD_ELLIPSIS
         grdClientes.CellDetails grdClientes.Rows, 7, !FecNac & "", DT_CENTER Or DT_WORD_ELLIPSIS
         .MoveNext
      Wend
   End With
   
   rc.Close

Error:
   Maneja_Error Err
   
   Set rc = Nothing
End Sub

Private Sub Seleccionar_Cliente()
   If grdClientes.SelectedRow > 0 Then
      m_ID = grdClientes.CellItemData(grdClientes.SelectedRow, 1)
   Else
      m_ID = 0
   End If
   Me.Hide
End Sub

Private Sub grdClientes_ColumnWidthChanging(ByVal lCol As Long, lWidth As Long, bCancel As Boolean)
   grdClientes.CellText(1, lCol) = lWidth
End Sub

Private Sub grdClientes_DblClick(ByVal lRow As Long, ByVal lCol As Long)
   If lRow > 0 Then Seleccionar_Cliente
End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
   Timer1.Enabled = False
   grdClientes.SelectedRow = 1
   grdClientes.SetFocus
End Sub
