VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmTiposTarjetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Tarjetas"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTiposTarjetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   4590
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstTarjetas 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox txtTarjeta 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   5760
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
      Picture         =   "frmTiposTarjetas.frx":000C
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Tarjeta:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmTiposTarjetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Inicializar
End Sub

Private Sub Inicializar()
   Screen.MousePointer = vbHourglass
   Cargar_Datos
   Screen.MousePointer = vbDefault
End Sub

Private Sub Cargar_Datos()
   On Error GoTo Error
   Dim rc As New ADODB.Recordset
   
   rc.Open "SELECT * FROM TarjetasPuntos ORDER BY ID", m_Conexion, adOpenForwardOnly, adLockOptimistic
   lstTarjetas.Clear
   
   With rc
      While Not .EOF
         DoEvents
         lstTarjetas.AddItem !TipoTarjeta
         lstTarjetas.ItemData(lstTarjetas.NewIndex) = !ID
         .MoveNext
      Wend
   End With
    
Error:
   Maneja_Error Err
   
End Sub

Private Sub txtTarjeta_GotFocus()
   Seleccionar_Texto txtTarjeta
   Cambiar_Color True, txtTarjeta
End Sub

Private Sub txtTarjeta_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Mayusculas(KeyAscii)

   If KeyAscii = vbKeyReturn Then
      Grabar_Datos txtTarjeta.Text
   End If
End Sub

Private Sub txtTarjeta_LostFocus()
   Cambiar_Color False, txtTarjeta
End Sub

Private Sub Grabar_Datos(Tarjeta As String)
   On Error GoTo Error
   
   If Tarjeta <> "" Then
      m_Conexion.Execute "INSERT INTO TarjetasPuntos (TipoTarjeta,FechaCreacion) VALUES ('" & Tarjeta & "','" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "')"
      Cargar_Datos
      txtTarjeta.Text = ""
      txtTarjeta.SetFocus
   End If
   
Error:
   Maneja_Error Err
   
End Sub

Private Function Mayusculas(Codigo As Integer) As Integer
    
    Mayusculas = Asc(UCase(Chr(Codigo)))

End Function


