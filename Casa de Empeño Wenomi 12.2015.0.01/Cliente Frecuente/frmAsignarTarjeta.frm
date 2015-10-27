VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmAsignarTarjeta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar Tarjeta"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAsignarTarjeta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFolio 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   3015
   End
   Begin VB.ComboBox cmbTarjetas 
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
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   1200
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
      Picture         =   "frmAsignarTarjeta.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdGrabar 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Aceptar"
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
      Picture         =   "frmAsignarTarjeta.frx":055E
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Numero de Tarjeta:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Tarjeta:"
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
      Left            =   585
      TabIndex        =   3
      Top             =   240
      Width           =   1950
   End
End
Attribute VB_Name = "frmAsignarTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_IDCliente As Long
Dim m_IDUsuario As Long

Public Property Let IDCliente(Valor As Long)
   m_IDCliente = Valor
End Property

Public Property Let IDUsuario(Valor As Long)
   m_IDUsuario = Valor
End Property

Private Sub cmbTarjetas_DropDown()
   Cambiar_Color True, cmbTarjetas
End Sub

Private Sub cmbTarjetas_GotFocus()
   Cambiar_Color True, cmbTarjetas
End Sub

Private Sub cmbTarjetas_LostFocus()
   Cambiar_Color False, cmbTarjetas
End Sub

Private Sub cmdGrabar_Click()
   If Validar Then Grabar_Datos txtFolio.Text, m_IDCliente, m_IDUsuario
End Sub

Private Function Validar() As Boolean
   Validar = True

   If m_IDCliente = 0 Then
      MsgBox "Favor de seleccionar el cliente", vbOKOnly Or vbCritical
      Validar = False
      Exit Function
   End If
   
   
   If cmbTarjetas.Text = "" Then
      MsgBox "Favor de seleccionar una tarjeta", vbOKOnly Or vbCritical
      Validar = False
      Exit Function
   End If
      
   If txtFolio.Text = "" Then
      MsgBox "Favor de introducir el numero de tarjeta", vbOKOnly Or vbCritical
      Validar = False
      Exit Function
   End If

End Function

Private Sub Grabar_Datos(NumeroTarjeta As String, IDCliente As Long, IDUsuario As Long)
   
On Error GoTo Error
       
    Dim rc As New ADODB.Recordset
    Dim Sql As String, Existe As Boolean

    rc.Open "SELECT COUNT(ID) AS Existe FROM asignaciontarjetas WHERE NumeroTarjeta = '" & NumeroTarjeta & "' AND IDTarjeta = " & cmbTarjetas.ItemData(cmbTarjetas.ListIndex), m_Conexion, adOpenStatic, adLockOptimistic
    Existe = rc!Existe
    rc.Close
    
    If Existe Then
        MsgBox "Ya existe el número de tarjeta", vbCritical
    Else
        Sql = "INSERT INTO AsignacionTarjetas (Fecha,NumeroTarjeta,IDTarjeta,IDCliente,IDUsuario,PC) VALUES ('" & _
        Format(Now, "YYYY/MM/DD HH:MM:SS") & "','" & NumeroTarjeta & "'," & cmbTarjetas.ItemData(cmbTarjetas.ListIndex) & "," & IDCliente & "," & IDUsuario & ",'" & Nombre_Pc & "')"

        m_Conexion.Execute Sql
        
        MsgBox "Tarjeta agregada correctamente", vbOKOnly Or vbInformation
    End If

    Unload Me
   
Error:
   Maneja_Error Err
   
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Inicializar
End Sub

Private Sub Inicializar()
   Cargar_Tarjetas
End Sub

Private Sub txtFolio_GotFocus()
   Seleccionar_Texto txtFolio
   Cambiar_Color True, txtFolio
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Numeros(KeyAscii)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtFolio_LostFocus()
   Cambiar_Color False, txtFolio
End Sub

Private Sub Cargar_Tarjetas()
   On Error GoTo Error
   Dim rc As New ADODB.Recordset
   
   rc.Open "SELECT * FROM TarjetasPuntos", m_Conexion, adOpenDynamic, adLockOptimistic
   
   cmbTarjetas.Clear
   With rc
      While Not .EOF
         cmbTarjetas.AddItem !TipoTarjeta
         cmbTarjetas.ItemData(cmbTarjetas.NewIndex) = !ID
         .MoveNext
      Wend
   End With

   rc.Close
   
Error:
   Maneja_Error Err
   
   Set rc = Nothing
End Sub
