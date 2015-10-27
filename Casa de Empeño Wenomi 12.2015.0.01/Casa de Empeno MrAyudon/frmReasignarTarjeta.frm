VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmReasignarTarjeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reasignar Tarjeta"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   9570
   Begin VB.CheckBox chkTransferirPuntos 
      Caption         =   "Transferir Puntos"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtTarjetaNueva 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   480
      MaxLength       =   20
      TabIndex        =   4
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox txtTarjetaAnterior 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   480
      MaxLength       =   20
      TabIndex        =   3
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox txtApellidoPaterno 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   480
      MaxLength       =   60
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtNombre 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   480
      MaxLength       =   20
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosCliente 
      Height          =   225
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   397
      AlignCaption    =   4
      AutoSize        =   0   'False
      Caption         =   ". . ."
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
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Top             =   4560
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
      Picture         =   "frmReasignarTarjeta.frx":0000
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   4560
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
      Picture         =   "frmReasignarTarjeta.frx":0552
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   9135
      Begin VB.TextBox txtApellidoMaterno 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   3960
         TabIndex        =   28
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtRFC 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   6120
         TabIndex        =   27
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtCuidad 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1080
         TabIndex        =   25
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtDireccion 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1200
         TabIndex        =   12
         Top             =   1800
         Width           =   7695
      End
      Begin VB.Label lblRFC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RFC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5640
         TabIndex        =   26
         Top             =   2280
         Width           =   345
      End
      Begin VB.Label lblCiudad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   630
      End
      Begin VB.Label lblPuntos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6120
         TabIndex        =   17
         Top             =   3000
         Width           =   120
      End
      Begin VB.Label lblPuntosAcumulados 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Puntos Acumulados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4080
         TabIndex        =   16
         Top             =   3000
         Width           =   1830
      End
      Begin VB.Label lblTarjetaA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarjeta a Reasignar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   3360
         Width           =   1770
      End
      Begin VB.Label lblTarjetaAnterior 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarjeta Anterior"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   2760
         Width           =   1470
      End
      Begin VB.Label lblDireccion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   825
      End
      Begin VB.Label lblApellidoMaterno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido Materno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3960
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblApellidoPaterno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido Paterno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido Paterno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   15
      Left            =   480
      TabIndex        =   23
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido Materno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   13
      Left            =   4200
      TabIndex        =   22
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   21
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tarjeta Anterior"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   600
      TabIndex        =   20
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tarjeta a Reasignar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   600
      TabIndex        =   19
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos Acumulados:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   4320
      TabIndex        =   18
      Top             =   2520
      Width           =   1890
   End
End
Attribute VB_Name = "frmReasignarTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Dim ClienteEmp As clientes
Dim m_IDCliente As Long
Dim m_IDUsuario As Long

Public Property Let IDCliente(Valor As Long)
   m_IDCliente = Valor
End Property

Public Property Let IDUsuario(Valor As Long)
   m_IDUsuario = Valor
End Property


Private Sub Inicializar()
   
End Sub


Private Sub cmdAceptar_Click()
IDUsuario = frmMDI.IDUsuario
'Aqui se guarda la nueva tarjeta
'Se valida la tarjeta que no exista en la base de datos
'If Validar Then Grabar_Datos txtTarjetaNueva.text, Me.IDCliente, frmMDI.IDUsuario
If Validar Then Grabar_Datos txtTarjetaNueva.text
End Sub

Private Function Validar() As Boolean
   Validar = True

   If m_IDCliente = 0 Then
      MsgBox "Favor de seleccionar el cliente", vbOKOnly Or vbCritical
      Validar = False
      Exit Function
   End If
   
        
   If txtTarjetaNueva.text = "" Then
      MsgBox "Favor de introducir el Nuevo numero de tarjeta", vbOKOnly Or vbCritical
      Validar = False
      Exit Function
   End If

End Function

Private Sub Grabar_Datos(NumeroTarjeta As String)
   
On Error GoTo Error
       
    Dim rc As New ADODB.Recordset
    Dim Sql As String, Existe As Boolean

    rc.Open "SELECT COUNT(ID) AS Existe FROM asignaciontarjetas WHERE NumeroTarjeta = '" & NumeroTarjeta & "' AND IDTarjeta = 1", dbDatos, adOpenStatic, adLockOptimistic
    Existe = rc!Existe
    rc.Close
    
    If Existe Then
        MsgBox "Ya existe el nuevo número de tarjeta asignado a otro cliente", vbCritical
    Else
                
        If chkTransferirPuntos.Value = "1" Then
             Sql = "UPDATE asignaciontarjetas SET NumeroTarjeta = '" & NumeroTarjeta & "', UsuarioMovimiento =" & m_IDUsuario & " WHERE IDCliente = " & m_IDCliente & " and IDTarjeta = 1 "
        Else
            Sql = "UPDATE asignaciontarjetas SET Puntos = 0, NumeroTarjeta = '" & NumeroTarjeta & "', UsuarioMovimiento =" & m_IDUsuario & " WHERE IDCliente = " & m_IDCliente & " and IDTarjeta = 1 "
        End If
        
         
        dbDatos.Execute Sql
        
        MsgBox "Tarjeta reasignada correctamente", vbOKOnly Or vbInformation
    End If

    Unload Me
   
Error:
   Maneja_Error Err
   
End Sub

Private Sub cmdMosCliente_Click()
frmMostrarCliente.Ver Me, txtNombre, True
End Sub

Private Sub cmdSalir_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Inicializar
End Sub

'buscamos al cliente y le mandamos el parametro del tipo de busqueda
Public Sub Buscar(Opcion As Integer, Optional Cliente As String = "", Optional Desde As String = "", Optional Hasta As String = "")
Dim rcBusqueda As New ADODB.Recordset
Dim i As Long
Dim diasComercializacion As Integer
Dim FechaComercializacion As Date
IDCliente = Opcion

On Error GoTo Error

    Screen.MousePointer = vbHourglass
  
  
    rcBusqueda.Open "SELECT * FROM clientes WHERE ID=" & Opcion, dbDatos, adOpenForwardOnly, adLockOptimistic
    
  
    If rcBusqueda.BOF Or rcBusqueda.EOF Then
         IDCliente = 0
        MsgBox "No se encontró información relacionada con el cliente !!", vbInformation, "Reasignar Tarjeta"
    
    Else
    
   
        With rcBusqueda
            While Not .EOF
                txtNombre.text = !Nombre
                txtApellidoPaterno.text = IIf(!ApellidoPaterno = "" And !ApellidoMaterno = "", !Apellido, !ApellidoPaterno)
                txtApellidoMaterno.text = !ApellidoMaterno
                txtDireccion.text = !Direccion & IIf(!NoExterior <> "", " #" & !NoExterior, "") & IIf(!NoInterior <> "", " INT." & !NoInterior, "") & " COL." & !Colonia & " C.P." & !CP
                txtCuidad.text = IIf(IsNull(!Municipio), "", !Municipio)
                txtRFC.text = IIf(IsNull(!RFC), "", !RFC)
                GetTargeta !ID
            .MoveNext
            Wend
            
        End With
    
        
    End If
    rcBusqueda.Close
    Set rcBusqueda = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcBusqueda = Nothing
     IDCliente = 0
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub GetTargeta(ID As Long)
Dim rcConsulta As New ADODB.Recordset

On Error GoTo Error
    rcConsulta.Open "SELECT ID,NumeroTarjeta,Puntos FROM asignaciontarjetas where  IDCliente = " & ID & " and IDTarjeta = 1", dbDatos, adOpenForwardOnly, adLockOptimistic
    
    
    If rcConsulta.BOF Or rcConsulta.EOF Then
        IDCliente = 0
        MsgBox "El Cliente no tiene Tarjeta de Cliente Frecuente asignada !!", vbInformation, "Reasignar Tarjeta"
    
    Else
    
   
        With rcConsulta
            While Not .EOF
                txtTarjetaAnterior.text = !NumeroTarjeta
                lblPuntos.Caption = !Puntos
            .MoveNext
            Wend
            
        End With
    
        
    End If
    rcConsulta.Close
    
    Set rcConsulta = Nothing

Error:

    If Err > 0 Then Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Private Sub TabControl1_BeforeClick(ByVal lTab As Long, bCancel As Boolean)

End Sub

Private Sub Label1_Click()

End Sub

