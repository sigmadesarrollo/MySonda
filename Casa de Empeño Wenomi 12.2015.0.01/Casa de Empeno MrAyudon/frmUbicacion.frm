VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmUbicacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de ubicación"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUbicacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   7995
   Begin Line3D.ucLine3D ucLine3D7 
      Height          =   1005
      Index           =   0
      Left            =   2925
      Top             =   1290
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   1773
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D6 
      Height          =   30
      Left            =   675
      Top             =   1530
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D5 
      Height          =   975
      Left            =   7485
      Top             =   1290
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   1720
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D4 
      Height          =   30
      Left            =   660
      Top             =   1275
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   1020
      Left            =   660
      Top             =   1275
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   1799
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   45
      Left            =   660
      Top             =   2265
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   79
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1020
      Left            =   675
      Top             =   1305
      Width           =   15
      _ExtentX        =   26
      _ExtentY        =   1799
      Orientation     =   0
      LineWidth       =   2
   End
   Begin VB.TextBox txtFila 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   5280
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtCajon 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   3000
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtCaja 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   720
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtDireccion 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1095
      MaxLength       =   30
      TabIndex        =   3
      Top             =   960
      Width           =   6720
   End
   Begin VB.TextBox txtNombre 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1095
      MaxLength       =   20
      TabIndex        =   1
      Top             =   600
      Width           =   2670
   End
   Begin VB.TextBox txtApellidos 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4845
      MaxLength       =   60
      TabIndex        =   2
      Top             =   600
      Width           =   2970
   End
   Begin VB.TextBox txtFolio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1095
      TabIndex        =   0
      Top             =   270
      Width           =   975
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   375
      Left            =   2205
      TabIndex        =   11
      Top             =   180
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16777215
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmUbicacion.frx":000C
   End
   Begin Line3D.ucLine3D ucLine3D7 
      Height          =   1005
      Index           =   1
      Left            =   5205
      Top             =   1275
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   1773
      Orientation     =   0
      LineWidth       =   2
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   6645
      TabIndex        =   16
      Top             =   2415
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
      Picture         =   "frmUbicacion.frx":0391
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   5475
      TabIndex        =   17
      Top             =   2415
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      Object.ToolTipText     =   ""
      Picture         =   "frmUbicacion.frx":08E3
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fila"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6217
      TabIndex        =   14
      Top             =   1275
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cajón"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3810
      TabIndex        =   13
      Top             =   1275
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1590
      TabIndex        =   12
      Top             =   1275
      Width           =   420
   End
   Begin VB.Label Label121 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   937
      Width           =   960
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   577
      Width           =   810
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "Apellidos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   8
      Top             =   570
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contrato:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   945
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000013&
      Height          =   255
      Left            =   690
      TabIndex        =   15
      Top             =   1305
      Width           =   6795
   End
End
Attribute VB_Name = "frmUbicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez

Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()

On Error GoTo error
    
    If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Registro de ubicación") = vbYes Then
        dbDatos.Execute "UPDATE empeno SET Caja='" & Trim(txtCaja.text) & "',Cajon='" & Trim(txtCajon.text) & "',Fila='" & Trim(txtFila.text) & "' WHERE ID=" & Val(txtFolio.Tag)
        Limpiar
        txtFolio.SetFocus
    End If
    Exit Sub
    
error:
    Maneja_Error Err
End Sub

Private Sub cmdBuscar_Click()
Dim rcConsulta As New ADODB.Recordset

On Error GoTo error
    
    With rcConsulta
    
        .Open "SELECT empeno.ID,empeno.Caja,empeno.Cajon,empeno.Fila,clientes.Nombre,clientes.Apellido,clientes.Direccion FROM empeno LEFT JOIN clientes ON empeno.IDCliente=clientes.ID WHERE empeno.NumContrato=" & Val(txtFolio.text) & " AND empeno.Pagado=0 AND empeno.Cancelado=0 AND (Serie=" & SERIE_A & " OR Serie=" & SERIE_C & ")", dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not .BOF And Not .EOF Then
            
            txtFolio.Tag = !ID
            txtNombre.text = !Nombre
            txtApellidos.text = !apellido
            txtDireccion.text = !Direccion
            txtCaja.text = IIf(IsNull(!caja), "", !caja)
            txtCajon.text = IIf(IsNull(!Cajon), "", !Cajon)
            txtFila.text = IIf(IsNull(!Fila), "", !Fila)
            
        Else
            
            MsgBox "No se encontró el contrato especificado !!", vbInformation, "Registro de ubicación"
            txtFolio.SetFocus
        End If
        .Close
        Set rcConsulta = Nothing
    End With
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl()
End Sub

Private Sub txtApellidos_GotFocus()
    Seleccionar_Texto txtApellidos
    Cambiar_Color True, txtApellidos
End Sub

Private Sub txtApellidos_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = 0
End Sub

Private Sub txtApellidos_LostFocus()
    Cambiar_Color False, txtApellidos
End Sub

Private Sub txtCaja_GotFocus()
    Cambiar_Color True, txtCaja
    Seleccionar_Texto txtCaja
End Sub

Private Sub txtCaja_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCaja_LostFocus()
    Cambiar_Color False, txtCaja
End Sub

Private Sub txtCajon_GotFocus()
    Cambiar_Color True, txtCajon
    Seleccionar_Texto txtCaja
End Sub

Private Sub txtCajon_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCajon_LostFocus()
    Cambiar_Color False, txtCajon
End Sub

Private Sub txtDireccion_GotFocus()
    Seleccionar_Texto txtDireccion
    Cambiar_Color True, txtDireccion
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = 0
End Sub

Private Sub txtDireccion_LostFocus()
    Cambiar_Color False, txtDireccion
End Sub

Private Sub txtFila_GotFocus()
    Cambiar_Color True, txtFila
    Seleccionar_Texto txtCaja
End Sub

Private Sub txtFila_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFila_LostFocus()
    Cambiar_Color False, txtFila
End Sub

Private Sub txtFolio_GotFocus()
    Seleccionar_Texto txtFolio
    Cambiar_Color True, txtFolio
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdBuscar_Click
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFolio_LostFocus()
    Cambiar_Color False, txtFolio
End Sub

Private Sub txtNombre_GotFocus()
    Seleccionar_Texto txtNombre
    Cambiar_Color True, txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = 0
End Sub

Private Sub txtNombre_LostFocus()
    Cambiar_Color False, txtNombre
End Sub

Sub Limpiar()
    txtFolio.text = ""
    txtFolio.Tag = ""
    txtNombre.text = ""
    txtApellidos.text = ""
    txtDireccion.text = ""
    txtCaja.text = ""
    txtCajon.text = ""
    txtFila.text = ""
End Sub
