VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmMensajes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Mensajes"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   5910
   Begin VB.CheckBox chkAutomovil 
      Appearance      =   0  'Flat
      Caption         =   "Automóvil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   240
      Width           =   1530
   End
   Begin VB.TextBox txtMensaje 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtFolio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   375
      Left            =   4170
      TabIndex        =   10
      Top             =   165
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
      Picture         =   "frmMensajes.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4740
      TabIndex        =   12
      Top             =   2865
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
      Picture         =   "frmMensajes.frx":0391
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3570
      TabIndex        =   13
      Top             =   2865
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
      ShadowColor     =   4210752
      TextColor       =   4210752
      Object.ToolTipText     =   ""
      Picture         =   "frmMensajes.frx":08E3
   End
   Begin VB.Label lblDireccion 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   9
      Top             =   1680
      Width           =   4275
   End
   Begin VB.Label lblApellidos 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   8
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   7
      Top             =   720
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Apellidos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1110
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mensaje:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contrato:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1050
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()

    If txtFolio.Tag = "" Then
        
        MsgBox "Seleccione el contrato de la boleta !!", vbInformation, "Mensajes"
        txtFolio.SetFocus
    Else
        
        dbDatos.Execute "UPDATE empeno SET Notas='" & Trim(txtMensaje.text) & "' where ID=" & Val(txtFolio.Tag)
        Limpiar
        txtMensaje.Enabled = False
        txtFolio.SetFocus
    End If

End Sub

Private Sub cmdBuscar_Click()

    If txtFolio.text <> "" Then BuscaFolio Val(txtFolio.text), IIf(chkAutomovil.Value = 0, 1, 2)

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
    txtMensaje.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtFolio_GotFocus()
    Seleccionar_Texto txtFolio
    Cambiar_Color True, txtFolio
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And txtFolio.text <> "" Then BuscaFolio Val(txtFolio.text), IIf(chkAutomovil.Value = 0, 1, 2)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFolio_LostFocus()
    Cambiar_Color False, txtFolio
End Sub

Private Sub txtMensaje_GotFocus()
    Seleccionar_Texto txtMensaje
    Cambiar_Color True, txtMensaje
End Sub

Private Sub txtMensaje_KeyPress(KeyAscii As Integer)
    KeyAscii = mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMensaje_LostFocus()
    Cambiar_Color False, txtMensaje
End Sub

Function BuscaFolio(Folio As Long, Serie As String)
Dim rcConsulta As ADODB.Recordset

On Error GoTo error

    Set rcConsulta = dbDatos.Execute("select empeno.ID,clientes.Nombre,clientes.Apellido,concat(clientes.Direccion,' Col. ',clientes.Colonia) as Domicilio,empeno.Notas from empeno inner join clientes on empeno.IDCliente=clientes.ID where empeno.Numcontrato=" & Folio & " and empeno.Serie=" & Serie & " and empeno.Pagado=0 And empeno.Cancelado=0")

    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        
        txtFolio.Tag = rcConsulta!ID
        lblNombre.Caption = rcConsulta!Nombre
        lblApellidos.Caption = rcConsulta!apellido
        lblDireccion.Caption = rcConsulta!Domicilio
        txtMensaje.Enabled = True
        txtMensaje.text = IIf(IsNull(rcConsulta!notas), "", rcConsulta!notas)
    Else
        
        MsgBox "No se encuentró el contrato especificado !!", vbInformation, "Mensajes"
        Limpiar
        txtMensaje.Enabled = False
        txtFolio.SetFocus
    End If

error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Function

Sub Limpiar()
    txtFolio.text = ""
    txtFolio.Tag = ""
    lblNombre.Caption = ""
    lblApellidos.Caption = ""
    lblDireccion.Caption = ""
    txtMensaje.text = ""
    chkAutomovil.Value = 0
End Sub
