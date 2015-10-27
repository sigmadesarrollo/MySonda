VERSION 5.00
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros Puntos"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   6150
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picPuntos 
      Height          =   4935
      Left            =   8040
      ScaleHeight     =   4875
      ScaleWidth      =   5595
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   480
      Width           =   5655
      Begin VB.TextBox txtPuntos 
         Alignment       =   1  'Right Justify
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
         Left            =   3120
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin DevPowerFlatBttn.FlatBttn cmdAceptarPuntos 
         Height          =   375
         Left            =   4200
         TabIndex        =   27
         Top             =   4440
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
         Picture         =   "frmParametros.frx":000C
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   3120
         TabIndex        =   26
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Equivalente a Pesos"
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
         Left            =   480
         TabIndex        =   25
         Top             =   1800
         Width           =   2460
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numeros de Puntos:"
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
         Left            =   480
         TabIndex        =   23
         Top             =   1080
         Width           =   2505
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   6480
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
      Picture         =   "frmParametros.frx":055E
   End
   Begin VB.PictureBox pParametrosTarjetas 
      Height          =   5535
      Left            =   240
      ScaleHeight     =   5475
      ScaleWidth      =   5595
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5655
      Begin VB.TextBox txtpRefrendoExt 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtpAbonos 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   11
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtpApartados 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   10
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtpEmpenoAutos 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtpEmpeno 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtpRefrendo 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtpDesempeno 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtpVentas 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   1215
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
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4455
      End
      Begin DevPowerFlatBttn.FlatBttn cmdGrabar 
         Height          =   375
         Left            =   4320
         TabIndex        =   17
         Top             =   5040
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
         Picture         =   "frmParametros.frx":0AB0
      End
      Begin DevPowerFlatBttn.FlatBttn cmdDesactivar 
         Height          =   375
         Left            =   3000
         TabIndex        =   21
         Top             =   5040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         AlignCaption    =   4
         AlignPicture    =   2
         AutoSize        =   0   'False
         Caption         =   "   &Desactivar"
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
      End
      Begin VB.Label lblRefrendo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Refrendo Extemporaneo"
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
         Left            =   1440
         TabIndex        =   28
         Top             =   2400
         Width           =   3360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Abonos"
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
         Left            =   1440
         TabIndex        =   20
         Top             =   4320
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Apartados"
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
         Left            =   1440
         TabIndex        =   19
         Top             =   3840
         Width           =   1635
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Empeño Autos"
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
         Left            =   1440
         TabIndex        =   18
         Top             =   1440
         Width           =   2145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Empeños"
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
         Left            =   1440
         TabIndex        =   16
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Refrendo"
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
         Left            =   1440
         TabIndex        =   15
         Top             =   1920
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Desempeño"
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
         Left            =   1440
         TabIndex        =   14
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Ventas"
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
         Left            =   1440
         TabIndex        =   13
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1950
      End
   End
   Begin vbalDTab6.vbalDTabControl dTab 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   10610
      TabAlign        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_IDTarjeta As Long

Private Sub cmbTarjetas_Click()
   Cargar_Valores cmbTarjetas.ItemData(cmbTarjetas.ListIndex)
End Sub

Private Sub cmbTarjetas_DropDown()
   Cambiar_Color True, cmbTarjetas
End Sub

Private Sub cmbTarjetas_GotFocus()
   Seleccionar_Texto cmbTarjetas
   Cambiar_Color True, cmbTarjetas
End Sub

Private Sub cmbTarjetas_LostFocus()
   Cambiar_Color False, cmbTarjetas
   Buscar_Tarjeta
End Sub

Private Sub cmdAceptarPuntos_Click()
   Grabar_Puntos
End Sub

Private Sub cmdDesactivar_Click()
    Desactivar
End Sub

Private Sub cmdGrabar_Click()
   Grabar_Porcentaje
End Sub

Private Sub Grabar_Porcentaje()
   
On Error GoTo Error
   
   Dim Sql As String
   
   If m_IDTarjeta = 0 Then
   
      Sql = "INSERT INTO TarjetasPuntos (TipoTarjeta,pEmpeno,pEmpenoAutos,pRefrendo,pRefrendoExt,pDesempeno,pVentas,pApartados,pAbonos,FechaCreacion,Activa) VALUES ('" & _
            cmbTarjetas.Text & "'," & Val(txtpEmpeno.Text) & "," & Val(txtpEmpenoAutos.Text) & "," & Val(txtpRefrendo.Text) & "," & Val(txtpRefrendoExt.Text) & "," & Val(txtpDesempeno.Text) & "," & _
            Val(txtpVentas.Text) & "," & Val(txtpApartados.Text) & "," & Val(txtpAbonos.Text) & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',1)"
            
   Else
         
         Sql = "UPDATE TarjetasPuntos SET pEmpeno=" & Val(txtpEmpeno.Text) & _
            ",pEmpenoAutos=" & Val(txtpEmpenoAutos.Text) & _
            ",pRefrendo=" & Val(txtpRefrendo.Text) & _
            ",pRefrendoExt=" & Val(txtpRefrendoExt.Text) & _
            ",pDesempeno=" & Val(txtpDesempeno.Text) & _
            ",pVentas=" & Val(txtpVentas.Text) & _
            ",pApartados=" & Val(txtpApartados.Text) & _
            ",pAbonos=" & Val(txtpAbonos.Text) & _
            ",Activa=1 WHERE ID=" & m_IDTarjeta
            
   End If
   
   m_Conexion.Execute Sql
   
   Limpiar
   
Error:
   Maneja_Error Err
   
End Sub

Private Sub Grabar_Puntos()
   On Error GoTo Error
   
   m_Conexion.Execute "UPDATE Parametros SET PuntosTarjeta=" & Val(txtPuntos.Text)
   
Error:
   Maneja_Error Err
End Sub

Private Sub Desactivar()
   On Error GoTo Error
   
   If m_IDTarjeta <> 0 Then
      m_Conexion.Execute "UPDATE TarjetasPuntos SET Activa=0 WHERE ID=" & m_IDTarjeta
      Limpiar
   End If
   
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
   Screen.MousePointer = vbHourglass
   pParametrosTarjetas.BorderStyle = 0
   picPuntos.BorderStyle = 0
   Crear_Tabs
   Cargar_Tarjetas
   txtPuntos.Text = Val(SacaValor("parametros", "PuntosTarjeta", ""))
   Screen.MousePointer = vbDefault
End Sub

Private Sub Crear_Tabs()
   Dim pTab As cTab
   With dTab
      Set pTab = .Tabs.Add("K1", , "Parametros Tarjetas")
      pTab.Panel = pParametrosTarjetas
      Set pTab = .Tabs.Add("K2", , "Redimir Puntos")
      pTab.Panel = picPuntos
   End With
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

Private Sub Text1_Change()

End Sub

Private Sub txtpAbonos_GotFocus()
   Seleccionar_Texto txtpAbonos
   Cambiar_Color True, txtpAbonos
End Sub

Private Sub txtpAbonos_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Numeros(KeyAscii)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtpAbonos_LostFocus()
   Cambiar_Color False, txtpAbonos
End Sub

Private Sub txtpApartados_GotFocus()
   Seleccionar_Texto txtpApartados
   Cambiar_Color True, txtpApartados
End Sub

Private Sub txtpApartados_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Numeros(KeyAscii)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtpApartados_LostFocus()
   Cambiar_Color False, txtpApartados
End Sub

Private Sub txtpEmpeno_GotFocus()
   Seleccionar_Texto txtpEmpeno
   Cambiar_Color True, txtpEmpeno
End Sub

Private Sub txtpEmpeno_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Numeros(KeyAscii)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtpEmpeno_LostFocus()
   Cambiar_Color False, txtpEmpeno
End Sub

Private Sub txtpEmpenoAutos_GotFocus()
   Seleccionar_Texto txtpEmpenoAutos
   Cambiar_Color True, txtpEmpenoAutos
End Sub

Private Sub txtpEmpenoAutos_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Numeros(KeyAscii)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtpEmpenoAutos_LostFocus()
   Cambiar_Color False, txtpEmpenoAutos
End Sub

Private Sub txtpRefrendo_GotFocus()
   Seleccionar_Texto txtpRefrendo
   Cambiar_Color True, txtpRefrendo
End Sub

Private Sub txtpRefrendo_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Numeros(KeyAscii)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtpRefrendo_LostFocus()
   Cambiar_Color False, txtpRefrendo
End Sub

Private Sub txtpDesempeno_GotFocus()
   Seleccionar_Texto txtpDesempeno
   Cambiar_Color True, txtpDesempeno
End Sub

Private Sub txtpDesempeno_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Numeros(KeyAscii)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtpDesempeno_LostFocus()
   Cambiar_Color False, txtpDesempeno
End Sub

Private Sub txtpRefrendoExt_GotFocus()
    Seleccionar_Texto txtpRefrendoExt
   Cambiar_Color True, txtpRefrendoExt
End Sub

Private Sub txtpRefrendoExt_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtpRefrendoExt_LostFocus()
    Cambiar_Color False, txtpRefrendoExt
End Sub

Private Sub txtPuntos_GotFocus()
   Seleccionar_Texto txtPuntos
   Cambiar_Color True, txtPuntos
End Sub

Private Sub txtPuntos_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Numeros(KeyAscii)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtPuntos_LostFocus()
   Cambiar_Color False, txtPuntos
End Sub

Private Sub txtpVentas_GotFocus()
   Seleccionar_Texto txtpVentas
   Cambiar_Color True, txtpVentas
End Sub

Private Sub txtpVentas_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Numeros(KeyAscii)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtpVentas_LostFocus()
   Cambiar_Color False, txtpVentas
End Sub

Private Sub Cargar_Valores(IDTarjeta As Long)
   On Error GoTo Error
   Dim rc As New ADODB.Recordset
     
   
   rc.Open "SELECT * FROM TarjetasPuntos WHERE ID=" & IDTarjeta, m_Conexion, adOpenForwardOnly, adLockOptimistic
   With rc
      If Not .EOF Then
         txtpEmpeno.Text = !pEmpeno
         txtpEmpenoAutos.Text = !pEmpenoAutos
         txtpRefrendo.Text = !pRefrendo
         txtpRefrendoExt.Text = !pRefrendoExt
         txtpDesempeno.Text = !pDesempeno
         txtpVentas.Text = !pVentas
         txtpApartados.Text = !pApartados
         txtpAbonos.Text = !pAbonos
      Else
         txtpEmpeno.Text = "0"
         txtpEmpenoAutos.Text = "0"
         txtpRefrendo.Text = "0"
         txtpRefrendoExt.Text = "0"
         txtpDesempeno.Text = "0"
         txtpVentas.Text = "0"
         txtpApartados.Text = "0"
         txtpAbonos.Text = "0"
      End If
   End With
      
   rc.Close
Error:
   Maneja_Error Err
   
   Set rc = Nothing
End Sub

Private Sub Limpiar()
   m_IDTarjeta = 0
   Cargar_Tarjetas
   txtpEmpeno.Text = "0"
   txtpEmpenoAutos.Text = "0"
   txtpRefrendo.Text = "0"
   txtpRefrendoExt.Text = "0"
   txtpDesempeno.Text = "0"
   txtpVentas.Text = "0"
   txtpApartados.Text = "0"
   txtpAbonos.Text = "0"
End Sub

Private Sub Buscar_Tarjeta()
   Dim Indice As Integer
   
   m_IDTarjeta = 0
   For Indice = 0 To cmbTarjetas.ListCount - 1
      If cmbTarjetas.Text = cmbTarjetas.List(Indice) Then
         m_IDTarjeta = cmbTarjetas.ItemData(Indice)
         Exit For
      End If
   Next Indice
   
   Cargar_Valores m_IDTarjeta
   
End Sub
