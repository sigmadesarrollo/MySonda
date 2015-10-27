VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmCambioPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Plan"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCambioPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   5820
   Begin vbAcceleratorGrid6.vbalGrid grdContratos 
      Height          =   2295
      Left            =   0
      TabIndex        =   24
      Top             =   3360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4048
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
   End
   Begin Line3D.ucLine3D ucLine3D8 
      Height          =   1155
      Left            =   1200
      Top             =   2100
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   2037
      Orientation     =   0
      LineWidth       =   2
   End
   Begin VB.TextBox txtNumContrato 
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
      Top             =   195
      Width           =   1095
   End
   Begin VB.ComboBox cmbTipoInteres 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCambioPlan.frx":000C
      Left            =   1335
      List            =   "frmCambioPlan.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2910
      Width           =   1560
   End
   Begin VB.ComboBox cmbPeriodo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCambioPlan.frx":0010
      Left            =   3045
      List            =   "frmCambioPlan.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2910
      Width           =   1455
   End
   Begin VB.ComboBox cmbPlazos 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCambioPlan.frx":0014
      Left            =   4680
      List            =   "frmCambioPlan.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2910
      Width           =   870
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3435
      TabIndex        =   1
      Top             =   5760
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
      TextColor       =   4210752
      Object.ToolTipText     =   ""
      Picture         =   "frmCambioPlan.frx":0018
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4635
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
      Picture         =   "frmCambioPlan.frx":056A
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   765
      Index           =   0
      Left            =   90
      Top             =   2505
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   1349
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D7 
      Height          =   30
      Index           =   0
      Left            =   90
      Top             =   3255
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D6 
      Height          =   1170
      Index           =   0
      Left            =   2955
      Top             =   2100
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2064
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D4 
      Height          =   30
      Left            =   90
      Top             =   2880
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   30
      Left            =   90
      Top             =   2505
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Left            =   1200
      Top             =   2100
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1515
      Index           =   1
      Left            =   -45
      Top             =   1455
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2672
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D6 
      Height          =   1170
      Index           =   1
      Left            =   4605
      Top             =   2100
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2064
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1155
      Index           =   2
      Left            =   5625
      Top             =   2115
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2037
      Orientation     =   0
      LineWidth       =   2
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   150
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
      Picture         =   "frmCambioPlan.frx":0ABC
   End
   Begin VB.Label lblCambio 
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   5760
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblTotal 
      Caption         =   "0.00"
      Height          =   255
      Left            =   1920
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblVencimiento 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblSeguro 
      Caption         =   "0.00"
      Height          =   255
      Left            =   3840
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblAlmacenaje 
      Caption         =   "0.00"
      Height          =   255
      Left            =   3840
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblTasa 
      Caption         =   "0.00"
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblPrestamo 
      Caption         =   "0.00"
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblFecha 
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
      TabIndex        =   23
      Top             =   600
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
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
      Index           =   6
      Left            =   120
      TabIndex        =   22
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label6 
      Height          =   405
      Left            =   90
      TabIndex        =   21
      Top             =   2100
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   270
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label7 
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
      TabIndex        =   19
      Top             =   1560
      Width           =   1110
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
      TabIndex        =   18
      Top             =   1080
      Width           =   75
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
      TabIndex        =   17
      Top             =   1560
      Width           =   75
   End
   Begin VB.Label lblPlazos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4680
      TabIndex        =   16
      Top             =   2565
      Width           =   870
   End
   Begin VB.Label lblPeriodo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3045
      TabIndex        =   15
      Top             =   2565
      Width           =   1455
   End
   Begin VB.Label lblTipoInteres 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1335
      TabIndex        =   14
      Top             =   2565
      Width           =   1560
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   225
      TabIndex        =   13
      Top             =   2955
      Width           =   870
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   225
      TabIndex        =   12
      Top             =   2565
      Width           =   870
   End
   Begin VB.Label Label5 
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
      TabIndex        =   11
      Top             =   195
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Interés"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1470
      TabIndex        =   8
      Top             =   2175
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3105
      TabIndex        =   7
      Top             =   2175
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Plazo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4605
      TabIndex        =   6
      Top             =   2175
      Width           =   1020
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Height          =   1140
      Left            =   105
      TabIndex        =   9
      Top             =   2130
      Width           =   5535
   End
End
Attribute VB_Name = "frmCambioPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbPeriodo_Click()
    
    If cmbPeriodo.ListIndex > -1 Then
        
        Cargar_Combos "DISTINCT plazos.Descripcion", "configuraciontasas INNER JOIN plazos ON plazos.ID=configuraciontasas.IDPlazo", cmbPlazos, " WHERE configuraciontasas.IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex) & " AND configuraciontasas.IDTipoPeriodo=" & cmbPeriodo.ItemData(cmbPeriodo.ListIndex), "plazos.Descripcion", , "plazos.ID"
    
    End If
    
    If cmbPlazos.ListCount > 0 Then cmbPlazos.ListIndex = 0
End Sub

Private Sub cmbPeriodo_GotFocus()
    Cambiar_Color True, cmbPeriodo
End Sub

Private Sub cmbPeriodo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPeriodo_LostFocus()
    Cambiar_Color False, cmbPeriodo
End Sub

Private Sub cmbPlazos_Click()
    SacaTasa CCur(lblPrestamo.Caption), cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), cmbPeriodo.ItemData(cmbPeriodo.ListIndex), cmbPlazos.ItemData(cmbPlazos.ListIndex), True
End Sub

Private Sub cmbPlazos_GotFocus()
    Cambiar_Color True, cmbPlazos
End Sub

Private Sub cmbPlazos_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPlazos_LostFocus()
    Cambiar_Color False, cmbPlazos
End Sub

Private Sub cmbTipoInteres_Click()
        
    cmbPeriodo.Clear
    cmbPlazos.Clear
    If cmbTipoInteres.ListIndex > -1 Then
                
        Cargar_Combos "DISTINCT tipoperiodo.Descripcion", "configuraciontasas INNER JOIN tipoperiodo ON tipoperiodo.ID=configuraciontasas.IDTipoPeriodo", cmbPeriodo, " WHERE configuraciontasas.IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), "tipoperiodo.Ordenamiento", , "tipoperiodo.ID"
    
    End If
    
    If cmbPeriodo.ListCount > 0 Then cmbPeriodo.ListIndex = 0
End Sub

Private Sub cmbTipoInteres_GotFocus()
    Cambiar_Color True, cmbTipoInteres
End Sub

Private Sub cmbTipoInteres_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoInteres_LostFocus()
    Cambiar_Color False, cmbTipoInteres
End Sub

Private Sub cmdAceptar_Click()
Dim crPrestamo As Double, Folio As Long, FolioAnterior As Long, Movimiento As Long, FolioNota As Long, strIniciales As String
Dim IDEmpeno As Long, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crIva As Double, strSql1 As String, strSql2 As String, strSql3 As String, Dias As Integer, Vencimiento As Date, Periodo As Long, Meses As Integer, Tasa As Double, Almacenaje As Double, Seguro As Double, crImporteTotal As Double, crEfectivo As Double, Hora As String
Dim rcEmpeño As New ADODB.Recordset
Dim rcArticulos As New ADODB.Recordset

On Error GoTo Error

    If Val(txtNumContrato.text) > 0 And (cmbTipoInteres.ListIndex > -1 And cmbPeriodo.ListIndex > -1 And cmbPlazos.ListIndex > -1) Then
        
        If MsgBox("Esta seguro que desea realizar el cambio de plan ??", vbQuestion + vbYesNo + vbDefaultButton2, "Cambio de plan") = vbYes Then
                
            crImporteTotal = CDbl(lblTotal.Caption)
            crEfectivo = frmEfectivo.RegresaCambio(crImporteTotal, 2)
            If crEfectivo < crImporteTotal Then Exit Sub
            CalculaCambio crEfectivo, crImporteTotal, 2
    
            'Tomo el Periodo
            Periodo = SacaValor("tipoperiodo", "Periodo", " WHERE ID=" & cmbPeriodo.ItemData(cmbPeriodo.ListIndex))
            
            'Saco las Iniciales
            strIniciales = Trim(lblNombre.Tag)
            
            'Saco el Nuevo Folio
            Folio = Regresa_NumContrato(False, SERIE_C)
            Regresa_NumContrato True, SERIE_C
            
            'Saco el Movimiento
            Movimiento = Regresa_Movimiento(False)
            Regresa_Movimiento True
            
            'Folio Notas
            FolioNota = Regresa_Movimiento(False, "FolioNotas")
            Regresa_Movimiento True, "FolioNotas"
            
            'Tomo los valores
            Tasa = CDbl(Mid(lblTasa.Caption, 1, Len(lblTasa.Caption) - 1))
            Almacenaje = CDbl(Mid(lblAlmacenaje.Caption, 1, Len(lblAlmacenaje.Caption) - 1))
            Seguro = CDbl(Mid(lblSeguro.Caption, 1, Len(lblSeguro.Caption) - 1))
                
            crIntereses = CDbl(grdContratos.CellText(grdContratos.Rows - 1, 4))
            crAlmacenaje = CDbl(grdContratos.CellText(grdContratos.Rows - 1, 5))
            crSeguro = CDbl(grdContratos.CellText(grdContratos.Rows - 1, 5))
            crIva = CDbl(grdContratos.CellText(grdContratos.Rows - 1, 7))
            
            'Leemos los datos del Empeno original
            rcEmpeño.Open "SELECT empeno.* FROM empeno WHERE empeno.ID=" & Val(txtNumContrato.Tag), dbDatos, adOpenForwardOnly, adLockOptimistic
            FolioAnterior = rcEmpeño!Folio
            crPrestamo = rcEmpeño!Prestamo
            
            'Actualizo el Empeño
            dbDatos.Execute "UPDATE empeno SET PC='" & NombrePc & "',Destino=" & OD_REFRENDO & ",FolioDestino=" & Folio & ",Pagado=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',IDUsuarioMov=" & frmMDI.IDUsuario & ",Intereses=" & crIntereses & ",Importeiva=" & crIva & ",ImporteAlmacenaje=" & crAlmacenaje & ",ImporteSeguro=" & crSeguro & ",Efectivo=" & crEfectivo & ",FolioNota=" & FolioNota & " WHERE ID=" & Val(txtNumContrato.Tag)
                                                       
            Select Case cmbPeriodo.text
            Case "MENSUAL"
                
                Meses = 1
            Case "QUINCENAL"
                
                Meses = 2
            Case "SEMANAL"
                
                Meses = 4
            End Select
        
            'Calculo la nuevo fecha de vencimiento
            Dias = Val(Periodo) * Val(cmbPlazos.text)
            Vencimiento = IIf(cmbTipoInteres.text = "FIJA", DateAdd("M", 2, Date), IIf(Periodo = 30, DateAdd("M", Val(cmbPlazos.text), Date), DateAdd("D", (Periodo * Val(cmbPlazos.text)) - 1, Date)))
            
            'Grabo el nuevo Empeño
            strSql1 = "INSERT INTO empeno (Fecha,Movimiento,NumContrato,Folio,Prestamo,PrestamoInicial,Avaluo,Origen,Vencimiento,FolioOrigen,Serie,PC,IDCliente,Responsable,Valuador,Notas,Tasa,Almacenaje,Seguro,Operacion,Comision,IVA,Periodo,Venperiodo,VenAlmoneda,Tipointeres,TipoTasa,IDSucursal,IDUsuario,IDAutorizacion,NumBolsa,Ubicacion,Caja,Cajon,Fila,IDUsuarioAutoriza,TipoAutoriza) VALUES "
            strSql2 = "('" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & rcEmpeño!NumContrato & "," & Folio & "," & crPrestamo & "," & crPrestamo & "," & rcEmpeño!Avaluo & "," & OD_REFRENDO & ",'" & Format(Vencimiento, "YYYY/MM/DD") & "'," & rcEmpeño!Folio & "," & IIf(cmbTipoInteres.text = "FIJA", SERIE_C, SERIE_A) & ",'" & Trim(NombrePc) & "'," & rcEmpeño!IDCliente & ",'" & Trim(rcEmpeño!Responsable) & "',"
            strSql3 = "'" & Trim(rcEmpeño!valuador) & "','" & Trim(rcEmpeño!Notas) & "'," & Tasa & "," & Almacenaje & "," & Seguro & "," & rcEmpeño!Operacion & "," & rcEmpeño!Comision & "," & rcEmpeño!Iva & "," & Periodo & "," & cmbPlazos.text & "," & rcEmpeño!VenAlmoneda & ",'" & cmbTipoInteres.text & "','" & cmbPeriodo.text & "'," & rcEmpeño!IDSucursal & "," & rcEmpeño!IDUsuario & "," & rcEmpeño!IDAutorizacion & ",'" & Trim(rcEmpeño!NumBolsa) & "','" & Trim(rcEmpeño!ubicacion) & "','" & Trim(rcEmpeño!caja) & "','" & Trim(rcEmpeño!Cajon) & "','" & Trim(rcEmpeño!Fila) & "'," & rcEmpeño!IDUsuarioAutoriza & "," & rcEmpeño!TipoAutoriza & ")"
            
            dbDatos.Execute strSql1 & strSql2 & strSql3
            
            rcEmpeño.Close
            
            'Saco el ID del Empeño
            IDEmpeno = SacaValor("empeno", "MAX(ID)")
            
            'Grabo el Detalle del Empeño
            rcArticulos.Open "SELECT * FROM detallesempeno WHERE IDEmpeno=" & Val(txtNumContrato.Tag), dbDatos, adOpenForwardOnly, adLockReadOnly
            With rcArticulos
                If Not rcArticulos.BOF And Not rcArticulos.EOF Then
                    
                    While Not .EOF
                                                                                    
                        dbDatos.Execute "INSERT INTO detallesempeno (IDEmpeno,Codigo,Tipo,Cantidad,Articulo,Peso,Kilates,Avaluo,Prestamo,Origen,CantidadPiedras,PesoPiedras,CantidadDiamantes,Puntos,PrestamoDiamante,Observaciones,TipoPrenda,Estado,Marca,Modelo,Serie,Color,Tamano) VALUES (" & _
                                        IDEmpeno & ",'" & Trim(!Codigo) & "'," & !Tipo & "," & !Cantidad & ",'" & !Articulo & "'," & !Peso & "," & !Kilates & "," & !Avaluo & "," & !Prestamo & ",1," & _
                                        !CantidadPiedras & "," & !PesoPiedras & "," & !CantidadDiamantes & "," & !Puntos & "," & !PrestamoDiamante & ",'" & Trim(!Observaciones) & "'," & !TipoPrenda & ",'" & !Estado & "','" & !Marca & "','" & !Modelo & "','" & !Serie & "','" & !Color & "','" & !Tamano & "')"
                    .MoveNext
                    Wend
                End If
            End With
            rcArticulos.Close
            
            'Tomo la Hora
            Hora = Time
            
            'Cuenta de Intereses
            If crIntereses > 0 Then
                
                'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & strIniciales & "','110101'," & crIntereses & "," & TIPO_CARGO & "," & SERIE_A & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & strIniciales & "','520450'," & crIntereses & "," & TIPO_ABONO & "," & SERIE_A & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            End If
            
            
            'Cuenta de Almacenaje
            If crAlmacenaje > 0 Then
                
                'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & strIniciales & "','110101'," & crAlmacenaje & "," & TIPO_CARGO & "," & SERIE_A & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & strIniciales & "','670350'," & crAlmacenaje & "," & TIPO_ABONO & "," & SERIE_A & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            End If
            
            
            'Cuenta de Seguro
            If crSeguro > 0 Then
                
                'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & strIniciales & "','110101'," & crSeguro & "," & TIPO_CARGO & "," & SERIE_A & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & strIniciales & "','680350'," & crSeguro & "," & TIPO_ABONO & "," & SERIE_A & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            End If
            
            
            'Cuenta de Iva
            If crIva > 0 Then
            
                'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & strIniciales & "','110101'," & crIva & "," & TIPO_CARGO & "," & SERIE_A & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & strIniciales & "','120150'," & crIva & "," & TIPO_ABONO & "," & SERIE_A & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            End If
                                            
'''            'Grabamos el cargo a 199401
'''            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
'''                            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & strIniciales & "','199401'," & crIntereses + crAlmacenaje + crSeguro + crIva & "," & TIPO_CARGO & "," & SERIE_A & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                        
            'Ponemos la entrada del nuevo Folio
            'Grabamos el cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & Folio & ",'" & strIniciales & "','201701'," & crPrestamo & "," & TIPO_CARGO & "," & SERIE_A & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                        
            'Grabamos el abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & strIniciales & "','201750'," & crPrestamo & "," & TIPO_ABONO & "," & SERIE_A & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            'Imprimo la boleta
            Imprimir_Boleta_CR IDEmpeno
                        
            'Imprimo el Recibo
            Imprimir_Nota Val(txtNumContrato.Tag), D_DESEMPEÑO
            
            Limpiar
        End If
        
    End If

Error:
    Maneja_Error Err
    Set rcArticulos = Nothing
    Set rcEmpeño = Nothing
End Sub

Private Sub cmdBuscar_Click()

    If Trim(txtNumContrato.text) = "" Then
        
        MsgBox "Introduzca el número de contrato !!", vbInformation, "Cambio de Plan"
        txtNumContrato.SetFocus
    Else
        
        Busca_Contrato CLng(txtNumContrato.text)
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Crear_Encabezados
    Cargar_Combos "Descripcion", "tipointeres", cmbTipoInteres, " WHERE Serie=" & SERIE_A, "Ordenamiento"
    Poner_Flat Fl(), Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtNumcontrato_GotFocus()
    Seleccionar_Texto txtNumContrato
    Cambiar_Color True, txtNumContrato
    
    If lblCambio.Tag = "1" Then
        
        lblCambio.Tag = ""
        lblCambio.Caption = ""
        lblCambio.Visible = False
    End If
End Sub

Private Sub txtNumcontrato_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn And Trim(txtNumContrato.text) <> "" Then
        Busca_Contrato CLng(txtNumContrato.text)
    End If
End Sub

Private Sub txtNumcontrato_LostFocus()
    Cambiar_Color False, txtNumContrato
End Sub

Sub Busca_Contrato(NumContrato As Long)
Dim rcConsulta As New ADODB.Recordset

On Error GoTo Error
    
    Limpiar False
    grdContratos.Clear
        
    rcConsulta.Open "SELECT empeno.ID,empeno.Folio,empeno.Prestamo,empeno.Fecha,empeno.TipoInteres,empeno.TipoTasa,empeno.VenPeriodo,empeno.Tasa,empeno.Almacenaje,empeno.Seguro,clientes.Nombre,clientes.Apellido,clientes.Iniciales FROM empeno INNER JOIN clientes ON empeno.IDCliente=clientes.ID WHERE empeno.NumContrato=" & NumContrato & " AND (empeno.Serie=" & SERIE_A & " OR empeno.Serie=" & SERIE_C & ") AND empeno.Destino=0 AND empeno.Pagado=0 And empeno.Cancelado=0", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        
        txtNumContrato.Tag = rcConsulta!ID
        lblFecha.Caption = Format(rcConsulta!Fecha, "DD/MMM/YYYY")
        lblNombre.Caption = rcConsulta!Nombre
        lblNombre.Tag = rcConsulta!Iniciales
        lblApellidos.Caption = rcConsulta!Apellido
        lblTipoInteres.Caption = rcConsulta!TipoInteres
        lblPeriodo.Caption = rcConsulta!TipoTasa
        lblPlazos.Caption = rcConsulta!VenPeriodo
        lblPrestamo.Caption = rcConsulta!Prestamo
        lblTasa.Caption = rcConsulta!Tasa
        lblAlmacenaje.Caption = rcConsulta!Almacenaje
        lblSeguro.Caption = rcConsulta!Seguro
        DetalleContrato rcConsulta!ID, rcConsulta!Folio, rcConsulta!Prestamo
    Else
        
        MsgBox "No se encuentró el contrato especificado !!", vbInformation, "Cambio de Plan"
        txtNumContrato.SetFocus
    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Sub DetalleContrato(IDEmpeno As Long, Folio As Long, crPrestamo As Double)
Dim crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crIva As Double
Dim rcAux As New ADODB.Recordset
    
    crIntereses = 0: crAlmacenaje = 0: crSeguro = 0: crIva = 0
    
    rcAux.Open "SELECT Fecha,Avaluo,Vencimiento,TipoTasa FROM empeno WHERE ID=" & IDEmpeno, dbDatos, adOpenForwardOnly, adLockOptimistic
    
    If Not rcAux.BOF And Not rcAux.EOF Then
        crIntereses = GeneraIntereses(IDEmpeno, "Tasa")
        crAlmacenaje = GeneraIntereses(IDEmpeno, "Almacenaje")
        crSeguro = GeneraIntereses(IDEmpeno, "Seguro")
        crIva = Redondeo(Regresa_Iva(crIntereses + crAlmacenaje + crSeguro, IDEmpeno))
    End If
    
    rcAux.Close
    Set rcAux = Nothing
    
    With grdContratos
        .AddRow
        .CellText(.Rows, 1) = Folio
        .CellText(.Rows, 2) = crPrestamo
        .CellTextAlign(.Rows, 2) = DT_RIGHT
        .CellText(.Rows, 3) = crIntereses + crAlmacenaje + crSeguro + crIva
        .CellTextAlign(.Rows, 3) = DT_RIGHT
        
        .CellText(.Rows, 4) = crIntereses
        .CellTextAlign(.Rows, 4) = DT_RIGHT
        .CellText(.Rows, 5) = crAlmacenaje
        .CellTextAlign(.Rows, 5) = DT_RIGHT
        .CellText(.Rows, 6) = crSeguro
        .CellTextAlign(.Rows, 6) = DT_RIGHT
        .CellText(.Rows, 7) = crIva
        .CellTextAlign(.Rows, 7) = DT_RIGHT
        If .Rows = 1 Then .AddRow
        Poner_Totales
    End With
    
End Sub

Sub Limpiar(Optional Ban As Boolean = True)
    If Ban Then txtNumContrato.text = "": txtNumContrato.Tag = ""
    lblFecha.Caption = ""
    lblNombre.Caption = ""
    lblNombre.Tag = ""
    lblApellidos.Caption = ""
    lblTipoInteres.Caption = ""
    lblPeriodo.Caption = ""
    lblPlazos.Caption = ""
    lblPrestamo.Caption = "0.00"
    lblTasa.Caption = "0.00"
    lblAlmacenaje.Caption = "0.00"
    lblSeguro.Caption = "0.00"
    lblTotal.Caption = "0.00"
    cmbTipoInteres.ListIndex = -1
    cmbPeriodo.ListIndex = -1
    cmbPlazos.ListIndex = -1
    grdContratos.Clear
End Sub

Sub Poner_Totales()
Dim columna As Integer, crIntereses As Double, i As Integer

    For i = 1 To grdContratos.Rows - 1
        
        crIntereses = crIntereses + grdContratos.CellText(i, 3)
    Next i
    
    grdContratos.CellText(grdContratos.Rows, 3) = crIntereses
    grdContratos.CellTextAlign(grdContratos.Rows, 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
    
    For i = 1 To grdContratos.Columns
        
        grdContratos.CellBackColor(grdContratos.Rows, i) = RGB(223, 208, 102)
        grdContratos.CellForeColor(grdContratos.Rows, i) = &HFF0000
    Next i
    
    lblTotal.Caption = crIntereses
End Sub

Sub Crear_Encabezados()
    
    With grdContratos
        .AddColumn "C1", "Folio", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortNumeric
        .AddColumn "C2", "Préstamo", ecgHdrTextALignRight, , 150, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C3", "Interés", ecgHdrTextALignRight, , 120, , , , , FMoneda, , CCLSortNumeric
        
        .AddColumn "C4", "Intereses", ecgHdrTextALignRight, , 120, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "C5", "Almacenaje", ecgHdrTextALignRight, , 120, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "C6", "Seguro", ecgHdrTextALignRight, , 120, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "C7", "Iva", ecgHdrTextALignRight, , 120, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "C8", "Total", ecgHdrTextALignRight, , 120, False, , , , FMoneda, , CCLSortNumeric
    End With

End Sub

Sub Imprimir_Nota(IDEmpeno As Long, Opcion As Integer)

Dim ImprDefault As Boolean

On Error GoTo Error
    
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\Nota.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.ID}=" & IDEmpeno & ""
        .Formulas(0) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(1) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
        .Formulas(2) = "Opcion=" & Opcion & ""
        .Formulas(3) = "Usuario='" & SacaValor("usuarios", "Nombre", " WHERE ID=" & frmMDI.IDUsuario) & "'"
        
        .SubreportToChange = "NuevosPagos"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .Formulas(0) = "Enajenacion=" & Regresa_Valor_BD("DiasEnajenacion") & ""
        .Formulas(1) = "LeyendaPres='PRÉSTAMO'"
        .DiscardSavedData = True
        .WindowState = crptMaximized
        .Destination = crptToWindow
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
                
        .WindowTitle = "Recibo"
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
    
'''''On Error GoTo error
'''''
'''''    With frmMDI.Cr
'''''        .Reset
'''''        .DiscardSavedData = True
'''''        .WindowShowPrintSetupBtn = True
'''''        .ReportFileName = Path & "\Reportes\Nota.rpt"
'''''        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'''''        .SelectionFormula = "{empeno.ID}=" & IDEmpeno & ""
'''''        .Formulas(0) = "Caja='" & Trim(UCase(NombrePc)) & "'"
'''''        .Formulas(1) = "Usuario='" & Trim(UCase(frmMDI.Usuario)) & "'"
'''''        .Formulas(2) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
'''''        .Formulas(3) = "Opcion=" & Opcion & ""
'''''
'''''        .SubreportToChange = "NuevosPagos"
'''''        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'''''        .Formulas(0) = "Enajenacion=" & Regresa_Valor_BD("DiasEnajenacion") & ""
'''''        .DiscardSavedData = True
'''''        .WindowState = crptMaximized
'''''
'''''        .Destination = crptToWindow
'''''        .WindowTitle = "Recibo"
'''''        .Action = 1
'''''    End With
'''''    Exit Sub
'''''
'''''error:
'''''    Maneja_Error Err
End Sub

Function MuestraTasa(TipoInteres As Integer, TipoPeriodo As Integer, TipoPlazo As Integer, crPrestamo As Double, ExisteCliente As Boolean, Etiqueta As Label, Autos As Boolean)
Dim rcTasas As New ADODB.Recordset
Dim TasaTipica As Double, TasaPromocion As Double, TasaPreferencial As Double, LimiteInferior As Double, LimiteSuperior As Double
    
    TasaTipica = 0
    TasaPromocion = 0
    TasaPreferencial = 0
    LimiteInferior = 0
    LimiteSuperior = 0
    
    LimiteInferior = Regresa_Valor_BD("LimiteInferior" & IIf(Autos, "Autos", ""))
    LimiteSuperior = Regresa_Valor_BD("LimiteSuperior" & IIf(Autos, "Autos", ""))
    
    rcTasas.Open "SELECT TasaTipica,TasaPromocion,TasaPreferencial FROM configuraciontasas WHERE IDTipoInteres=" & TipoInteres & " AND IDTipoPeriodo=" & TipoPeriodo & " AND IDPlazo=" & TipoPlazo, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcTasas.BOF And Not rcTasas.EOF Then
        
        TasaTipica = rcTasas!TasaTipica
        TasaPromocion = rcTasas!TasaPromocion
        TasaPreferencial = rcTasas!TasaPreferencial
    End If
    rcTasas.Close
    Set rcTasas = Nothing
    
    If ExisteCliente = False Then
        
        If crPrestamo >= LimiteSuperior Then
            
            MuestraTasa = TasaPreferencial
            
        ElseIf crPrestamo >= LimiteInferior Then
            
            MuestraTasa = TasaPromocion
        
        Else
            
            MuestraTasa = TasaTipica
        End If
    
    Else
        
        If crPrestamo >= LimiteInferior Then
            
            MuestraTasa = TasaPreferencial
        Else
            
            MuestraTasa = TasaPromocion
        End If
        
    End If
    
    Etiqueta.Caption = Format(MuestraTasa, "0.00") & "%"
End Function

Sub SacaTasa(crPrestamo As Double, TipoInteres As Integer, TipoPeriodo As Integer, TipoPlazo As Integer, ExisteCliente As Boolean)
Dim rcConsulta As New ADODB.Recordset
Dim Meses As Integer, Almacenaje As Double, Seguro As Double

On Error GoTo Error
    
    'Tasa
    rcConsulta.Open "SELECT tipointeres.Descripcion AS TipoInteres,tipoperiodo.Descripcion AS TipoPeriodo,tipoperiodo.Periodo,plazos.Descripcion AS Vencimiento " _
                    & "FROM configuraciontasas INNER JOIN plazos ON configuraciontasas.IDPlazo=plazos.ID INNER JOIN tipoperiodo ON configuraciontasas.IDTipoPeriodo=tipoperiodo.ID INNER JOIN tipointeres ON configuraciontasas.IDTipoInteres=tipointeres.ID WHERE " _
                    & "configuraciontasas.IDTipoInteres=" & TipoInteres & " AND configuraciontasas.IDTipoPeriodo=" & TipoPeriodo & " AND configuraciontasas.IDPlazo=" & TipoPlazo, dbDatos, adOpenForwardOnly, adLockReadOnly

    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        
        Select Case rcConsulta!TipoPeriodo
        Case "MENSUAL"
            
            Meses = 1
        Case "QUINCENAL"
            
            Meses = 2
        Case "SEMANAL"
            
            Meses = 4
            
        End Select
        
        Almacenaje = Regresa_Valor_BD("Almacenaje")
        Seguro = Regresa_Valor_BD("Seguro")
        
        lblAlmacenaje.Caption = Format((Almacenaje / 30) * rcConsulta!Periodo, "0.00") & "%"
        lblSeguro.Caption = Format((Seguro / 30) * rcConsulta!Periodo, "0.00") & "%"
        lblVencimiento.Caption = Format(IIf(cmbTipoInteres.text = "FIJA", DateAdd("M", 2, Date), IIf(rcConsulta!Periodo = 30, DateAdd("M", rcConsulta!Vencimiento, Date), DateAdd("D", (rcConsulta!Periodo * rcConsulta!Vencimiento) - 1, Date))), "DD/MMM/YYYY")
        MuestraTasa TipoInteres, TipoPeriodo, TipoPlazo, crPrestamo, ExisteCliente, lblTasa, False
    
    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub

Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Function CalculaCambio(crEfectivo As Double, crImporte As Double, Pestana As Integer) As Boolean
    
    lblCambio.Caption = "CAMBIO: " & Format(crEfectivo - crImporte, FMoneda)
    lblCambio.ForeColor = &HFF&
    lblCambio.Visible = True
    lblCambio.Tag = 1
    Abrir_Cajon
End Function
