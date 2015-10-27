VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmSelecAlertaLavado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccion de Alertas"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelecAlertaLavado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   9825
   Begin VB.Frame frameDescripcion 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   9735
      Begin VB.TextBox txtDescripcion 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1440
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   150
         Width           =   8100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8550
      TabIndex        =   0
      Top             =   4860
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmSelecAlertaLavado.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   4140
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmSelecAlertaLavado.frx":055E
   End
   Begin vbAcceleratorGrid6.vbalGrid grdAlertas 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5741
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      HighlightBackColor=   -2147483645
      HighlightForeColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
End
Attribute VB_Name = "frmSelecAlertaLavado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim strDescripcion As String
Dim vIdAlerta As Integer, vDescAlerta As String
Dim vAntIdAlerta As Integer, vAntDescAlerta As String
Dim vModulo As String

Private Sub Crear_Encabezados()

On Error Resume Next

    With grdAlertas
        .Clear True
        .AddColumn "K1", "Clave", ecgHdrTextALignLeft, , 50, False, , , , , , CCLSortString
        .AddColumn "K2", "Descripcion", ecgHdrTextALignLeft, , 1000, , , , , , , CCLSortString
        .AddColumn "K3", "Req. Desc", ecgHdrTextALignLeft, , 50, False, , , , , , CCLSortNumeric
    End With
End Sub

Sub Inicializar()
    Poner_Flat Fl, Me.Controls, Me
    
    Crear_Encabezados
    Cargar_Datos
    CentrarForm Me, frmMDI
End Sub

Public Sub Mostrar(ByRef tTipoAlerta As Integer, ByRef tDescAlerta As String, ByVal tModulo As Integer)
    strDescripcion = ""
    
    vAntIdAlerta = tTipoAlerta: vAntDescAlerta = tDescAlerta
    
    vModulo = tModulo
    Me.Show vbModal
    
    If vIdAlerta = 0 Then
        tTipoAlerta = vAntIdAlerta
        tDescAlerta = vAntDescAlerta
    Else
        tTipoAlerta = vIdAlerta
        tDescAlerta = vDescAlerta
    End If
    
End Sub

Private Sub grdAlertas_Click(ByVal lRow As Long, ByVal lCol As Long)
    
    If grdAlertas.Rows > 0 Then
        If grdAlertas.SelectedRow > 0 Then
            If grdAlertas.CellText(grdAlertas.SelectedRow, 3) = 1 Then
                'txtDescripcion.Enabled = True
                frameDescripcion.Visible = True
            Else
                'txtDescripcion.Enabled = False
                frameDescripcion.Visible = False
                txtDescripcion.text = ""
            End If
        End If
    End If
    
End Sub

Private Sub grdAlertas_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    If grdAlertas.Rows > 0 Then
        If grdAlertas.SelectedRow > 0 Then
            If grdAlertas.CellText(grdAlertas.SelectedRow, 3) = 1 Then
                'txtDescripcion.Enabled = True
                frameDescripcion.Visible = True
            Else
                'txtDescripcion.Enabled = False
                frameDescripcion.Visible = False
                txtDescripcion.text = ""
            End If
        End If
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
    Seleccionar_Texto txtDescripcion
    Cambiar_Color True, txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDescripcion_LostFocus()
    Cambiar_Color False, txtDescripcion
End Sub


Private Sub cmdAceptar_Click()
    If Validar = True Then
        Seleccionar
        'Unload Me
    End If
    Unload Me
End Sub

'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub Cargar_Datos()
Dim rcAlertas As New ADODB.Recordset
Dim vTabla As String

On Error GoTo Error
    
    Select Case vModulo
        Case MLD_PRESTAMO: vTabla = "mld_prestamos_tipo_alertas"
        Case MLD_METALES: vTabla = "mld_metales_tipo_alertas"
        Case MLD_VEHICULOS: vTabla = "mld_vehiculos_tipo_alertas"
        Case MLD_INMUEBLES: vTabla = "mld_inmuebles_tipo_alertas"
    End Select
    
    rcAlertas.Open "SELECT * FROM " & vTabla & " ORDER BY Clave ASC", dbDatos, adOpenForwardOnly, adLockReadOnly
   
    grdAlertas.Redraw = False
    With rcAlertas
        While Not .EOF
            
            grdAlertas.AddRow
            grdAlertas.CellText(grdAlertas.Rows, 1) = !Clave
            grdAlertas.CellItemData(grdAlertas.Rows, 1) = !ID
            grdAlertas.CellTextAlign(grdAlertas.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdAlertas.CellText(grdAlertas.Rows, 2) = Trim(!Descripcion)
            grdAlertas.CellTextAlign(grdAlertas.Rows, 2) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdAlertas.CellText(grdAlertas.Rows, 3) = !ReqDesc
            grdAlertas.CellTextAlign(grdAlertas.Rows, 3) = DT_CENTER Or DT_WORD_ELLIPSIS
        
        .MoveNext
        Wend
    End With
    rcAlertas.Close
    Set rcAlertas = Nothing
    grdAlertas.Redraw = True
Exit Sub
  
Error:
    Maneja_Error Err
    Set rcAlertas = Nothing
End Sub

Public Function Seleccionar() As Integer
    'obj.text = grdUsuarios.CellText(grdUsuarios.SelectedRow, 1)
    'obj.Tag = grdUsuarios.CellItemData(grdUsuarios.SelectedRow, 1)
    If grdAlertas.SelectedRow > 0 Then
        vIdAlerta = grdAlertas.CellItemData(grdAlertas.SelectedRow, 1)
        vDescAlerta = Trim(txtDescripcion)
    End If
    'If Mostrar Then frm.Buscar grdUsuarios.CellItemData(grdUsuarios.SelectedRow, 1)
    Unload Me
End Function


Public Function Validar() As Boolean
    
    Validar = True
    
'    If grdAlertas.SelectedRow = 0 Then
'        Validar = False
'        MsgBox "Seleccione la Alerta !!!", vbInformation, Me.Caption
'        Exit Function
'    End If
    
    If grdAlertas.SelectedRow > 0 Then
        If grdAlertas.CellText(grdAlertas.SelectedRow, 3) = 1 Then
            If Trim(txtDescripcion.text) = "" Then
                Validar = False
                MsgBox "Especifique la Descripción de la Alerta !!!", vbInformation, Me.Caption
                If frameDescripcion.Visible = True Then txtDescripcion.SetFocus
                Exit Function
            End If
        End If
    End If
End Function
