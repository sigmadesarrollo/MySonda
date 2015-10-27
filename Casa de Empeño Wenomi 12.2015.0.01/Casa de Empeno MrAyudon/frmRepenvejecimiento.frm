VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmRepenvejecimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de envejecimiento"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepenvejecimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   900
   ScaleWidth      =   5205
   Begin VB.ComboBox cmbTipo 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   240
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
      Picture         =   "frmRepenvejecimiento.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   2
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "       &Imprimir"
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
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmRepenvejecimiento.frx":009D
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmRepenvejecimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tipo As String
Dim Fl() As cFlatControl

Private Sub cmbTipo_GotFocus()
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmbTipo_LostFocus()
    Cambiar_Color False, cmbTipo
End Sub

Private Sub cmdAceptar_Click()
    With frmMDI.Cr
        .Reset
        .ReportFileName = Path & "\Reportes\Envejecimiento.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{DetallesEntradaInventario.cantidad} > 0 AND {DetallesEntradaInventario.TipoEntrada} <> " & D_FUNDICION & " AND {DetallesEntradaInventario.TipoEntrada} <> " & D_OTRO & " " & IIf(Trim(cmbTipo.text) = "(TODOS)", "", " and {Detallesentradainventario.Tipo}=" & Trim(cmbTipo.ItemData(cmbTipo.ListIndex)) & "")
        .DiscardSavedData = True
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowTitle = "Reporte de envejecimiento"
        .Action = 1
    End With
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Sub Inicializar()
    Cargar_Combos "Descripcion", "tipo", cmbTipo
    cmbTipo.ListIndex = 0
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

'cargamos el tipo de prenda
Private Sub Cargar_Combos(Campo As String, Tabla As String, Combo As ComboBox)
On Error GoTo error

   Dim rcTipos As New ADODB.Recordset
   
   rcTipos.Open "SELECT * FROM " & Tabla, dbDatos, adOpenDynamic, adLockOptimistic
   
   Combo.Clear
   Combo.AddItem "(TODOS)"
   With rcTipos
      While Not .EOF
         Combo.AddItem .Fields(Campo)
         Combo.ItemData(Combo.NewIndex) = !ID
         .MoveNext
      Wend
   End With
   
   rcTipos.Close
error:
   Maneja_Error Err
   
   Set rcTipos = Nothing
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Public Function regresa_tipoprenda() As String
    Cargar_Combos "Tipo", "Tipo", cmbTipo
    regresa_tipoprenda = cmbTipo.text
End Function
