VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmInventariooo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dotación a inventario"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInventariooo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11025
   Begin VB.Frame frmEntradas 
      Caption         =   "ENTRADAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      Left            =   30
      TabIndex        =   2
      Top             =   675
      Width           =   10950
      Begin vbAcceleratorGrid6.vbalGrid grdArticulos 
         Height          =   4905
         Left            =   45
         TabIndex        =   3
         Top             =   225
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   8652
         GridLines       =   -1  'True
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
         ScrollBarStyle  =   2
         Editable        =   -1  'True
         DisableIcons    =   -1  'True
         Begin VB.TextBox txtSerie 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   10440
            TabIndex        =   13
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cmbTipo 
            Height          =   315
            Left            =   9120
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtGrupo 
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   840
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox txtPesoo 
            BorderStyle     =   0  'None
            Height          =   405
            Left            =   8040
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox dcbKilates 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmInventariooo.frx":000C
            Left            =   5955
            List            =   "frmInventariooo.frx":002E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   0
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtCode 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   120
            MaxLength       =   8
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtArticulo 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox txtPrecioo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   7920
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtCosto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   6885
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   5160
            TabIndex        =   4
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Label lblCodigo1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   4560
         Width           =   75
      End
      Begin VB.Label lblDescripcion1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   4560
         Width           =   75
      End
      Begin VB.Label lblKilates1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6360
         TabIndex        =   20
         Top             =   4560
         Width           =   75
      End
      Begin VB.Label lblPrecio1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7560
         TabIndex        =   19
         Top             =   4560
         Width           =   75
      End
      Begin VB.Label lblNumCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   975
         TabIndex        =   18
         Top             =   5490
         Width           =   75
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Num:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   255
         TabIndex        =   17
         Top             =   5490
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "<Total>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6615
         TabIndex        =   16
         Top             =   5490
         Width           =   1035
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5535
         TabIndex        =   15
         Top             =   5490
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         TabIndex        =   14
         Top             =   5340
         Width           =   11415
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   6645
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
      Picture         =   "frmInventariooo.frx":0074
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   6645
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   3
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
      MaskColor       =   16777215
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmInventariooo.frx":0105
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   26
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   795
   End
   Begin VB.Label lblFolio 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8760
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Folio:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "frmInventariooo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 27/05/2002
' Modulo frmInventario - frmInventario.frm
' Ultima Modificacion - 27/05/2002
' Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

Option Explicit
Dim Fl() As cFlatControl
Dim Band As Boolean

Private Function Validar() As Boolean
Dim i As Integer

Validar = True

For i = 1 To grdArticulos.Rows
    If grdArticulos.CellText(i, 1) = "" Or grdArticulos.CellText(i, 2) = "" Or grdArticulos.CellText(i, 3) = "" Or grdArticulos.CellText(i, 4) = "" Or grdArticulos.CellText(i, 5) = "" Or grdArticulos.CellText(i, 6) = "" Or grdArticulos.CellText(i, 7) = "" Or grdArticulos.CellText(i, 8) = "" Or grdArticulos.CellText(i, 9) = "" Then
        If grdArticulos.CellText(i, 1) = "" And grdArticulos.CellText(i, 2) = "" And grdArticulos.CellText(i, 3) = "" And grdArticulos.CellText(i, 4) = "" And grdArticulos.CellText(i, 5) = "" And grdArticulos.CellText(i, 6) = "" And grdArticulos.CellText(i, 7) = "" And grdArticulos.CellText(i, 8) = "" And grdArticulos.CellText(i, 9) = "" Then GoTo 125
        
        If grdArticulos.CellText(i, 8) = "" Then MsgBox "Seleccione el tipo del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 8, 13, False: Exit Function
        If grdArticulos.CellText(i, 1) = "" Then MsgBox "Seleccione el grupo del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 1, 13, False: Exit Function
        If grdArticulos.CellText(i, 2) = "" Then MsgBox "Introduzca la descripción del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 2, 13, False: Exit Function
        If grdArticulos.CellText(i, 3) = "" Then MsgBox "Introduzca la cantidad de artículos !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 3, 13, False: Exit Function
        If grdArticulos.CellText(i, 4) = "" And grdArticulos.CellText(i, 9) = "METAL" Then MsgBox "Seleccione el kilataje del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 4, 13, False: Exit Function
        If grdArticulos.CellText(i, 5) = "" And grdArticulos.CellText(i, 9) = "METAL" Then MsgBox "Introduzca el peso del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 5, 13, False: Exit Function
        If grdArticulos.CellText(i, 6) = "" Then MsgBox "Introduzca el costo del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 6, 13, False: Exit Function
        If grdArticulos.CellText(i, 7) = "" Then MsgBox "Introduzca el precio del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 7, 13, False: Exit Function
    End If
125:
Next i
End Function

Private Sub cmbTipo_Click()
grdArticulos.CellText(grdArticulos.SelectedRow, 8) = cmbTipo.Text
grdArticulos.CellItemData(grdArticulos.SelectedRow, 8) = cmbTipo.ItemData(cmbTipo.ListIndex)
cmbTipo.Visible = False
grdArticulos_CancelEdit
End Sub

Private Sub cmbTipo_GotFocus()
cmbTipo.BackColor = &HC0FFFF
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 9) = cmbTipo.Text
    cmbTipo.Visible = False
    grdArticulos.CancelEdit
    grdArticulos.SetFocus
End If
End Sub

Private Sub cmbTipo_LostFocus()
cmbTipo.BackColor = vbWhite
End Sub

Private Sub cmdImprimir_Click()
Screen.MousePointer = vbHourglass
If Validar_datos Then If Validar Then Grabar_Entradas
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub dcbKilates_Click()
Dim Cantidad As Integer, peso As Double, Total As Double, kilataje As String, kilates As String


If dcbKilates.ListIndex = -1 Then Exit Sub
If dcbKilates.Text <> "" Then kilates = dcbKilates.Text Else kilates = grdArticulos.CellText(grdArticulos.SelectedRow, 4)
kilataje = dcbKilates.ItemData(dcbKilates.ListIndex)

grdArticulos.CellText(grdArticulos.SelectedRow, 4) = dcbKilates.Text 'sacakilates(grdArticulos.CellItemData(grdArticulos.SelectedRow, 5))
grdArticulos.CellItemData(grdArticulos.SelectedRow, 4) = kilataje

If grdArticulos.CellText(grdArticulos.SelectedRow, 3) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 3) Else Cantidad = 0
If grdArticulos.CellText(grdArticulos.SelectedRow, 5) <> "" Then peso = grdArticulos.CellText(grdArticulos.SelectedRow, 5) Else peso = 0

If grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "" Then
    Set rcConsulta = dbDatos.Execute("select " & "Venta" & dcbKilates.Text & " as costo from parametros")
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        Total = (peso * rcConsulta!costo) ''* cantidad
        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
    End If
Else
    Total = 0
    grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
End If

grdArticulos.CancelEdit
dcbKilates.Visible = False
lblTotal.Caption = Format(regresa_total, "##,###0.00")
Set rcConsulta = Nothing
End Sub

Private Sub dcbKilates_GotFocus()
dcbKilates.BackColor = &HC0FFFF
End Sub

Private Sub dcbKilates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then dcbKilates.Visible = False
End Sub

Private Sub dcbKilates_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 5) = dcbKilates.Text
    dcbKilates.Visible = False
    grdArticulos.CancelEdit
    grdArticulos.SetFocus
End If
End Sub

Private Sub dcbKilates_LostFocus()
dcbKilates.BackColor = vbWhite
End Sub

Private Sub Form_Load()
inicializar
End Sub

'Inicializamos la forma
Private Sub inicializar()
Screen.MousePointer = vbHourglass
frmEntradas.BorderStyle = 0
lblFecha.Caption = Format(Date, "DD/MMM/YYYY")
Cargar_Combos "Descripcion", "Tipo", cmbTipo
Crear_Encabezados
Poner_Flat Fl, Me.Controls, Me
lblTotal.Caption = "0.00"
lblFolio.Caption = Regresa_Movimiento(False, "FolioDotacion")
Screen.MousePointer = vbDefault
CentrarForm Me, frmMDI
End Sub

'Creamos los encabezados
Private Sub Crear_Encabezados()
With grdArticulos
'   .AddColumn "K1", "Código", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
   .AddColumn "K8", "Grupo", ecgHdrTextALignLeft, , 45, , , , , , , CCLSortString
   .AddColumn "K2", "Descripción", ecgHdrTextALignLeft, , 180, , , , , , , CCLSortString
   .AddColumn "K5", "Cantidad", ecgHdrTextALignRight, , 55, , , , , , , CCLSortString
   .AddColumn "K3", "Kilates", ecgHdrTextALignLeft, , 62, , , , , , , CCLSortNumeric
   .AddColumn "K7", "Peso", ecgHdrTextALignRight, , 50, , , , , "0.000", , CCLSortNumeric
   .AddColumn "K6", "Costo", ecgHdrTextALignRight, , 65, , , , , "###,###,###,##0.00", , CCLSortString
   .AddColumn "K4", "Precio", ecgHdrTextALignRight, , 75, , , , , "###,###,###,##0.00", , CCLSortNumeric
   .AddColumn "K9", "Tipo", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
   .AddColumn "K10", "Serie", ecgHdrTextALignLeft, , 76, , , , , , , CCLSortString
   .GridLines = True
   .Rows = 20
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Quitar_Flat Fl()
Unload Me
End Sub

Private Sub grdArticulos_CancelEdit()
txtCantidad.Visible = False
txtCosto.Visible = False
txtPrecioo.Visible = False
txtCode.Visible = False
txtArticulo.Visible = False
dcbKilates.Visible = False
dcbKilates.ListIndex = -1
txtPesoo.Visible = False
txtGrupo.Visible = False
cmbTipo.Visible = False
txtSerie.Visible = False
End Sub

Private Sub grdArticulos_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
Dim Cantidad As Integer, precio As Double, Total As Double, i As Integer

If grdArticulos.SelectedRow > 0 Then
      If KeyCode = vbKeyDelete Then
       If MsgBox("Desea Eliminar Estos Articulos ??", vbQuestion + vbYesNo + vbDefaultButton2, "Dotación a Inventario") = vbYes Then
            Cantidad = 0
            precio = 0
            Total = 0
            grdArticulos.RemoveRow grdArticulos.SelectedRow
            For i = 1 To grdArticulos.Rows
               Cantidad = grdArticulos.CellText(i, 3)
               precio = grdArticulos.CellText(i, 6)
               Total = Total + (Cantidad * precio)
            Next i
            'lblTotal.Caption = "$ " & Format(Total, "##,###0.00")
            'lblNumCap.Caption = grdArticulos.Rows
            grdArticulos.CancelEdit
            'txtClave.SetFocus
        End If
      End If
   End If
End Sub

Private Sub grdArticulos_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
 Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
   Dim sText As String
   Dim obj As Object
   
   txtCantidad.Visible = False
   grdArticulos_CancelEdit
   
    If (lCol = 4 Or lCol = 5) And (grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "METAL" And grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "ORO") Then Exit Sub
    If lCol = 10 And grdArticulos.CellText(grdArticulos.SelectedRow, 8) = "METAL" Then Exit Sub
    
    Select Case lCol
    Case 1: Set obj = txtGrupo
    Case 2: Set obj = txtArticulo
    Case 3: Set obj = txtCantidad
    Case 4: Set obj = dcbKilates
    Case 5: Set obj = txtPesoo
    Case 6: Set obj = txtCosto
    Case 7: Set obj = txtPrecioo
    Case 8: Set obj = cmbTipo
    Case 9: Set obj = txtSerie
    Case Else: Exit Sub
    End Select
   
   grdArticulos.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight

   If Not IsMissing(grdArticulos.CellText(lRow, lCol)) Then
      sText = grdArticulos.CellText(lRow, lCol)
   Else
      sText = ""
   End If
   
   If lCol <> 4 And lCol <> 8 Then
         If lCol = 1 Or lCol = 2 Then
            obj.Alignment = vbLeftJustify
            grdArticulos.CellTextAlign(grdArticulos.SelectedRow, lCol) = DT_LEFT
         Else
            obj.Alignment = vbRightJustify
         End If
   
      'iKeyAscii = Solo_Numeros(iKeyAscii)
      If (iKeyAscii > 13) Then
         sText = Chr$(iKeyAscii) & sText
         obj.Text = sText
         obj.SelStart = 1
         obj.SelLength = Len(sText)
      Else
         obj.Text = sText
         obj.SelStart = 0
         obj.SelLength = Len(sText)
      End If
      
      Set txtCantidad.Font = grdArticulos.CellFont(lRow, lCol)
      If grdArticulos.CellBackColor(lRow, lCol) = -1 Then
         txtCantidad.BackColor = grdArticulos.BackColor
      Else
         txtCantidad.BackColor = grdArticulos.CellBackColor(lRow, lCol)
      End If
   End If
   
   If lCol <> 4 And lCol <> 8 Then
      obj.Move lLeft + 40, lTop + 25, lWidth - 60, lHeight - 50
   Else
      obj.Move lLeft + 40, lTop + 25, lWidth - 60
   End If
   
   obj.Visible = True
   obj.ZOrder
   
   If lCol = 1 Then
    txtGrupo.Visible = True
    frmMostrarGrupo.Ver Me, txtGrupo, 1, 2
    Exit Sub
   End If
   
   obj.SetFocus
End Sub

Private Sub tTab_BeforeClick(ByVal lTab As Long, bCancel As Boolean)

End Sub

Private Sub txtArticulo_GotFocus()
txtArticulo.BackColor = &HC0FFFF
End Sub

Private Sub txtArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then txtArticulo.Visible = False
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
mayusculas KeyAscii
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 2) = txtArticulo.Text
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 2) = DT_LEFT
    grdArticulos.CancelEdit
    txtArticulo.Visible = False
    grdArticulos.SetFocus
End If
End Sub

Private Sub txtArticulo_LostFocus()
txtArticulo.BackColor = vbWhite
End Sub

Private Sub txtCantidad_GotFocus()
txtCantidad.BackColor = &HC0FFFF
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then txtCantidad.Visible = False
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
Dim i As Integer, Total As Double, Cantidad As Integer, precio As Double, x As Integer, peso As Double, kilataje As String


KeyAscii = Solo_Numeros(KeyAscii)

If KeyAscii = vbKeyReturn Then
   
   kilataje = RegresaKilates(IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 8) = "METAL", grdArticulos.CellText(grdArticulos.SelectedRow, 4), ""), grdArticulos.CellText(grdArticulos.SelectedRow, 8))

   Total = 0
   Cantidad = 0
   precio = 0
   
   If txtCantidad.Text <> "" Then grdArticulos.CellText(grdArticulos.SelectedRow, 3) = txtCantidad.Text Else grdArticulos.CellText(grdArticulos.SelectedRow, 3) = "": GoTo 125
   grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 3) = DT_RIGHT
   For i = 1 To grdArticulos.Rows
        Cantidad = IIf(grdArticulos.CellText(i, 3) = "", 0, Val(grdArticulos.CellText(i, 3)))
        If grdArticulos.CellText(i, 6) <> "" Then precio = grdArticulos.CellText(i, 7) Else precio = 0
        Total = Total + (Cantidad * precio)
   Next i
  
    '--------
    If grdArticulos.CellText(grdArticulos.SelectedRow, 3) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 3) Else Cantidad = 0
    If grdArticulos.CellText(grdArticulos.SelectedRow, 5) <> "" Then peso = grdArticulos.CellText(grdArticulos.SelectedRow, 5) Else peso = 0
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Fino" Then
        Set rcConsulta = dbDatos.Execute("select " & "Venta" & IIf(dcbKilates.Text = "", grdArticulos.CellText(grdArticulos.SelectedRow, 4), dcbKilates.Text) & " as costo from parametros")
        If Not rcConsulta.BOF And Not rcConsulta.EOF Then
            Total = (peso * rcConsulta!costo) * Cantidad
        End If
        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
    Else
        Total = 0
        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
    End If
   '--------

125:
   grdArticulos.CancelEdit
   txtCantidad.Visible = False
   lblTotal.Caption = Format(regresa_total, "##,###0.00")
   KeyAscii = 0
   grdArticulos.SetFocus
ElseIf KeyAscii = vbKeyEscape Then
   grdArticulos.CancelEdit
   KeyAscii = 0
End If

Set rcConsulta = Nothing
End Sub

Private Sub txtCantidad_LostFocus()
txtCantidad.BackColor = vbWhite
End Sub

'Buscamos el grupo y lo visualizamos
Private Function Buscar_Clave(Codigo As String) As Boolean
   On Error GoTo error
   Dim rcGrupos As New ADODB.Recordset
   
   Buscar_Clave = True
   
   rcGrupos.Open "SELECT * FROM Grupos WHERE Clave='" & Codigo & "'", dbDatos, adOpenKeyset, adLockOptimistic
   
   If rcGrupos.RecordCount = 0 Then
      Buscar_Clave = False
   End If

   rcGrupos.Close
   
error:
   Maneja_Error Err
   
   Set rcGrupos = Nothing
      
End Function

Private Sub txtCode_GotFocus()
txtCode.BackColor = &HC0FFFF
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then txtCode.Visible = False
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 1) = txtCode.Text
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 1) = DT_LEFT
    grdArticulos.CancelEdit
    txtCode.Visible = False
End If
End Sub

Private Sub txtCode_LostFocus()
txtCode.BackColor = vbWhite
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Numeros(KeyAscii)
   If KeyAscii = vbKeyReturn Then
      'If Len(txtCodigo.Text) = 8 Then 'txtCodigo.Text = "0" & txtCodigo.Text
         'Buscar_Articulo txtCodigo.Text, grdSalidas, lblTotalSalida, txtCodigo
      'End If
   End If
   'Pasar_Foco KeyAscii
End Sub

Private Sub txtCosto_GotFocus()
txtCosto.BackColor = &HC0FFFF
End Sub

Private Sub txtCosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then txtCosto.Visible = False
End Sub

Private Sub txtCosto_KeyPress(KeyAscii As Integer)
Dim kilataje As String, Cantidad As Integer, peso As Double, Total As Double
'

KeyAscii = Solo_Numeros(KeyAscii, 1)

If KeyAscii = vbKeyReturn Then
   
   kilataje = RegresaKilates(grdArticulos.CellText(grdArticulos.SelectedRow, 4))
   
   grdArticulos.CellText(grdArticulos.SelectedRow, 6) = txtCosto.Text
   grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
   grdArticulos.CancelEdit
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 3) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 3) Else Cantidad = 0
    If grdArticulos.CellText(grdArticulos.SelectedRow, 5) <> "" Then peso = grdArticulos.CellText(grdArticulos.SelectedRow, 5) Else peso = 0
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Fino" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Plata" Then
        Set rcConsulta = dbDatos.Execute("select " & "Venta" & IIf(dcbKilates.Text = "", grdArticulos.CellText(grdArticulos.SelectedRow, 4), dcbKilates.Text) & " as costo from parametros")
        If Not rcConsulta.BOF And Not rcConsulta.EOF Then
            Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 6) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 6))
            grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
        End If
    Else
        
        Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 6) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 6))
        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
    End If
   
   txtCosto.Visible = False
   lblTotal.Caption = Format(regresa_total, "##,###0.00")

   KeyAscii = 0
   grdArticulos.SetFocus
ElseIf KeyAscii = vbKeyEscape Then
   grdArticulos.CancelEdit
   grdArticulos.SetFocus
   KeyAscii = 0
End If
Set rcConsulta = Nothing
End Sub

Private Sub txtCosto_LostFocus()
txtCosto.BackColor = vbWhite
End Sub

Private Sub txtGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then txtGrupo.Visible = False
End Sub

Public Sub txtGrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 1) = txtGrupo.Text
    'grdArticulos.CellText(grdArticulos.SelectedRow, 1) = genera_codigo(grdArticulos.CellText(grdArticulos.SelectedRow, 2), Format(IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 8) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 8)), "00000"))
    txtGrupo.Visible = False
End If
End Sub

Private Sub txtNombre2_KeyPress(KeyAscii As Integer)
Pasar_Foco KeyAscii
End Sub

Private Sub txtPesoo_GotFocus()
txtPesoo.BackColor = &HC0FFFF
End Sub

Private Sub txtPesoo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then txtPesoo.Visible = False
End Sub

Private Sub txtPesoo_KeyPress(KeyAscii As Integer)
Dim peso As Double, Cantidad As Integer, costo As Double, i As Integer
Dim Total As Double

KeyAscii = Solo_Numeros(KeyAscii, 1)
If KeyAscii = vbKeyReturn Then
    
   If txtPesoo.Text <> "" Then grdArticulos.CellText(grdArticulos.SelectedRow, 5) = txtPesoo.Text: peso = txtPesoo.Text Else grdArticulos.CellText(grdArticulos.SelectedRow, 5) = "": peso = 0
   grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 5) = DT_RIGHT

    '--------
    If grdArticulos.CellText(grdArticulos.SelectedRow, 3) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 3) Else Cantidad = 0
    If grdArticulos.CellText(grdArticulos.SelectedRow, 5) <> "" Then peso = grdArticulos.CellText(grdArticulos.SelectedRow, 5) Else peso = 0
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Fino" Then
        Set rcConsulta = dbDatos.Execute("select " & "Venta" & IIf(dcbKilates.Text = "", grdArticulos.CellText(grdArticulos.SelectedRow, 4), dcbKilates.Text) & " as costo from parametros")
        If Not rcConsulta.BOF And Not rcConsulta.EOF Then
            Total = (peso * rcConsulta!costo) * Cantidad
        End If
        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
    Else
        Total = 0
        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
    End If
   '--------
   
   grdArticulos.CancelEdit
   txtPesoo.Visible = False
   KeyAscii = 0
   grdArticulos.SetFocus
ElseIf KeyAscii = vbKeyEscape Then
   grdArticulos.CancelEdit
   KeyAscii = 0
End If
End Sub

Private Sub txtPesoo_LostFocus()
txtPesoo.BackColor = vbWhite
End Sub

'Buscamos la categoria por medio de la clave
Public Sub Buscar_Grupo(Codigo As String)
   On Error GoTo error
   Dim rcGrupo As New ADODB.Recordset
   
   rcGrupo.Open "SELECT * FROM Grupos WHERE Clave='" & Codigo & "'", dbDatos, adOpenKeyset, adLockOptimistic
      
   With rcGrupo
      If .RecordCount = 0 Then
         MsgBox "El Grupo no se encuentra dado de alta", vbOKOnly + vbCritical
      End If
   End With
   
   rcGrupo.Close

error:
   Maneja_Error Err
   
   Set rcGrupo = Nothing

End Sub

'Grabamos todos los datos necesarios
Private Sub Grabar_Entradas()

If grdArticulos.Rows > 0 Then
   Screen.MousePointer = vbHourglass
   Grabar_Encabezado
   Sleep 1000
   Limpiar
   grdArticulos.GridLines = True
   grdArticulos.Rows = 20
   lblFolio.Caption = Regresa_Movimiento(False, "FolioDotacion")
   Screen.MousePointer = vbDefault
Else
    MsgBox "Introduzca los datos de los artículos que desea dotar !!", vbInformation, "Dotación a inventario"
    txtCode.SetFocus
End If
End Sub

'imprimimos la entrada del inventario
Private Sub Imprimir_Entrada()
Dim Folio As Integer

With frmMDI.Cr
     Folio = lblFolio.Caption
    .Reset
    .DataFiles(0) = Path & "\Base De Datos\Datos.mdb"
    .DataFiles(1) = Path & "\Base De Datos\Datos.mdb"
    .password = Chr(10) & "administrativo"
    .ReportFileName = Path & "\Reportes\EntradaInventarioo.rpt"
    .DiscardSavedData = True
    .SelectionFormula = "{DetallesEntradaInventario.folio}=" & Folio & ""
    .Formulas(0) = "Folio='" & Trim(lblFolio.Caption) & "'"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
End With
End Sub

Private Sub Limpiar()
grdArticulos.Clear
grdArticulos.GridLines = True
grdArticulos.Rows = 20
lblTotal.Caption = "0.00"
End Sub

'Grabamos el encabezado de la entrada
Private Function Grabar_Encabezado() As Long
   On Error GoTo error
   Dim rcID As New ADODB.Recordset
   Dim Folio As Long
   
   Folio = Regresa_Movimiento(False, "FolioDotacion")
   Regresa_Movimiento True, "FolioDotacion"
   
   dbDatos.Execute "INSERT INTO EntradaInventario (Folio,Fecha,IDUsuario) VALUES (" & Folio & ",'" & Format(lblFecha.Caption, "YYYY/MM/DD") & "'," & frmMDI.IDUsuario & ")"
   rcID.Open "SELECT MAX(ID) AS IDD FROM EntradaInventario", dbDatos, adOpenDynamic, adLockOptimistic
   
   Grabar_Inventario rcID!IDD, Folio
   
   rcID.Close
   
error:
      Maneja_Error Err
      
      Set rcID = Nothing
   
End Function

'Grabamos el inventario
Private Sub Grabar_Inventario(ID As Long, Folio As Long)
On Error GoTo error

Dim Indice As Integer, Movimiento As Long, crImporte As Double, kilataje As Integer
Dim Codigo As String

    crImporte = lblTotal.Caption
   
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
   
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
            "('" & Format(Date, "YYYY/MM/DD") & "'," & Movimiento & "," & Folio & ",'EN01','620301'," & crImporte & "," & TIPO_CARGO & ",1,'Entrada Inventario','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    
    'Grabamos el abono
    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
            "('" & Format(Date, "YYYY/MM/DD") & "'," & Movimiento & "," & Folio & ",'EN50','200950'," & crImporte & "," & TIPO_ABONO & ",1,'Entrada Inventario','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
           
    For Indice = 1 To grdArticulos.Rows
    
        If grdArticulos.CellText(Indice, 1) = "" And grdArticulos.CellText(Indice, 2) = "" And grdArticulos.CellText(Indice, 3) = "" And grdArticulos.CellText(Indice, 4) = "" And grdArticulos.CellText(Indice, 5) = "" And grdArticulos.CellText(Indice, 6) = "" And grdArticulos.CellText(Indice, 7) = "" Then GoTo 126
        
        kilataje = 0
        kilataje = RegresaKilates(IIf(grdArticulos.CellText(Indice, 8) = "METAL", grdArticulos.CellText(Indice, 4), ""), grdArticulos.CellText(Indice, 8))
        Codigo = CreaCodigoBarras(frmMDI.IDSucursal, ENTRADADOTACION, Trim(Folio), Indice, grdArticulos.CellItemData(Indice, 8))
        
        dbDatos.Execute "INSERT INTO DetallesEntradaInventario (IDEntrada,Codigo,Descripcion,Kilates,Peso,Costo,Precio,Cantidad,Tipo,Serie,SucursalOrigen,TipoEntrada) VALUES (" & _
                         ID & ",'" & Trim(Codigo) & "','" & grdArticulos.CellText(Indice, 2) & "'," & kilataje & "," & _
                         CDbl(grdArticulos.CellText(Indice, 5)) & "," & CDbl(grdArticulos.CellText(Indice, 6)) & "," & CDbl(grdArticulos.CellText(Indice, 7)) & "," & grdArticulos.CellText(Indice, 3) & "," & _
                         grdArticulos.CellItemData(Indice, 8) & ",'" & grdArticulos.CellText(Indice, 9) & "'," & frmMDI.IDSucursal & ", " & ENTRADADOTACION & ")"
         
126:
      Next Indice
   
error:
   Maneja_Error Err
   
End Sub

Private Sub txtPrecioo_GotFocus()
txtPrecioo.BackColor = &HC0FFFF
End Sub

Private Sub txtPrecioo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then txtPrecioo.Visible = False
End Sub

Private Sub txtPrecioo_KeyPress(KeyAscii As Integer)
Dim i As Integer, Cantidad As Integer, precio As Double, Total As Double

KeyAscii = Solo_Numeros(KeyAscii, 1)
If KeyAscii = vbKeyReturn Then
   
    grdArticulos.CellText(grdArticulos.SelectedRow, 7) = txtPrecioo.Text
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 7) = DT_RIGHT
    grdArticulos.CancelEdit
        
    If grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Fino" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Plata" Then
        Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 7) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 7))
    Else
        Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 7) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 7))
    End If

    txtCosto.Visible = False
    lblTotal.Caption = Format(regresa_total, "##,###0.00")
    KeyAscii = 0
    grdArticulos.SetFocus

ElseIf KeyAscii = vbKeyEscape Then
    grdArticulos.CancelEdit
    grdArticulos.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtPrecioo_LostFocus()
txtPrecioo.BackColor = vbWhite
End Sub

'Function REMATE(str As String) As Boolean
'Dim busca As ADODB.Recordset
'Set busca = New ADODB.Recordset
'
'Set busca = dbDatos.Execute("select clave from grupos where clave='" & Trim(str) & "'")
'If Not busca.BOF And Not busca.EOF Then REMATE = False Else REMATE = True
'End Function

'Function remates(folio As String, codigo As String)
'Dim busca As ADODB.Recordset
'Dim i As Integer, Cantidad As Integer, precio As Double, total As Double
'
'Set busca = New ADODB.Recordset
'
'Set busca = dbDatos.Execute("select id,destino from Empeno where codigoremate='" & Trim(folio) & "'")
'If Not busca.BOF And Not busca.EOF Then
'    If busca!destino = 5 Then
'        Set busca = dbDatos.Execute("select * from detallesempeño where idempeño=" & busca!ID & " order by articulo")
'        If Not busca.BOF And Not busca.EOF Then
'            busca.MoveFirst
'            While Not busca.EOF
'                grdArticulos.AddRow
'                grdArticulos.CellText(grdArticulos.Rows, 1) = codigo
'                grdArticulos.CellText(grdArticulos.Rows, 2) = busca!Articulo
'                grdArticulos.CellText(grdArticulos.Rows, 3) = busca!Cantidad
'                grdArticulos.CellTextAlign(grdArticulos.Rows, 3) = DT_RIGHT
'                grdArticulos.CellText(grdArticulos.Rows, 4) = busca!kilates
'                grdArticulos.CellTextAlign(grdArticulos.Rows, 4) = DT_RIGHT
'                grdArticulos.CellText(grdArticulos.Rows, 5) = busca!prestamo
'                grdArticulos.CellTextAlign(grdArticulos.Rows, 5) = DT_RIGHT
'                grdArticulos.CellText(grdArticulos.Rows, 6) = busca!avaluo
'                grdArticulos.CellTextAlign(grdArticulos.Rows, 6) = DT_RIGHT
'                'Poner_Colores2 grdArticulos, grdArticulos.Rows, &HE0E0E0
'            busca.MoveNext
'            Wend
'            grdArticulos.CancelEdit
'
'            For i = 1 To grdArticulos.Rows
'                Cantidad = grdArticulos.CellText(i, 3)
'                precio = grdArticulos.CellText(i, 6)
'                total = total + (Cantidad * precio)
'            Next i
'            'lblTotal.Caption = "$ " & Format(Total, "##,###0.00")
'            'lblNumCap.Caption = grdArticulos.Rows
'        End If
'    Else
'        MsgBox "Esta Boleta todavia no ha sido Marcada como Remate !!", vbInformation, "Dotación a Inventario"
'        'txtClave.SetFocus
'    End If
'End If
'End Function

Function sacagrupo(codegrupo As String) As Integer
Dim str As String


'Set rcConsulta = New ADODB.Recordset

str = Left(codegrupo, 2)

rcConsulta.Open "select id from grupos where clave='" & Trim(str) & "'", dbDatos, adOpenDynamic, adLockOptimistic
If Not rcConsulta.BOF And Not rcConsulta.EOF Then sacagrupo = rcConsulta!ID

rcConsulta.Close
Set rcConsulta = Nothing
End Function

Function mayusculas(ascii As Integer) As Integer
If (ascii >= 97) And (ascii <= 122) Then ascii = ascii - 32 Else If ascii = 39 Then ascii = 0
mayusculas = ascii
End Function

Private Function Validar_datos() As Boolean
Dim i As Integer, x As Integer

x = 0
For i = 1 To grdArticulos.Rows
    If grdArticulos.CellText(i, 1) = "" Or grdArticulos.CellText(i, 2) = "" Or grdArticulos.CellText(i, 3) = "" Or grdArticulos.CellText(i, 4) = "" Or Trim(grdArticulos.CellText(i, 5)) = "" Or Trim(grdArticulos.CellText(i, 6)) = "" Or grdArticulos.CellText(i, 7) = "" Or grdArticulos.CellText(i, 8) = "" Or grdArticulos.CellText(i, 9) = "" Then
        If grdArticulos.CellText(i, 1) = "" And grdArticulos.CellText(i, 2) = "" And grdArticulos.CellText(i, 3) = "" And grdArticulos.CellText(i, 4) = "" And grdArticulos.CellText(i, 5) = "" And grdArticulos.CellText(i, 6) = "" And grdArticulos.CellText(i, 7) = "" And grdArticulos.CellText(i, 8) = "" And grdArticulos.CellText(i, 9) = "" Then x = x + 1
    End If
Next i

If x = grdArticulos.Rows Then MsgBox "Introduzca los datos de los artículos que desea comprar !!", vbInformation, "Dotación a Inventario": Validar_datos = False: grdArticulos.SetFocus Else Validar_datos = True
End Function

Function genera_codigo(Sucursal As String, TipoEntrada As Integer, boleta As String, partida As Integer, TipoPrenda As Integer) As String

'genera_codigo = grupo & precio
'genera_codigo = genera_codigo & digitoverificador(genera_codigo)
End Function

Function digitoverificador(Codigo As String) As String
Dim i As Integer, x(7) As Integer, suma As Integer, residuo As Double, vc As Integer, dc As Integer
Dim length As Integer

length = Len(Trim(Codigo))

For i = 1 To 7
    x(i) = Mid(Trim(Codigo), i, 1)
Next i

suma = 0

For i = 1 To 7
    residuo = i Mod 2
    If residuo <> 0 Then vc = 3 Else vc = 1
    suma = suma + (x(i) * vc)
Next i

dc = 10 - (suma Mod 10)
If dc = 10 Then dc = 0

digitoverificador = dc
End Function

Private Sub txtSerie_GotFocus()
Cambiar_Color True, txtSerie
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 9) = txtSerie
    txtSerie.Visible = False
    grdArticulos.CancelEdit
    grdArticulos.SetFocus
End If
End Sub

Private Sub txtSerie_LostFocus()
Cambiar_Color False, txtSerie
End Sub

Function regresa_total() As Double
Dim Total As Double, i As Integer

Total = 0
For i = 1 To grdArticulos.Rows
    Total = Total + (grdArticulos.CellText(i, 6) * IIf(grdArticulos.CellText(i, 3) = "", 1, grdArticulos.CellText(i, 3)))
Next i
regresa_total = Total
End Function
