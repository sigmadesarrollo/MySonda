VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmDotacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada a Inventario"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDotacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   13425
   Begin VB.TextBox txtConcepto 
      BorderStyle     =   0  'None
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
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1320
      Width           =   7335
   End
   Begin VB.TextBox txtCodigo 
      BorderStyle     =   0  'None
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
      Left            =   120
      MaxLength       =   13
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Frame frmEntradas 
      Caption         =   "ENTRADAS"
      Height          =   5805
      Left            =   0
      TabIndex        =   2
      Top             =   1650
      Width           =   13425
      Begin vbAcceleratorGrid6.vbalGrid grdArticulos 
         Height          =   4950
         Left            =   30
         TabIndex        =   3
         Top             =   240
         Width           =   13350
         _ExtentX        =   23548
         _ExtentY        =   8731
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
         Begin VB.ComboBox cmbTipo 
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   0
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.ComboBox dcbKilates 
            Height          =   315
            ItemData        =   "frmDotacion.frx":000C
            Left            =   5880
            List            =   "frmDotacion.frx":001F
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   5160
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtCosto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   6885
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   930
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
         Begin VB.TextBox txtArticulo 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox txtPesoo 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   9240
            TabIndex        =   4
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
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
         Left            =   9480
         TabIndex        =   16
         Top             =   5370
         Width           =   735
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
         Left            =   10140
         TabIndex        =   15
         Top             =   5370
         Width           =   1725
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
         TabIndex        =   14
         Top             =   5370
         Visible         =   0   'False
         Width           =   645
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
         TabIndex        =   13
         Top             =   5370
         Width           =   75
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
         Left            =   15
         TabIndex        =   17
         Top             =   5190
         Width           =   13380
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   4560
         Width           =   75
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
         TabIndex        =   9
         Top             =   4560
         Width           =   75
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   12240
      TabIndex        =   18
      Top             =   7575
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
      Picture         =   "frmDotacion.frx":003B
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   11040
      TabIndex        =   26
      Top             =   7575
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
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmDotacion.frx":00CC
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Concepto de entrada:"
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
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   2280
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Codigo:"
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
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   840
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
      Left            =   5880
      TabIndex        =   22
      Top             =   3480
      Visible         =   0   'False
      Width           =   675
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
      Left            =   8730
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
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
      Left            =   10650
      TabIndex        =   20
      Top             =   240
      Width           =   690
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
      Left            =   11520
      TabIndex        =   19
      Top             =   240
      Width           =   75
   End
End
Attribute VB_Name = "frmDotacion"
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

Private Function Validar() As Boolean
Dim i As Integer

Validar = True

For i = 1 To grdArticulos.Rows
    If grdArticulos.CellText(i, 1) = "" Or grdArticulos.CellText(i, 2) = "" Or grdArticulos.CellText(i, 3) = "" Or grdArticulos.CellText(i, 4) = "" Or grdArticulos.CellText(i, 5) = "" Or grdArticulos.CellText(i, 6) = "" Or grdArticulos.CellText(i, 7) = "" Or grdArticulos.CellText(i, 8) = "" Or grdArticulos.CellText(i, 9) = "" Or grdArticulos.CellText(i, 10) = "" Or grdArticulos.CellText(i, 11) = "" Then
        If grdArticulos.CellText(i, 1) = "" And grdArticulos.CellText(i, 2) = "" And grdArticulos.CellText(i, 3) = "" And grdArticulos.CellText(i, 4) = "" And grdArticulos.CellText(i, 5) = "" And grdArticulos.CellText(i, 6) = "" And grdArticulos.CellText(i, 7) = "" And grdArticulos.CellText(i, 8) = "" And grdArticulos.CellText(i, 9) = "" And grdArticulos.CellText(i, 10) = "" And grdArticulos.CellText(i, 11) = "" Then GoTo 125
               
        If grdArticulos.CellText(i, 6) = "" Then MsgBox "Introduzca la descripción del artículo !!", vbInformation, "Entrada a Inventario": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 6, 13, False: Exit Function
        If grdArticulos.CellText(i, 7) = "" Then MsgBox "Introduzca la cantidad de artículos !!", vbInformation, "Entrada a Inventario": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 7, 13, False: Exit Function
        If grdArticulos.CellText(i, 8) = "" And grdArticulos.CellText(i, 5) = "ORO" Then MsgBox "Seleccione el kilataje del artículo !!", vbInformation, "Entrada a Inventario": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 8, 13, False: Exit Function
        If grdArticulos.CellText(i, 9) = "" And grdArticulos.CellText(i, 5) = "ORO" Then MsgBox "Introduzca el peso del artículo !!", vbInformation, "Entrada a Inventario": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 9, 13, False: Exit Function
        If grdArticulos.CellText(i, 10) = "" Then MsgBox "Introduzca el costo del artículo !!", vbInformation, "Entrada a Inventario": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 10, 13, False: Exit Function
        If grdArticulos.CellText(i, 11) = "" Then MsgBox "Introduzca el precio del artículo !!", vbInformation, "Entrada a Inventario": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 11, 13, False: Exit Function
    End If
125:
Next i
End Function

Private Sub cmbTipo_Click()
    If cmbTipo.ListIndex = -1 Then Exit Sub
    grdArticulos.CellText(grdArticulos.SelectedRow, 5) = cmbTipo.text
    grdArticulos.CellItemData(grdArticulos.SelectedRow, 5) = cmbTipo.ItemData(cmbTipo.ListIndex)
    cmbTipo.Visible = False
    grdArticulos.SetFocus
    'grdArticulos_CancelEdit
End Sub

Private Sub cmbTipo_GotFocus()
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmbTipo.ListIndex > -1 Then
        grdArticulos.CellText(grdArticulos.SelectedRow, 5) = cmbTipo.text
        grdArticulos.CellItemData(grdArticulos.SelectedRow, 5) = cmbTipo.ItemData(cmbTipo.ListIndex)
        cmbTipo.Visible = False
        grdArticulos.CancelEdit
        grdArticulos.SetFocus
    End If
End Sub

Private Sub cmbTipo_LostFocus()
    Cambiar_Color False, cmbTipo
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass

    If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Entrada a Inventario") = vbYes Then
        If Validar_Datos Then If Validar Then Grabar_Entradas
    End If

    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub dcbKilates_Click()
Dim Cantidad As Integer, Peso As Double, Total As Double, Kilataje As String, Kilates As String

If dcbKilates.ListIndex = -1 Then Exit Sub

If dcbKilates.text <> "" Then Kilates = dcbKilates.text Else Kilates = grdArticulos.CellText(grdArticulos.SelectedRow, 8)
Kilataje = dcbKilates.ItemData(dcbKilates.ListIndex)

grdArticulos.CellText(grdArticulos.SelectedRow, 8) = dcbKilates.text
grdArticulos.CellItemData(grdArticulos.SelectedRow, 8) = RegresaKilates(grdArticulos.CellText(grdArticulos.SelectedRow, 8))

If grdArticulos.CellText(grdArticulos.SelectedRow, 7) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 7) Else Cantidad = 0
If grdArticulos.CellText(grdArticulos.SelectedRow, 9) <> "" Then Peso = grdArticulos.CellText(grdArticulos.SelectedRow, 9) Else Peso = 0

If grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "" Then
    Set rcConsulta = dbDatos.Execute("select " & "Venta" & dcbKilates.text & " as costo from parametros")
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        Total = (Peso * rcConsulta!Costo)
        grdArticulos.CellText(grdArticulos.SelectedRow, 10) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 10) = DT_RIGHT
    End If
Else
    Total = 0
    grdArticulos.CellText(grdArticulos.SelectedRow, 10) = Total
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 10) = DT_RIGHT
End If

'grdArticulos.CancelEdit
dcbKilates.Visible = False
lblTotal.Caption = Format(Regresa_Total, "##,###0.00")
grdArticulos.SetFocus
Set rcConsulta = Nothing
End Sub

Private Sub dcbKilates_GotFocus()
dcbKilates.BackColor = &HC0FFFF
End Sub

Private Sub dcbKilates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then dcbKilates.Visible = False
End Sub

Private Sub dcbKilates_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 8) = dcbKilates.text
    grdArticulos.CellItemData(grdArticulos.SelectedRow, 8) = RegresaKilates(dcbKilates.text)
    dcbKilates.Visible = False
    grdArticulos.SetFocus
End If
End Sub

Private Sub dcbKilates_LostFocus()
dcbKilates.BackColor = vbWhite
End Sub

Private Sub Form_Load()
Inicializar
End Sub

'Inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    frmEntradas.BorderStyle = 0
    lblFecha.Caption = Format(Date, "DD/MMM/YYYY")
    Crear_Encabezados
    Cargar_Combos "Descripcion", "Tipo", cmbTipo
    Poner_Flat Fl, Me.Controls, Me
    lblTotal.Caption = "0.00"
    lblFolio.Caption = Regresa_Movimiento(False, "FolioInventario")
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

'Creamos los encabezados
Private Sub Crear_Encabezados()

    With grdArticulos
        .AddColumn "K1", "Codigo", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "K2", "Sucursal", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
        .AddColumn "K3", "Boleta", ecgHdrTextALignRight, , 70, , , , , , , CCLSortString
        .AddColumn "K4", "Partida", ecgHdrTextALignRight, , 50, , , , , , , CCLSortString
        .AddColumn "K6", "Tipo", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "K7", "Descripción", ecgHdrTextALignLeft, , 188, , , , , , , CCLSortString
        .AddColumn "K8", "Cantidad", ecgHdrTextALignRight, , 55, , , , , , , CCLSortString
        .AddColumn "K9", "Kilates", ecgHdrTextALignLeft, , 62, , , , , , , CCLSortNumeric
        .AddColumn "K10", "Peso", ecgHdrTextALignRight, , 50, , , , , "0.000", , CCLSortNumeric
        .AddColumn "K11", "Costo", ecgHdrTextALignRight, , 65, , , , , "###,###,###,##0.00", , CCLSortString
        .AddColumn "K12", "Precio", ecgHdrTextALignRight, , 75, , , , , "###,###,###,##0.00", , CCLSortNumeric
        .GridLines = True
        .Rows = 20
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl()
End Sub

Private Sub grdArticulos_CancelEdit()
    txtCantidad.Visible = False
    txtCosto.Visible = False
    txtPrecioo.Visible = False
    txtArticulo.Visible = False
    dcbKilates.Visible = False
    'dcbKilates.ListIndex = -1
    cmbTipo.Visible = False
    'cmbTipo.ListIndex = -1
    txtPesoo.Visible = False
End Sub

Private Sub grdArticulos_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
If grdArticulos.SelectedRow > 0 Then
      If KeyCode = vbKeyDelete Then
       If MsgBox("Desea eliminar el articulo seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Dotación a Inventario") = vbYes Then

            grdArticulos.RemoveRow grdArticulos.SelectedRow
            lblTotal.Caption = Format(Regresa_Total, "##,###0.00")
            grdArticulos.CancelEdit
        End If
      End If
   End If
End Sub

Private Sub grdArticulos_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, sText As String, obj As Object
   
    txtCantidad.Visible = False
    grdArticulos_CancelEdit
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 5) <> "ORO" And (lCol = 8 Or lCol = 9) Then Exit Sub
    If lCol < 6 And lCol <> 5 Then Exit Sub
    
    Select Case lCol
    Case 5: Set obj = cmbTipo
    Case 6: Set obj = txtArticulo
    Case 7: Set obj = txtCantidad
    Case 8: Set obj = dcbKilates
    Case 9: Set obj = txtPesoo
    Case 10: Set obj = txtCosto
    Case 11: Set obj = txtPrecioo
    End Select
   
    grdArticulos.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight

    If Not IsMissing(grdArticulos.CellText(lRow, lCol)) Then
        sText = grdArticulos.CellText(lRow, lCol)
    Else
        sText = ""
    End If
   
    If lCol <> 8 And lCol <> 5 Then
        If lCol = 6 Then
            obj.Alignment = vbLeftJustify
            grdArticulos.CellTextAlign(grdArticulos.SelectedRow, lCol) = DT_LEFT
        Else
            obj.Alignment = vbRightJustify
        End If

        If (iKeyAscii > 13) Then
            sText = Chr$(iKeyAscii) & sText
            obj.text = UCase(sText)
            obj.SelStart = 1
            obj.SelLength = Len(sText)
        Else
            obj.text = sText
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
   
    If lCol <> 8 And lCol <> 5 Then
        obj.Move lLeft + 40, lTop + 25, lWidth - 60, lHeight - 50
    Else
        obj.Move lLeft + 40, lTop + 25, lWidth - 60
    End If
   
    obj.Visible = True
    obj.ZOrder
    obj.SetFocus
End Sub

Private Sub txtArticulo_GotFocus()
txtArticulo.BackColor = &HC0FFFF
End Sub

Private Sub txtArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then txtArticulo.Visible = False
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 6) = txtArticulo.text
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_LEFT
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
If KeyCode = vbKeyEscape Then txtCantidad.Visible = False
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
Dim i As Integer, Total As Double, Cantidad As Integer, Precio As Double, x As Integer, Peso As Double, Kilataje As String

KeyAscii = Solo_Numeros(KeyAscii)

If KeyAscii = vbKeyReturn Then
   
   Kilataje = RegresaKilates(IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 5) = "ORO", grdArticulos.CellText(grdArticulos.SelectedRow, 8), ""), grdArticulos.CellText(grdArticulos.SelectedRow, 5))

   Total = 0
   Cantidad = 0
   Precio = 0
   
   If txtCantidad.text <> "" Then grdArticulos.CellText(grdArticulos.SelectedRow, 7) = txtCantidad.text Else grdArticulos.CellText(grdArticulos.SelectedRow, 7) = "": GoTo 125
   grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 7) = DT_RIGHT
   For i = 1 To grdArticulos.Rows
        Cantidad = IIf(grdArticulos.CellText(i, 7) = "", 0, Val(grdArticulos.CellText(i, 7)))
        If grdArticulos.CellText(i, 11) <> "" Then Precio = grdArticulos.CellText(i, 11) Else Precio = 0
        Total = Total + (Cantidad * Precio)
   Next i

    If grdArticulos.CellText(grdArticulos.SelectedRow, 7) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 7) Else Cantidad = 0
    If grdArticulos.CellText(grdArticulos.SelectedRow, 9) <> "" Then Peso = grdArticulos.CellText(grdArticulos.SelectedRow, 9) Else Peso = 0
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "Fino" Then
        Set rcConsulta = dbDatos.Execute("select " & "Venta" & IIf(dcbKilates.text = "", grdArticulos.CellText(grdArticulos.SelectedRow, 8), dcbKilates.text) & " as costo from parametros")
        If Not rcConsulta.BOF And Not rcConsulta.EOF Then
            Total = (Peso * rcConsulta!Costo) '* Cantidad
        End If
        grdArticulos.CellText(grdArticulos.SelectedRow, 10) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 10) = DT_RIGHT
    Else
        Total = 0
        grdArticulos.CellText(grdArticulos.SelectedRow, 10) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 10) = DT_RIGHT
    End If
   '--------

125:
   grdArticulos.CancelEdit
   txtCantidad.Visible = False
   lblTotal.Caption = Format(Regresa_Total, "###,###,###,###0.00")
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

Private Sub txtCodigo_GotFocus()
Seleccionar_Texto txtCodigo
Cambiar_Color True, txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
KeyAscii = Solo_Numeros(KeyAscii)
If KeyAscii = vbKeyReturn Then
    If Len(txtCodigo.text) = 13 Then AgregaArticulo txtCodigo.text
End If
Pasar_Foco KeyAscii
End Sub

Private Sub txtCodigo_LostFocus()
Cambiar_Color False, txtCodigo
End Sub

Private Sub txtConcepto_GotFocus()
Seleccionar_Texto txtConcepto
Cambiar_Color True, txtConcepto
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtConcepto_LostFocus()
Cambiar_Color False, txtConcepto
End Sub

Private Sub txtCosto_GotFocus()
txtCosto.BackColor = &HC0FFFF
End Sub

Private Sub txtCosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then txtCosto.Visible = False
End Sub

Private Sub txtCosto_KeyPress(KeyAscii As Integer)
Dim Kilataje As String, Cantidad As Integer, Peso As Double, Total As Double

On Error GoTo error

KeyAscii = Solo_Numeros(KeyAscii, 1)
If KeyAscii = vbKeyReturn Then
   
    Kilataje = RegresaKilates(IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 5) = "ORO", grdArticulos.CellText(grdArticulos.SelectedRow, 8), ""), grdArticulos.CellText(grdArticulos.SelectedRow, 5))
   
    grdArticulos.CellText(grdArticulos.SelectedRow, 10) = txtCosto.text
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 10) = DT_RIGHT
    grdArticulos.CancelEdit
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 7) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 7) Else Cantidad = 0
    If grdArticulos.CellText(grdArticulos.SelectedRow, 9) <> "" Then Peso = grdArticulos.CellText(grdArticulos.SelectedRow, 9) Else Peso = 0
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "Fino" And grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "Plata" Then
        Set rcConsulta = dbDatos.Execute("select " & "Venta" & IIf(dcbKilates.text = "", grdArticulos.CellText(grdArticulos.SelectedRow, 8), dcbKilates.text) & " as costo from parametros")
        If Not rcConsulta.BOF And Not rcConsulta.EOF Then
            Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 10) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 10))
            grdArticulos.CellText(grdArticulos.SelectedRow, 10) = Total
        End If
    Else
        
        Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 10) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 10))
        grdArticulos.CellText(grdArticulos.SelectedRow, 10) = Total
    End If
    
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 10) = DT_RIGHT
    txtCosto.Visible = False
    lblTotal.Caption = Format(Regresa_Total, "###,###,###,###0.00")
    KeyAscii = 0
    grdArticulos.CancelEdit
    grdArticulos.SetFocus
ElseIf KeyAscii = vbKeyEscape Then
    grdArticulos.CancelEdit
    grdArticulos.SetFocus
    KeyAscii = 0
End If

error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Private Sub txtCosto_LostFocus()
txtCosto.BackColor = vbWhite
End Sub

Private Sub txtPesoo_GotFocus()
txtPesoo.BackColor = &HC0FFFF
End Sub

Private Sub txtPesoo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then txtPesoo.Visible = False
End Sub

Private Sub txtPesoo_KeyPress(KeyAscii As Integer)
Dim Peso As Double, Cantidad As Integer, Costo As Double, i As Integer
Dim Total As Double

KeyAscii = Solo_Numeros(KeyAscii, 1)
If KeyAscii = vbKeyReturn Then
    
    If txtPesoo.text <> "" Then grdArticulos.CellText(grdArticulos.SelectedRow, 9) = txtPesoo.text: Peso = txtPesoo.text Else grdArticulos.CellText(grdArticulos.SelectedRow, 9) = "": Peso = 0
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 9) = DT_RIGHT

    If grdArticulos.CellText(grdArticulos.SelectedRow, 7) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 7) Else Cantidad = 0
    If grdArticulos.CellText(grdArticulos.SelectedRow, 9) <> "" Then Peso = grdArticulos.CellText(grdArticulos.SelectedRow, 9) Else Peso = 0
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "Fino" Then
        Set rcConsulta = dbDatos.Execute("select " & "Venta" & IIf(dcbKilates.text = "", grdArticulos.CellText(grdArticulos.SelectedRow, 8), dcbKilates.text) & " as costo from parametros")
        If Not rcConsulta.BOF And Not rcConsulta.EOF Then
            Total = (Peso * rcConsulta!Costo) '* Cantidad
        End If
        grdArticulos.CellText(grdArticulos.SelectedRow, 10) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 10) = DT_RIGHT
    Else
        Total = 0
        grdArticulos.CellText(grdArticulos.SelectedRow, 10) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 10) = DT_RIGHT
    End If
   
    grdArticulos.CancelEdit
    txtPesoo.Visible = False
    lblTotal.Caption = Format(Regresa_Total, "###,###,###,###0.00")
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

'Grabamos la Entrada
Private Sub Grabar_Entradas()
    Grabar_Encabezado
    Sleep 1000
    Limpiar
    txtCodigo.SetFocus
End Sub

Private Sub Limpiar()
    grdArticulos.Clear
    grdArticulos.GridLines = True
    grdArticulos.Rows = 20
    lblTotal.Caption = "0.00"
    txtCodigo.text = ""
    txtConcepto.text = ""
End Sub

'Grabamos el encabezado de la entrada
Private Function Grabar_Encabezado() As Long
Dim rcID As New ADODB.Recordset
Dim Folio As Long
   
On Error GoTo error
    
    Folio = Regresa_Movimiento(False, "FolioInventario")
    Regresa_Movimiento True, "FolioInventario"
   
    dbDatos.Execute "INSERT INTO entradainventario (Folio,Fecha,IDUsuario,IDSucursal,TipoEntrada) VALUES (" & Folio & ",'" & Format(lblFecha.Caption, "YYYY/MM/DD") & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & ENTRADADOTACION & ")"
    rcID.Open "SELECT MAX(ID) AS IDD FROM entradainventario", dbDatos, adOpenForwardOnly, adLockOptimistic
   
    Grabar_Inventario rcID!idd, Folio
    
    ImprimirEntradas rcID!idd
    
    rcID.Close
   
error:
    Maneja_Error Err
    Set rcID = Nothing
End Function

'Grabamos el inventario
Private Sub Grabar_Inventario(ID As Long, Folio As Long)
Dim Indice As Integer, Movimiento As Long, crImporte As Double, Kilataje As Integer, Codigo As String

On Error GoTo error
    
    'Tomo el Importe
    crImporte = lblTotal.Caption
    
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
   
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "'," & Movimiento & "," & Folio & ",'EN01','620301'," & crImporte & "," & TIPO_CARGO & ",1,'" & Trim(txtConcepto.text) & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabamos el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "'," & Movimiento & "," & Folio & ",'EN50','200950'," & crImporte & "," & TIPO_ABONO & ",1,'" & Trim(txtConcepto.text) & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
           
    For Indice = 1 To grdArticulos.Rows
    
        If grdArticulos.CellText(Indice, 1) = "" And grdArticulos.CellText(Indice, 2) = "" And grdArticulos.CellText(Indice, 3) = "" And grdArticulos.CellText(Indice, 4) = "" And grdArticulos.CellText(Indice, 5) = "" And grdArticulos.CellText(Indice, 6) = "" And grdArticulos.CellText(Indice, 7) = "" And grdArticulos.CellText(Indice, 8) = "" And grdArticulos.CellText(Indice, 9) = "" And grdArticulos.CellText(Indice, 10) = "" And grdArticulos.CellText(Indice, 11) = "" Then GoTo 126
        
        Kilataje = 0
        Kilataje = RegresaKilates(IIf(grdArticulos.CellText(Indice, 5) = "ORO", grdArticulos.CellText(Indice, 8), ""), grdArticulos.CellText(Indice, 5))
        Codigo = grdArticulos.CellText(Indice, 1)
        
        dbDatos.Execute "INSERT INTO detallesentradainventario (IDEntrada,Codigo,Descripcion,Kilates,Peso,Costo,Precio,Cantidad,Tipo,SucursalOrigen,TipoEntrada,PrecioVitrina) VALUES (" & _
                         ID & ",'" & Trim(Codigo) & "','" & Trim(grdArticulos.CellText(Indice, 6)) & "'," & Kilataje & "," & _
                         CDbl(grdArticulos.CellText(Indice, 9)) & "," & CDbl(grdArticulos.CellText(Indice, 10)) & "," & CDbl(grdArticulos.CellText(Indice, 11)) & "," & grdArticulos.CellText(Indice, 7) & "," & _
                         grdArticulos.CellItemData(Indice, 5) & "," & frmMDI.IDSucursal & ", " & ENTRADADOTACION & "," & CDbl(grdArticulos.CellText(Indice, 11)) & ")"
         
126:
      Next Indice
   
error:
   Maneja_Error Err
End Sub

Private Sub txtPrecioo_GotFocus()
txtPrecioo.BackColor = &HC0FFFF
End Sub

Private Sub txtPrecioo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then txtPrecioo.Visible = False
End Sub

Private Sub txtPrecioo_KeyPress(KeyAscii As Integer)
Dim i As Integer, Cantidad As Integer, Precio As Double, Total As Double

KeyAscii = Solo_Numeros(KeyAscii, 1)
If KeyAscii = vbKeyReturn Then
   
    grdArticulos.CellText(grdArticulos.SelectedRow, 11) = txtPrecioo.text
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 11) = DT_RIGHT
    grdArticulos.CancelEdit
        
    If grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "Fino" And grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "Plata" Then
        Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 11) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 11))
    Else
        Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 11) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 11))
    End If

    txtCosto.Visible = False
    lblTotal.Caption = Format(Regresa_Total, "###,###,###,###0.00")
    KeyAscii = 0
    txtPrecioo.Visible = False
    grdArticulos.CancelEdit
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


Private Function Validar_Datos() As Boolean
Dim i As Integer, x As Integer

    If txtConcepto.text = "" Then
        MsgBox "Introduzca el concepto de entrada !!", vbInformation, "Entrada de Inventario"
        Validar_Datos = False
        txtConcepto.SetFocus
        Exit Function
    End If

    x = 0

    For i = 1 To grdArticulos.Rows

        If grdArticulos.CellText(i, 1) = "" Or grdArticulos.CellText(i, 2) = "" Or grdArticulos.CellText(i, 3) = "" Or grdArticulos.CellText(i, 4) = "" Or Trim(grdArticulos.CellText(i, 5)) = "" Or Trim(grdArticulos.CellText(i, 6)) = "" Or grdArticulos.CellText(i, 7) = "" Or grdArticulos.CellText(i, 8) = "" Or grdArticulos.CellText(i, 9) = "" Or grdArticulos.CellText(i, 10) = "" Or grdArticulos.CellText(i, 11) = "" Then
            If grdArticulos.CellText(i, 1) = "" And grdArticulos.CellText(i, 2) = "" And grdArticulos.CellText(i, 3) = "" And grdArticulos.CellText(i, 4) = "" And grdArticulos.CellText(i, 5) = "" And grdArticulos.CellText(i, 6) = "" And grdArticulos.CellText(i, 7) = "" And grdArticulos.CellText(i, 8) = "" And grdArticulos.CellText(i, 9) = "" And grdArticulos.CellText(i, 10) = "" And grdArticulos.CellText(i, 11) = "" Then x = x + 1
        End If

    Next i

    If x = grdArticulos.Rows Then MsgBox "Introduzca los artículos a dotar !!", vbInformation, "Entrada a Inventario": Validar_Datos = False: txtCodigo.SetFocus Else Validar_Datos = True
End Function

Function Regresa_Total() As Double
    Dim Total As Double, i As Integer

    Total = 0

    For i = 1 To grdArticulos.Rows
        Total = Total + (grdArticulos.CellText(i, 10) * IIf(grdArticulos.CellText(i, 7) = "", 1, grdArticulos.CellText(i, 7)))
    Next i

    Regresa_Total = Total
End Function

Function AgregaArticulo(Codigo As String)
Dim Sucursal As String, Boleta As String, Partida As Integer, TipoPrenda As Integer, Renglon As Integer, i As Integer
Dim NombreSucursal As String, TipoEntrada As Integer

    Sucursal = Mid(Codigo, 1, 3)
    TipoEntrada = Mid(Codigo, 4, 1)
    Boleta = Mid(Codigo, 5, 6)
    Partida = Mid(Codigo, 11, 2)

    AgregaArticulo = CreaCodigoBarras(Sucursal, TipoEntrada, Boleta, Partida)
    
    With grdArticulos
        
        For i = 1 To .Rows

            If Val(.CellText(i, 1)) = 0 Then Exit For
        Next i
    
        .CellText(i, 1) = AgregaArticulo
        .CellText(i, 2) = SacaValor("sucursales", "NombreComercial", " WHERE Activa=1")
        .CellText(i, 3) = Val(Boleta)
        .CellTextAlign(i, 3) = DT_RIGHT
        .CellText(i, 4) = Partida
        .CellTextAlign(i, 4) = DT_RIGHT
    End With

End Function

Function ImprimirEntradas(IDEntrada As Long)

    With frmMDI.Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\EntradasInventario.rpt"
        .SelectionFormula = "{detallesentradainventario.IDEntrada}=" & IDEntrada & ""
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado=''"
        .WindowTitle = "Reporte Entradas a Inventario"
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With

End Function
