VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~2.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmInventariofisico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Inventario físico"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10155
   Icon            =   "frmInventariofisico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   10155
   Begin vbAcceleratorGrid6.vbalGrid grdInventario 
      Height          =   4770
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8414
      RowMode         =   -1  'True
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
      DisableIcons    =   -1  'True
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9030
      TabIndex        =   1
      Top             =   4830
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
      Picture         =   "frmInventariofisico.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   4830
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Imprimir"
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
      Picture         =   "frmInventariofisico.frx":009D
   End
End
Attribute VB_Name = "frmInventariofisico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TipoDePrenda As Integer

Private Sub cmdImprimir_Click()
Dim Opcion As String

On Error GoTo Error

    If TipoDePrenda = -1 Then
        Opcion = ""
    Else
        Set rcTmp = dbDatos.Execute("select descripcion from tipo where id=" & TipoDePrenda & "")
        Opcion = rcTmp!Descripcion
    End If
    
    Screen.MousePointer = vbHourglass
    With frmMDI.Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\InventarioFisico.rpt"
        .SelectionFormula = "{Detallesentradainventario.Cantidad}>0" & IIf(TipoDePrenda = "-1", "", " and {DetallesEntradaInventario.Tipo}=" & Trim(TipoDePrenda) & "")
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='" & IIf(TipoDePrenda = "-1", "", "TIPO DE PRENDA - " & Opcion) & "'"
        .DiscardSavedData = True
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .WindowTitle = "Reporte de inventario físico"
        .Action = 1
    End With
Error:
    Maneja_Error Err
    Set rcTmp = Nothing
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Inicializar
End Sub

Public Sub Inicializar()
    Crear_Encabezados
    MuestraArticulos TipoDePrenda
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 5
End Sub

Public Sub MuestraArticulos(TipoPrenda As Integer)
Dim rcTmp As New ADODB.Recordset

On Error GoTo Error

    TipoDePrenda = TipoPrenda
    With grdInventario
        
        .Redraw = False
        .Clear
        
        rcTmp.Open "SELECT * FROM detallesentradainventario d WHERE d.Cantidad>0" & IIf(TipoPrenda = -1, "", " AND Tipo=" & Trim(TipoPrenda) & "") & " ORDER BY codigo,descripcion", dbDatos, adOpenForwardOnly, adLockReadOnly
        While Not rcTmp.EOF
            
            DoEvents
            .AddRow
            .CellText(.Rows, 1) = rcTmp!Codigo
            .CellTextAlign(.Rows, 1) = DT_LEFT
            .CellText(.Rows, 2) = rcTmp!Descripcion
            .CellTextAlign(.Rows, 2) = DT_LEFT
            .CellText(.Rows, 3) = SacaKilates(rcTmp!Kilates)
            .CellTextAlign(.Rows, 3) = DT_RIGHT
            .CellText(.Rows, 4) = rcTmp!Peso
            .CellTextAlign(.Rows, 4) = DT_RIGHT
            .CellText(.Rows, 5) = rcTmp!Costo
            .CellTextAlign(.Rows, 5) = DT_RIGHT
            .CellText(.Rows, 6) = rcTmp!Cantidad
            .CellTextAlign(.Rows, 6) = DT_RIGHT
        
        rcTmp.MoveNext
        Wend
        rcTmp.Close
        Set rcTmp = Nothing
        .Redraw = False
    
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Sub

Sub Crear_Encabezados()
With grdInventario
    .AddColumn "c1", "Código", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
    .AddColumn "c2", "Descripción", ecgHdrTextALignLeft, , 330, , , , , , , CCLSortString
    .AddColumn "c3", "Kilates", ecgHdrTextALignRight, , 55, , , , , , , CCLSortString
    .AddColumn "c4", "Peso", ecgHdrTextALignRight, , 55, , , , , "###0.000", , CCLSortString
    .AddColumn "c5", "Costo", ecgHdrTextALignRight, , 65, , , , , "##,###0.00", , CCLSortString
    .AddColumn "c6", "Existencia", ecgHdrTextALignRight, , 65, , , , , , , CCLSortNumeric
End With
End Sub
