VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmExistencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Existencia de artículos"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExistencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   13740
   Begin VB.Frame Frame1 
      Caption         =   "Filtrar por:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   13695
      Begin VB.ComboBox cmbKilates 
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
         Left            =   8880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   285
         Width           =   3090
      End
      Begin VB.ComboBox cmbPrenda 
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
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   285
         Width           =   3015
      End
      Begin VB.ComboBox cmbTipo 
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
         Left            =   690
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   285
         Width           =   2535
      End
      Begin VB.ComboBox cmbKilataje 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   900
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cmbFiltrar 
         Height          =   315
         Left            =   1755
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   900
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton opKilataje 
         Appearance      =   0  'Flat
         Caption         =   "Kilataje"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3960
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.OptionButton opTipoPrenda 
         Appearance      =   0  'Flat
         Caption         =   "Tipo de Prenda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1470
      End
      Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
         Height          =   375
         Left            =   12075
         TabIndex        =   3
         Top             =   240
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
         Picture         =   "frmExistencias.frx":000C
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Kilataje/Marca:"
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
         Left            =   7320
         TabIndex        =   19
         Top             =   330
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Prenda:"
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
         Left            =   3360
         TabIndex        =   18
         Top             =   330
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         TabIndex        =   17
         Top             =   330
         Width           =   465
      End
   End
   Begin vbAcceleratorGrid6.vbalGrid grdExistencias 
      Height          =   6075
      Left            =   0
      TabIndex        =   4
      Top             =   885
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   10716
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
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
      DefaultRowHeight=   17
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   10680
         TabIndex        =   14
         Top             =   30
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   12540
      TabIndex        =   15
      Top             =   7440
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
      Picture         =   "frmExistencias.frx":0391
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   11385
      TabIndex        =   16
      Top             =   7440
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
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmExistencias.frx":08E3
   End
   Begin VB.Label lblTotalCosto 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Costo:"
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
      Index           =   1
      Left            =   4200
      TabIndex        =   20
      Top             =   7005
      Width           =   1005
   End
   Begin VB.Label lblTotalVitrina 
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
      Left            =   10560
      TabIndex        =   13
      Top             =   7005
      Width           =   1035
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "Total Precio Vitrina:"
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
      Left            =   7440
      TabIndex        =   6
      Top             =   7005
      Width           =   2475
   End
   Begin VB.Label lblCosto 
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
      Left            =   5640
      TabIndex        =   5
      Top             =   7005
      Width           =   1035
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   435
      Left            =   0
      TabIndex        =   7
      Top             =   6960
      Width           =   13695
   End
End
Attribute VB_Name = "frmExistencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbFiltrar_Click()
    
    If opKilataje.Value And cmbFiltrar.ListIndex > 0 Then
        
        cmbKilataje.Clear
        cmbKilataje.AddItem "(TODOS)", 0
        Cargar_Combos "Descripcion", "Kilatajes", cmbKilataje, " where IDTipo=" & cmbFiltrar.ItemData(cmbFiltrar.ListIndex), , False
        cmbKilataje.ListIndex = 0
        cmbKilataje.Visible = True
    
    ElseIf opKilataje.Value And cmbFiltrar.text = "(TODOS)" Then
        
        cmbKilataje.Visible = False
        cmbKilataje.Clear
    Else
        
        cmbKilataje.Visible = False
    End If

End Sub

Private Sub cmbFiltrar_GotFocus()
    Cambiar_Color True, cmbFiltrar
End Sub

Private Sub cmbFiltrar_LostFocus()
    Cambiar_Color False, cmbFiltrar
End Sub

Private Sub cmbKilataje_GotFocus()
    Cambiar_Color True, cmbKilataje
End Sub

Private Sub cmbKilataje_LostFocus()
    Cambiar_Color False, cmbKilataje
End Sub

Private Sub cmbKilates_GotFocus()
    Cambiar_Color True, cmbKilates
End Sub

Private Sub cmbKilates_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbKilates_LostFocus()
    Cambiar_Color False, cmbKilates
End Sub

Private Sub cmbPrenda_Click()
    
    If cmbPrenda.ListIndex > -1 Then
        
        If cmbTipo.ListIndex = 0 Then
            
            cmbKilates.Clear
            cmbKilates.AddItem "(TODOS)", 0
            Cargar_Combos "Descripcion", "kilatajes", cmbKilates, " WHERE IDTipo=1", "Ordenamiento", False
        
        Else
            
            cmbKilates.Clear
            cmbKilates.AddItem "(TODAS)", 0
            Cargar_Combos "Descripcion", "marcas", cmbKilates, , "Descripcion", False

        End If
        
    End If

End Sub

Private Sub cmbPrenda_GotFocus()
    Cambiar_Color True, cmbPrenda
End Sub

Private Sub cmbPrenda_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPrenda_LostFocus()
    Cambiar_Color False, cmbPrenda
End Sub

Private Sub cmbTipo_Click()

    If cmbTipo.ListIndex > -1 Then
        
        cmbKilates.Clear
        cmbPrenda.Clear
        cmbPrenda.AddItem "(TODAS)"
        Cargar_Combos "Descripcion", "tipoprenda", cmbPrenda, " WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Descripcion", False
    End If

End Sub

Private Sub cmbTipo_GotFocus()
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipo_LostFocus()
    Cambiar_Color False, cmbTipo
End Sub

Private Sub cmdBuscar_Click()
Dim Tipo As Integer

    Tipo = 0
    If cmbTipo.ListIndex = 0 Then
        
        Tipo = -1
    ElseIf cmbTipo.ListIndex > 0 Then
        
        Tipo = cmbTipo.ItemData(cmbTipo.ListIndex)
    End If
            
    Existencias Tipo, cmbPrenda.text, cmbKilates.text
    

'''''    If cmbFiltrar.ListIndex > -1 Then
'''''
'''''        If opTipoPrenda.Value And cmbFiltrar.text <> "(TODOS)" Then
'''''
'''''            Existencias " And detallesEntradaInventario.TipoPrenda=" & cmbFiltrar.ItemData(cmbFiltrar.ListIndex)
'''''
'''''        ElseIf opTipoPrenda.Value And cmbFiltrar.text = "(TODOS)" Then
'''''
'''''                Existencias
'''''
'''''        ElseIf cmbFiltrar.text = "(TODOS)" And cmbKilataje.ListIndex = -1 Then
'''''
'''''                Existencias
'''''
'''''        ElseIf cmbFiltrar.ListIndex > -1 And cmbKilataje.text = "(TODOS)" Then
'''''
'''''            Existencias " And detallesEntradaInventario.Tipo=" & cmbFiltrar.ItemData(cmbFiltrar.ListIndex)
'''''
'''''        Else
'''''
'''''            Existencias " And detallesEntradaInventario.Tipo=" & cmbFiltrar.ItemData(cmbFiltrar.ListIndex) & " And detallesentradainventario.Kilates=" & cmbKilataje.ItemData(cmbKilataje.ListIndex)
'''''
'''''        End If
'''''
'''''    End If
    
End Sub

Private Sub cmdImprimir_Click()
Dim Tipo As Integer
    
On Error GoTo Error

    If grdExistencias.Rows > 0 Then
                
        Tipo = 0
        If cmbTipo.ListIndex = 0 Then
            
            Tipo = -1
        ElseIf cmbTipo.ListIndex > 0 Then
            
            Tipo = cmbTipo.ItemData(cmbTipo.ListIndex)
        End If
    
        With frmMDI.Cr
            .Reset
            .WindowShowPrintSetupBtn = True
            .WindowShowExportBtn = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\Existencias.rpt"
            .SelectionFormula = "({detallesentradainventario.TipoEntrada}=" & ENTRADAALMONEDA & " OR {detallesentradainventario.TipoEntrada}=" & ENTRADACOMPRA & " OR {detallesentradainventario.TipoEntrada}=" & ENTRADADOTACION & ") AND {detallesentradainventario.Cantidad}>0" & IIf(Tipo = -1, "", " AND {detallesentradainventario.Tipo}=" & Tipo) & IIf(cmbPrenda.text <> "" And cmbPrenda.text <> "(TODAS)", " AND {detallesentradainventario.Descripcion}='" & cmbPrenda.text & "'", "") & IIf(Tipo > 1 And cmbKilates.text <> "" And cmbKilates.text <> "(TODAS)", " AND {detallesentradainventario.Marca}='" & cmbKilates.text & "'", IIf(Tipo = 1 And cmbKilates.text <> "" And cmbKilates.text <> "(TODOS)", " AND {kilatajes.Descripcion}='" & cmbKilates.text & "'", ""))
            .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
            .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
            .Formulas(2) = "Encabezado='A la fecha: " & Format(Date, "dd/mmm/yyyy") & "'"
            .Formulas(3) = "Iva=" & Regresa_Valor_BD("IvaVentas") & ""
            .DiscardSavedData = True
            .WindowState = crptMaximized
            .Destination = crptToWindow
            .WindowTitle = "Reporte de existencias"
            .Action = 1
        End With
        
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
    CentrarForm Me, frmMDI
End Sub

'Inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    lblCosto.Caption = ""
    lblTotalVitrina.Caption = ""
    cmbTipo.AddItem "(TODOS)"
    Cargar_Combos "Descripcion", "tipo", cmbTipo, , "ID", False
    cmbTipo.AddItem "AUTOS"
    Crear_Encabezados
    Poner_Flat Fl, Me.Controls, Me
    Screen.MousePointer = vbDefault
End Sub

'Creamos los encabezados
Private Sub Crear_Encabezados()

    With grdExistencias
        .AddColumn "K1", "Código", ecgHdrTextALignLeft, , 88, , , , , , , CCLSortString
        .AddColumn "K2", "Descripción", ecgHdrTextALignLeft, , 250, , , , , , , CCLSortString
        .AddColumn "K3", "Exis.", ecgHdrTextALignCentre, , 38, , , , , , , CCLSortNumeric
        .AddColumn "K4", "Kt", ecgHdrTextALignCentre, , 38, , , , , , , CCLSortString
        .AddColumn "K5", "Costo", ecgHdrTextALignRight, , 73, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Precio", ecgHdrTextALignRight, , 73, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Precio Vitrina", ecgHdrTextALignRight, , 75, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K8", "Total Costo", ecgHdrTextALignRight, , 85, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K9", "Total Precio V.", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K10", "Marca", ecgHdrTextALignLeft, , 84, , , , , , , CCLSortNumeric
        .AddColumn "K11", "Modelo", ecgHdrTextALignLeft, , 84, , , , , , , CCLSortNumeric
    End With
    
End Sub

'Mostramos las existencias
Private Sub Existencias(Tipo As Integer, Familia As String, Marca As String)
Dim rcInventario As New ADODB.Recordset
Dim Total As Double, TotalVitrina As Double, Iva As Double, strCondicion As String

On Error GoTo Error
    
    With rcInventario
        
        strCondicion = IIf(Tipo = -1, "", " AND d.Tipo=" & Tipo) & IIf(Familia <> "" And Familia <> "(TODAS)", " AND d.Descripcion='" & Familia & "'", "") & IIf(Tipo > 1 And Marca <> "" And Marca <> "(TODAS)", " AND d.Marca='" & Marca & "'", IIf(Tipo = 1 And Marca <> "" And Marca <> "(TODOS)", " AND kilatajes.Descripcion='" & Marca & "'", ""))
        
       ' rcInventario.Open "SELECT d.ID,d.PrecioVitrina,d.Codigo,d.Descripcion,d.Cantidad,d.Costo,d.Precio,d.TipoEntrada,d.Marca,d.Modelo,kilatajes.Descripcion AS Kilates " _
       '                    & "FROM detallesentradainventario d LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave WHERE d.Cantidad>0 AND (d.TipoEntrada=" & ENTRADAALMONEDA & " OR d.TipoEntrada=" & ENTRADACOMPRA & " OR d.TipoEntrada=" & ENTRADADOTACION & " OR d.TipoEntrada=" & ENTRADATRASPASO & ")" & strCondicion & " ORDER BY d.Codigo,d.Descripcion", dbDatos, adOpenForwardOnly, adLockReadOnly
        rcInventario.Open "SELECT d.ID,d.PrecioVitrina,d.Codigo,d.Descripcion,concat(c.nombre,' ',c.Apellido) as Nombre,d.Cantidad,d.Costo,d.Precio,d.TipoEntrada,d.Marca,d.Modelo,kilatajes.Descripcion AS Kilates,e.Fecha AS FechaEntrada " _
                         & "FROM detallesentradainventario d INNER JOIN entradainventario e ON d.IDEntrada=e.ID LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave LEFT JOIN empeno em on em.NumContrato=convert(mid(d.Codigo,5,6),UNSIGNED INTEGER) LEFT JOIN clientes c ON c.ID=em.IDCliente WHERE d.Cantidad>0 AND (d.TipoEntrada=" & D_VENTA & " OR d.TipoEntrada=" & ENTRADAALMONEDA & " OR d.TipoEntrada=" & ENTRADACOMPRA & " OR d.TipoEntrada=" & ENTRADADOTACION & ")" & strCondicion & " GROUP BY d.Codigo ORDER BY d.Codigo,d.Descripcion", dbDatos, adOpenForwardOnly, adLockReadOnly
        
        
        grdExistencias.Clear
        grdExistencias.Redraw = False
        Iva = Regresa_Valor_BD("IVAVentas") / 100
                
        While Not .EOF
            grdExistencias.AddRow
            grdExistencias.CellText(grdExistencias.Rows, 1) = !Codigo
            grdExistencias.CellItemData(grdExistencias.Rows, 1) = !ID
            grdExistencias.CellTextAlign(grdExistencias.Rows, 1) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdExistencias.CellText(grdExistencias.Rows, 2) = !Descripcion
            grdExistencias.CellTextAlign(grdExistencias.Rows, 2) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdExistencias.CellText(grdExistencias.Rows, 3) = !Cantidad
            grdExistencias.CellTextAlign(grdExistencias.Rows, 3) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdExistencias.CellText(grdExistencias.Rows, 4) = !Kilates
            grdExistencias.CellTextAlign(grdExistencias.Rows, 4) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdExistencias.CellText(grdExistencias.Rows, 5) = !Costo
            grdExistencias.CellTextAlign(grdExistencias.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdExistencias.CellText(grdExistencias.Rows, 6) = !Precio
            grdExistencias.CellTextAlign(grdExistencias.Rows, 6) = DT_RIGHT
            grdExistencias.CellText(grdExistencias.Rows, 7) = IIf(!TipoEntrada = ENTRADACOMPRA, !PrecioVitrina * (1 + Iva), !PrecioVitrina)
            grdExistencias.CellTextAlign(grdExistencias.Rows, 7) = DT_RIGHT
            
            grdExistencias.CellText(grdExistencias.Rows, 8) = !Costo
            grdExistencias.CellTextAlign(grdExistencias.Rows, 8) = DT_RIGHT
            grdExistencias.CellText(grdExistencias.Rows, 9) = !PrecioVitrina * (1 + Iva)
            grdExistencias.CellTextAlign(grdExistencias.Rows, 9) = DT_RIGHT
            grdExistencias.CellText(grdExistencias.Rows, 10) = !Marca
            grdExistencias.CellText(grdExistencias.Rows, 11) = !Modelo
            
            Total = Total + !Costo
            TotalVitrina = TotalVitrina + (!PrecioVitrina * (1 + Iva))
        .MoveNext
        Wend
        
        .Close
        Set rcInventario = Nothing
        grdExistencias.Redraw = True
        
    End With
        
    lblCosto.Caption = Format(Total, FMoneda)
    lblTotalVitrina.Caption = Format(TotalVitrina, FMoneda)
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcInventario = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdExistencias_ColumnClick(ByVal lCol As Long)
    
    If lCol = 1 Or lCol = 2 Then
        
        Ordenar_Grid lCol, grdExistencias, 1, 0
    End If

End Sub

Private Sub grdExistencias_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, sText As String
   
    If lCol <> 7 Then txtEdit.Visible = False: Exit Sub
    
    frmPasswords.ConexSuc = 0
    frmPasswords.DescuentoVentas = 0
    frmPasswords.Cancel = 0
    frmPasswords.Ventas = 0
    frmPasswords.ModificaPrecio = 0
    frmPasswords.ModificaCorte = 0
    frmPasswords.HacerCorte = 0
    frmPasswords.InteresRefrendo = 0
    frmPasswords.InteresDesempeño = 0
    frmPasswords.RecalculoPrecios = 0
    frmPasswords.AutorizaPrestamo = 0
    frmPasswords.Vencido = 0
    frmPasswords.CancelaCierre = 0
    frmPasswords.PrecioVitrina = 1
    
    If frmPasswords.Password(GERENTE, 1) = False Then txtEdit.Visible = False: Exit Sub
    
    grdExistencias.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
       
    iKeyAscii = Solo_Numeros(iKeyAscii, 1)
    If (iKeyAscii > 13) Then
        sText = Chr$(iKeyAscii) & sText
        txtEdit.text = sText
        txtEdit.SelStart = 1
        txtEdit.SelLength = Len(sText)
    Else
        txtEdit.text = grdExistencias.CellText(lRow, lCol)
        txtEdit.SelStart = 0
        txtEdit.SelLength = Len(sText)
    End If
       
    Set txtEdit.Font = grdExistencias.CellFont(lRow, lCol)
    If grdExistencias.CellBackColor(lRow, lCol) = -1 Then
        txtEdit.BackColor = grdExistencias.BackColor
    Else
        txtEdit.BackColor = grdExistencias.CellBackColor(lRow, lCol)
    End If
    
    txtEdit.Move lLeft + 40, lTop + 25, lWidth - 60, lHeight - 50
    txtEdit.Visible = True
    txtEdit.ZOrder
    txtEdit.SetFocus
End Sub

Private Sub opKilataje_Click()
    cmbFiltrar.Clear
    cmbFiltrar.AddItem "(TODOS)", 0
    Cargar_Combos "Descripcion", "Tipo", cmbFiltrar, " where Kilataje=1 or Peso=1", , False
    cmbFiltrar.ListIndex = 0
    cmbKilataje.ListIndex = -1
End Sub

Private Sub opTipoPrenda_Click()

'''''    cmbFiltrar.Clear
'''''    cmbFiltrar.AddItem "(TODOS)", 0
'''''    Cargar_Combos "Descripcion", "TipoPrenda", cmbFiltrar, , , False
'''''    cmbFiltrar.ListIndex = 0
'''''    cmbKilataje.Visible = False
End Sub

Private Sub txtEdit_GotFocus()
    Cambiar_Color True, txtEdit
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
Dim Precio As Double

On Error GoTo Error

    If KeyAscii = vbKeyReturn Then
    
        If Val(txtEdit.text) > 0 And txtEdit.text <> "" Then
            
            Precio = txtEdit.text
            dbDatos.Execute "update detallesentradainventario set PrecioVitrina=" & Precio & " where ID=" & grdExistencias.CellItemData(grdExistencias.SelectedRow, 1)
            grdExistencias.CellText(grdExistencias.SelectedRow, 7) = Precio
        End If
    
        txtEdit.Visible = False
    End If
    KeyAscii = Solo_Numeros(KeyAscii, 1)

Error:
    Maneja_Error Err
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then txtEdit.Visible = False
End Sub

Private Sub txtEdit_LostFocus()
    Cambiar_Color False, txtEdit
End Sub
