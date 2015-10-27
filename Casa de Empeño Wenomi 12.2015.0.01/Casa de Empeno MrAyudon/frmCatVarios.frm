VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCatVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo prendas varios"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   13665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatVarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   13665
   Begin VB.ComboBox cmbModelo 
      Height          =   315
      Left            =   8670
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   210
      Width           =   1695
   End
   Begin VB.ComboBox cmbMarca 
      Height          =   315
      Left            =   5790
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   210
      Width           =   1980
   End
   Begin VB.ComboBox cmbFamilia 
      Height          =   315
      Left            =   3045
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   1920
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      Left            =   585
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   1545
   End
   Begin vbAcceleratorGrid6.vbalGrid grdCatPrendas 
      Height          =   7635
      Left            =   0
      TabIndex        =   5
      Top             =   645
      Width           =   13620
      _ExtentX        =   24024
      _ExtentY        =   13467
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      GridLineColor   =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
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
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   375
      Left            =   10515
      TabIndex        =   4
      Top             =   165
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "    &Buscar"
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
      Picture         =   "frmCatVarios.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   11550
      TabIndex        =   10
      Top             =   165
      Width           =   1005
      _ExtentX        =   1773
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
      Picture         =   "frmCatVarios.frx":0391
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
      Height          =   375
      Left            =   12585
      TabIndex        =   11
      Top             =   165
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "    &Agregar"
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
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmCatVarios.frx":08E3
      PictureDisabled =   "frmCatVarios.frx":0C4D
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   7830
      TabIndex        =   9
      Top             =   255
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Marca:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   5070
      TabIndex        =   8
      Top             =   255
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Familia:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2205
      TabIndex        =   7
      Top             =   255
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   45
      TabIndex        =   6
      Top             =   255
      Width           =   495
   End
End
Attribute VB_Name = "frmCatVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim IDPrenda As Long, IDTipoPrenda As Integer

Private Sub cmbFamilia_Click()
    cmbMarca.Clear
    cmbMarca.AddItem ""
    Cargar_Combos "DISTINCT marcas.Descripcion", "tipoprenda INNER JOIN prendaselec ON tipoprenda.ID=prendaselec.IDFamilia INNER JOIN marcas on marcas.ID=prendaselec.IDMarca", cmbMarca, " WHERE prendaselec.IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND prendaselec.IDFamilia=" & cmbFamilia.ItemData(cmbFamilia.ListIndex), , False, "marcas.ID"

    cmbModelo.Clear
End Sub

Private Sub cmbFamilia_GotFocus()
    Cambiar_Color True, cmbFamilia
End Sub

Private Sub cmbFamilia_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbFamilia_LostFocus()
    Cambiar_Color False, cmbFamilia
End Sub

Private Sub cmbMarca_Click()
    cmbModelo.Clear
    cmbModelo.AddItem ""
    Cargar_Combos "DISTINCT prendaselec.Modelo", "tipoprenda INNER JOIN prendaselec ON tipoprenda.ID=prendaselec.IDFamilia INNER JOIN marcas on marcas.ID=prendaselec.IDMarca", cmbModelo, " WHERE prendaselec.IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND prendaselec.IDFamilia=" & cmbFamilia.ItemData(cmbFamilia.ListIndex) & " AND prendaselec.IDMarca = " & cmbMarca.ItemData(cmbMarca.ListIndex), , False, "marcas.ID"
End Sub

Private Sub cmbMarca_GotFocus()
    Cambiar_Color True, cmbMarca
End Sub

Private Sub cmbMarca_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbMarca_LostFocus()
    Cambiar_Color False, cmbMarca
End Sub

Private Sub cmbModelo_GotFocus()
    Cambiar_Color True, cmbModelo
End Sub

Private Sub cmbModelo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbModelo_LostFocus()
    Cambiar_Color False, cmbModelo
End Sub

Private Sub cmbTipo_Click()
    cmbFamilia.Clear
    cmbFamilia.AddItem ""
    Cargar_Combos "DISTINCT tipoprenda.Descripcion", "tipoprenda INNER JOIN prendaselec ON tipoprenda.ID=prendaselec.IDFamilia", cmbFamilia, " WHERE prendaselec.IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), , False, "tipoprenda.ID"
    
    cmbMarca.Clear
    cmbModelo.Clear
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

Private Sub cmdAgregar_Click()
Dim IDTipo As Integer, IDFamilia As Integer, IDMarca As Integer
    
    IDTipo = 0: IDFamilia = 0: IDMarca = 0
    IDTipo = cmbTipo.ItemData(cmbTipo.ListIndex)
    If cmbFamilia.ListIndex > -1 Then IDFamilia = cmbFamilia.ItemData(cmbFamilia.ListIndex)
    If cmbMarca.ListIndex > -1 Then IDMarca = cmbMarca.ItemData(cmbMarca.ListIndex)
    Hide
    Unload Me
    frmAgregaPrenda.Mostrar IDTipo, IDFamilia, IDMarca
End Sub

Private Sub cmdBuscar_Click()
Dim Tipo As Integer, Familia As Integer, Marca As Integer

    Tipo = 0
    Familia = 0
    Marca = 0
    
    If cmbTipo.ListIndex > -1 Then
        
        Tipo = cmbTipo.ItemData(cmbTipo.ListIndex)
    End If
    
    If cmbFamilia.ListIndex > -1 Then
        
        Familia = cmbFamilia.ItemData(cmbFamilia.ListIndex)
    End If
    
    If cmbMarca.ListIndex > -1 Then
        
        Marca = cmbMarca.ItemData(cmbMarca.ListIndex)
    End If
    
    CargaDatos Tipo, Familia, Marca, cmbModelo.text
End Sub

Private Sub cmdImprimir_Click()
Dim strFiltro As String
    
    If grdCatPrendas.Rows > 0 Then
        
        strFiltro = ""
        
        'Tipo
        strFiltro = "{prendaselec.IDTipo}=" & cmbTipo.ItemData(cmbTipo.ListIndex)
        
        'Familia
        If cmbFamilia.ListIndex > 0 Then
            
            strFiltro = strFiltro & " AND {prendaselec.IDFamilia}=" & cmbFamilia.ItemData(cmbFamilia.ListIndex)
        End If
        
        'Marca
        If cmbMarca.ListIndex > 0 Then
            
            strFiltro = strFiltro & " AND {prendaselec.IDMarca}=" & cmbMarca.ItemData(cmbMarca.ListIndex)
        End If
        
        'Modelo
        If cmbModelo.ListIndex > 0 Then
            
            strFiltro = strFiltro & " AND {prendaselec.Modelo}='" & cmbModelo.text & "'"
        End If
        
        With frmMDI.Cr
            .Reset
            .WindowShowPrintSetupBtn = True
            .WindowShowExportBtn = True
            .DiscardSavedData = True
            .ReportFileName = Path & "\Reportes\RepCatalogoPrendasVarios.rpt"
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .SelectionFormula = strFiltro
            .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
            .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .WindowTitle = "Catálogo prendas varios"
            .Action = 1
        End With
    
    End If
    
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'''''    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Cargar_Combos "Descripcion", "tipo", cmbTipo, " WHERE Kilataje=0 AND Peso=0", , False
    cmbTipo.ListIndex = ComboInformacion(cmbTipo, IDTipoPrenda)
    Crea_Encabezado
    Poner_Flat Fl, Me.Controls, Me
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Screen.MousePointer = vbDefault
End Sub

Sub Crea_Encabezado()
    
    With grdCatPrendas
        .AddColumn "C1", "Tipo", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
        .AddColumn "C2", "Familia", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "C3", "Marca", ecgHdrTextALignLeft, , 120, , , , , , , CCLSortString
        .AddColumn "C4", "Modelo", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
        .AddColumn "C5", "Funciones", ecgHdrTextALignLeft, , 152, , , , , , , CCLSortString
        .AddColumn "C6", "Características", ecgHdrTextALignLeft, , 152, , , , , , , CCLSortString
        .AddColumn "C7", "Mínimo", ecgHdrTextALignRight, , 84, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C8", "Máximo", ecgHdrTextALignRight, , 84, , , , , FMoneda, , CCLSortNumeric
    End With

End Sub

Sub CargaDatos(Tipo As Integer, Familia As Integer, Marca As Integer, Modelo As String)
Dim rcPrendas As New ADODB.Recordset
    
    With grdCatPrendas
        
        .Redraw = False
        .Clear
            
        rcPrendas.Open "SELECT prendaselec.ID,prendaselec.IDTipo,prendaselec.IDFamilia,prendaselec.IDMarca,prendaselec.Modelo,prendaselec.Funciones,prendaselec.Caracteristicas,prendaselec.Minimo,prendaselec.Maximo,tipo.Descripcion AS Desc_Tipo,marcas.Descripcion AS Desc_Marca,tipoprenda.Descripcion AS Desc_Familia FROM prendaselec LEFT JOIN tipo ON prendaselec.IDTipo=Tipo.ID LEFT JOIN marcas ON prendaselec.IDMarca=marcas.ID LEFT JOIN tipoprenda ON prendaselec.IDFamilia=tipoprenda.ID " & _
                        "WHERE prendaselec.IDTipo=" & Tipo & IIf(Familia > 0, " AND IDFamilia=" & Familia, "") & IIf(Marca > 0, " AND IDMarca=" & Marca, "") & IIf(Trim(Modelo) <> "", " AND Modelo='" & Modelo & "'", "") & " ORDER BY tipo.Descripcion,marcas.Descripcion,tipoprenda.Descripcion,prendaselec.Modelo", dbDatos, adOpenForwardOnly, adLockReadOnly
        
        While Not rcPrendas.EOF
            .AddRow
            .CellText(.Rows, 1) = rcPrendas!Desc_Tipo
            .CellItemData(.Rows, 1) = rcPrendas!IDTipo
            .CellText(.Rows, 2) = rcPrendas!Desc_Familia
            .CellItemData(.Rows, 2) = rcPrendas!IDMarca
            .CellText(.Rows, 3) = rcPrendas!Desc_Marca
            .CellItemData(.Rows, 3) = rcPrendas!IDFamilia
            .CellText(.Rows, 4) = rcPrendas!Modelo
            .CellItemData(.Rows, 4) = rcPrendas!ID
            .CellText(.Rows, 5) = rcPrendas!Funciones
            .CellText(.Rows, 6) = rcPrendas!Caracteristicas
            .CellText(.Rows, 7) = rcPrendas!Minimo
            .CellTextAlign(.Rows, 7) = DT_RIGHT
            .CellText(.Rows, 8) = rcPrendas!Maximo
            .CellTextAlign(.Rows, 8) = DT_RIGHT
            
            Colorea grdCatPrendas, .Rows, IIf(.Rows Mod 2 <> 0, Trim(RGB(242, 254, 255)), Trim(RGB(255, 255, 255)))
        
        rcPrendas.MoveNext
        Wend
        .Redraw = True
        
    End With
    rcPrendas.Close
    Set rcPrendas = Nothing

End Sub

Public Function Mostrar(IDTipo As Integer) As Long
    IDTipoPrenda = IDTipo
    Me.Show vbModal
    Mostrar = IDPrenda
End Function

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdCatPrendas_ColumnClick(ByVal lCol As Long)
    
    If lCol = 2 Or lCol = 3 Or lCol = 4 Then
            
        Ordenar_Grid lCol, grdCatPrendas, 1, 0
        grdCatPrendas.ClearSelection
    
    End If

End Sub

Private Sub grdCatPrendas_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    
    If grdCatPrendas.SelectedRow > 0 Then
        
        IDPrenda = grdCatPrendas.CellItemData(lRow, 4)
        Unload Me
    End If

End Sub

Sub Filtra(IDTipo As Long, Optional IDFamilia As Long = 0, Optional IDMarca As Long = 0)
Dim i As Long

    With grdCatPrendas
        
        .Redraw = False
            
            For i = 1 To .Rows
            
                If .CellItemData(i, 1) = IDTipo And IIf(IDMarca > 0, .CellItemData(i, 2) = IDMarca, .CellItemData(i, 1) = IDTipo) And IIf(IDFamilia > 0, .CellItemData(i, 3) = IDFamilia, .CellItemData(i, 1) = IDTipo) Then
                    
                    .RowVisible(i) = True
                    
                Else
                    
                    .RowVisible(i) = False
                End If
                
            Next i
            
        .Redraw = True
        
    End With

End Sub

Private Sub grdCatPrendas_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And grdCatPrendas.SelectedRow > 0 Then
        
        IDPrenda = grdCatPrendas.CellItemData(grdCatPrendas.SelectedRow, 4)
        Unload Me
    End If
    
End Sub
