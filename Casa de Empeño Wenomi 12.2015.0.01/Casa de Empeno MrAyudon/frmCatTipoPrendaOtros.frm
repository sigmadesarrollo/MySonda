VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCatTipoPrendaOtros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo Prendas Varios"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatTipoPrendaOtros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   12105
   Begin VB.TextBox txtModelo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5775
      TabIndex        =   3
      Top             =   510
      Width           =   2760
   End
   Begin VB.ComboBox cmbFamilia 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCatTipoPrendaOtros.frx":000C
      Left            =   1635
      List            =   "frmCatTipoPrendaOtros.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   3120
   End
   Begin VB.ComboBox cmbMarca 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCatTipoPrendaOtros.frx":0010
      Left            =   5730
      List            =   "frmCatTipoPrendaOtros.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   135
      Width           =   2850
   End
   Begin VB.TextBox txtCaracteristicas 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1680
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1395
      Width           =   6855
   End
   Begin VB.TextBox txtFunciones 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1680
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   855
      Width           =   6855
   End
   Begin VB.TextBox txtMaximo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6135
      TabIndex        =   7
      Top             =   1920
      Width           =   2400
   End
   Begin VB.TextBox txtMinimo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   1920
      Width           =   2400
   End
   Begin VB.ComboBox cmbTipo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCatTipoPrendaOtros.frx":0014
      Left            =   1635
      List            =   "frmCatTipoPrendaOtros.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   3120
   End
   Begin vbAcceleratorGrid6.vbalGrid grdCatPrendas 
      Height          =   6525
      Left            =   30
      TabIndex        =   8
      Top             =   2325
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   11509
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
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10110
      TabIndex        =   9
      Top             =   585
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
      Picture         =   "frmCatTipoPrendaOtros.frx":0018
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   120
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
      Object.ToolTipText     =   ""
      Picture         =   "frmCatTipoPrendaOtros.frx":056A
   End
   Begin DevPowerFlatBttn.FlatBttn cmdEliminar 
      Height          =   375
      Left            =   10110
      TabIndex        =   13
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Eliminar"
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
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmCatTipoPrendaOtros.frx":0ABC
      PictureDisabled =   "frmCatTipoPrendaOtros.frx":100E
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCancelar 
      Height          =   375
      Left            =   8880
      TabIndex        =   20
      Top             =   585
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Cancelar"
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
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmCatTipoPrendaOtros.frx":1BE0
      PictureDisabled =   "frmCatTipoPrendaOtros.frx":1E2F
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   9360
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   2
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "        &Imprimir"
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
      Picture         =   "frmCatTipoPrendaOtros.frx":2A01
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
      Left            =   4920
      TabIndex        =   19
      Top             =   540
      Width           =   780
   End
   Begin VB.Label Label28 
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
      Index           =   0
      Left            =   4920
      TabIndex        =   18
      Top             =   195
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Características:"
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
      Left            =   75
      TabIndex        =   17
      Top             =   1455
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Funciones:"
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
      Left            =   75
      TabIndex        =   16
      Top             =   855
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Máximo:"
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
      Left            =   5160
      TabIndex        =   15
      Top             =   1965
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mínimo:"
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
      Left            =   75
      TabIndex        =   14
      Top             =   1965
      Width           =   780
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
      Index           =   0
      Left            =   75
      TabIndex        =   12
      Top             =   540
      Width           =   795
   End
   Begin VB.Label Label28 
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
      Index           =   30
      Left            =   75
      TabIndex        =   11
      Top             =   195
      Width           =   495
   End
End
Attribute VB_Name = "frmCatTipoPrendaOtros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbFamilia_Click()
Dim IDFamilia As Integer

    If cmbFamilia.text = "[1. AGREGAR FAMILIA]" And cmbTipo.ListIndex > -1 Then
        
        IDFamilia = 0
        IDFamilia = frmAgregaFamilia.Mostrar(cmbTipo.ItemData(cmbTipo.ListIndex))
        If IDFamilia > 0 Then
            
            cmbFamilia.Clear
            cmbFamilia.AddItem "[1. AGREGAR FAMILIA]"
            Cargar_Combos "tipoprenda.Descripcion", "tipoprenda INNER JOIN tipo ON tipoprenda.IDTipo=tipo.ID", cmbFamilia, " WHERE tipo.Kilataje=0 AND tipo.Peso=0", , False, "tipoprenda.ID"
            cmbFamilia.ListIndex = ComboInformacion(cmbFamilia, IDFamilia)
        
        Else
            
            cmbFamilia.ListIndex = -1
        End If
    
    End If
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
Dim IDMarca As Integer

    If cmbMarca.text = "[1. AGREGAR MARCA]" Then
        
        IDMarca = 0
        IDMarca = frmAgregaMarca.Mostrar()
        If IDMarca > 0 Then
            
            cmbMarca.Clear
            cmbMarca.AddItem "[1. AGREGAR MARCA]"
            Cargar_Combos "Descripcion", "marcas", cmbMarca, "", "Descripcion", False
            cmbMarca.ListIndex = ComboInformacion(cmbMarca, IDMarca)
        
        Else
            
            cmbMarca.ListIndex = -1
        End If
    
    End If
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

Private Sub cmbTipo_GotFocus()
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipo_LostFocus()
    Cambiar_Color False, cmbTipo
End Sub

Private Sub cmdAceptar_Click()
Dim Bandera As Boolean, crMinimo As Double, crMaximo As Double
            
    If cmbTipo.ListIndex = -1 Then
        
        MsgBox "Seleccione el tipo !!", vbInformation, "Prendas Varios"
        cmbTipo.SetFocus
        Exit Sub
    End If
    
    If cmbMarca.ListIndex = -1 Then
        
        MsgBox "Seleccione la marca !!", vbInformation, "Prendas Varios"
        cmbMarca.SetFocus
        Exit Sub
    End If
    
    If cmbFamilia.ListIndex = -1 Then
        
        MsgBox "Seleccione la familia !!", vbInformation, "Prendas Varios"
        cmbFamilia.SetFocus
        Exit Sub
    End If
    
    Bandera = False
    
    If Val(txtMinimo.text) > 0 Or Trim(txtMinimo.text) <> "" Then
        
        crMinimo = CDbl(txtMinimo.text)
    Else
        
        crMinimo = 0
    End If
    
    If Val(txtMaximo.text) > 0 Or Trim(txtMaximo.text) <> "" Then
        
        crMaximo = CDbl(txtMaximo.text)
    Else
        
        crMaximo = 0
    End If
    
    If Val(txtFunciones.Tag) = 0 Then
        
        dbDatos.Execute "INSERT INTO prendaselec (IDTipo,IDMarca,IDFamilia,Modelo,Minimo,Maximo,Funciones,Caracteristicas) VALUES (" & _
                        cmbTipo.ItemData(cmbTipo.ListIndex) & "," & cmbMarca.ItemData(cmbMarca.ListIndex) & "," & cmbFamilia.ItemData(cmbFamilia.ListIndex) & ",'" & Trim(txtModelo.text) & "'," & crMinimo & "," & crMaximo & ",'" & Trim(txtFunciones.text) & "','" & Trim(txtCaracteristicas.text) & "')"
        
        Bandera = True
        
    ElseIf Val(txtFunciones.Tag) > 0 Then
        
        dbDatos.Execute "UPDATE prendaselec SET IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex) & ",IDMarca=" & cmbMarca.ItemData(cmbMarca.ListIndex) & ",IDFamilia=" & cmbFamilia.ItemData(cmbFamilia.ListIndex) & ",Modelo='" & Trim(txtModelo.text) & "',Minimo= " & crMinimo & ",Maximo=" & crMaximo & ",Funciones='" & Trim(txtFunciones.text) & "',Caracteristicas='" & Trim(txtCaracteristicas.text) & "' WHERE ID=" & Val(txtFunciones.Tag)
        
        Bandera = True
        
    End If
    
    If Bandera Then
        CargarPrendas
        cmbMarca.ListIndex = -1
        cmbFamilia.ListIndex = -1
        txtModelo.text = ""
        txtMinimo.text = ""
        txtMaximo.text = ""
        txtFunciones.text = ""
        txtFunciones.Tag = ""
        txtCaracteristicas.text = ""
    End If
    
End Sub


Private Sub cmdEliminar_Click()

    If grdCatPrendas.Rows > 0 Then
        
        If grdCatPrendas.SelectedRow > 0 Then
            
            If MsgBox("Desea eliminar la prenda seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton2, "Prendas Varios") = vbYes Then
                
                dbDatos.Execute "DELETE FROM prendaselec WHERE ID=" & grdCatPrendas.CellItemData(grdCatPrendas.SelectedRow, 4)
                CargarPrendas
                cmbMarca.ListIndex = -1
                cmbFamilia.ListIndex = -1
                txtModelo.text = ""
                txtFunciones.text = ""
                txtFunciones.Tag = ""
                txtCaracteristicas.text = ""
                cmbMarca.SetFocus
            End If

        Else
            
            cmbMarca.SetFocus
        End If

    Else
        
        cmbMarca.SetFocus
    End If

End Sub

Private Sub cmdCancelar_Click()
    cmbMarca.ListIndex = -1
    cmbFamilia.ListIndex = -1
    txtModelo.text = ""
    txtMinimo.text = ""
    txtMaximo.text = ""
    txtFunciones.text = ""
    txtFunciones.Tag = ""
    txtCaracteristicas.text = ""
    grdCatPrendas.ClearSelection
    cmbTipo.SetFocus
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
        
'''''        'Modelo
'''''        If cmbModelo.ListIndex > 0 Then
'''''
'''''            strFiltro = strFiltro & " AND {prendaselec.Modelo}='" & cmbModelo.text & "'"
'''''        End If
        
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

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Cargar_Combos "Descripcion", "tipo", cmbTipo, " WHERE tipo.Kilataje=0 AND Peso=0", "Ordenamiento"
    cmbMarca.AddItem "[1. AGREGAR MARCA]"
    Cargar_Combos "Descripcion", "marcas", cmbMarca, "", "Descripcion", False
    cmbFamilia.AddItem "[1. AGREGAR FAMILIA]"
    Cargar_Combos "tipoprenda.Descripcion", "tipoprenda INNER JOIN tipo ON tipoprenda.IDTipo=tipo.ID", cmbFamilia, " WHERE tipo.Kilataje=0 AND tipo.Peso=0", , False, "tipoprenda.ID"
    Crear_Encabezado
    CargarPrendas
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Sub Crear_Encabezado()

    With grdCatPrendas
    
        .AddColumn "C1", "TIPO", ecgHdrTextALignLeft, , 110, , , , , , , CCLSortString
        .AddColumn "C2", "MARCA", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "C3", "FAMILIA", ecgHdrTextALignLeft, , 190, , , , , , , CCLSortString
        .AddColumn "C4", "MODELO", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
        .AddColumn "C5", "MÍNIMO", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C6", "MÁXIMO", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C7", "FUNCIONES", ecgHdrTextALignLeft, , 150, , , , , , , CCLSortString
        .AddColumn "C8", "CARACTERÍSTICAS", ecgHdrTextALignLeft, , 250, , , , , , , CCLSortString
    
    End With

End Sub

Sub CargarPrendas()
Dim rcTmp As New ADODB.Recordset

On Error GoTo error

    rcTmp.Open "SELECT prendaselec.ID,prendaselec.IDTipo,prendaselec.IDMarca,prendaselec.IDFamilia,tipo.Descripcion AS Desc_Tipo,tipoprenda.Descripcion AS Desc_Familia,prendaselec.Modelo,marcas.Descripcion AS Desc_Marca,prendaselec.Minimo,prendaselec.Maximo,prendaselec.Funciones,prendaselec.Caracteristicas " _
                & "FROM prendaselec INNER JOIN tipo ON prendaselec.IDTipo=tipo.ID LEFT JOIN marcas ON prendaselec.IDMarca=marcas.ID LEFT JOIN tipoprenda ON prendaselec.IDFamilia=tipoprenda.ID ORDER BY prendaselec.IDTipo,tipo.Descripcion,marcas.Descripcion", dbDatos, adOpenForwardOnly, adLockReadOnly
    
    With grdCatPrendas
        
        .Redraw = False
        .Clear
        While Not rcTmp.EOF
            .AddRow
            .CellText(.Rows, 1) = rcTmp!Desc_Tipo
            .CellItemData(.Rows, 1) = rcTmp!IDTipo
            .CellTextAlign(.Rows, 1) = DT_LEFT
            
            .CellText(.Rows, 2) = rcTmp!Desc_Marca
            .CellItemData(.Rows, 2) = rcTmp!IDMarca
            .CellTextAlign(.Rows, 2) = DT_LEFT
            
            .CellText(.Rows, 3) = rcTmp!Desc_Familia
            .CellItemData(.Rows, 3) = rcTmp!IDFamilia
            .CellTextAlign(.Rows, 3) = DT_LEFT
            
            .CellText(.Rows, 4) = rcTmp!Modelo
            .CellItemData(.Rows, 4) = rcTmp!ID
            .CellTextAlign(.Rows, 4) = DT_LEFT
            
            .CellText(.Rows, 5) = rcTmp!Minimo
            .CellTextAlign(.Rows, 5) = DT_RIGHT
            
            .CellText(.Rows, 6) = rcTmp!Maximo
            .CellTextAlign(.Rows, 6) = DT_RIGHT
            
            .CellText(.Rows, 7) = IIf(IsNull(rcTmp!Funciones), "", rcTmp!Funciones)
            .CellTextAlign(.Rows, 7) = DT_LEFT
            
            .CellText(.Rows, 8) = IIf(IsNull(rcTmp!Caracteristicas), "", rcTmp!Caracteristicas)
            .CellTextAlign(.Rows, 8) = DT_LEFT
            
        rcTmp.MoveNext
        Wend
        
        .Redraw = True
    End With
    rcTmp.Close
    Set rcTmp = Nothing
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Sub

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
    
    With grdCatPrendas
        
        If .Rows > 0 Then
            
            If .SelectedRow > 0 Then
                
                cmbTipo.ListIndex = ComboInformacion(cmbTipo, .CellItemData(lRow, 1))
                cmbMarca.ListIndex = ComboInformacion(cmbMarca, .CellItemData(lRow, 2))
                cmbFamilia.ListIndex = ComboInformacion(cmbFamilia, .CellItemData(lRow, 3))
                         
                txtModelo.text = IIf(IsNull(.CellText(lRow, 4)), "", .CellText(lRow, 4))
                txtFunciones.text = .CellText(lRow, 7)
                txtFunciones.Tag = .CellItemData(lRow, 4)
                txtCaracteristicas.text = .CellText(lRow, 8)
                txtMinimo.text = Format(.CellText(lRow, 5), FMoneda)
                txtMaximo.text = Format(.CellText(lRow, 6), FMoneda)
                
                grdCatPrendas.ClearSelection
                txtFunciones.SetFocus
            End If
        
        End If
        
    End With

End Sub

Private Sub txtCaracteristicas_GotFocus()
    Seleccionar_Texto txtCaracteristicas
    Cambiar_Color True, txtCaracteristicas
End Sub

Private Sub txtCaracteristicas_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCaracteristicas_LostFocus()
    Cambiar_Color False, txtCaracteristicas
End Sub

Private Sub txtFunciones_GotFocus()
    Seleccionar_Texto txtFunciones
    Cambiar_Color True, txtFunciones
End Sub

Private Sub txtFunciones_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFunciones_LostFocus()
    Cambiar_Color False, txtFunciones
End Sub

Private Sub txtMaximo_GotFocus()
    Seleccionar_Texto txtMaximo
    Cambiar_Color True, txtMaximo
End Sub

Private Sub txtMaximo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMaximo_LostFocus()
    txtMaximo.text = Format(txtMaximo.text, FMoneda)
    Cambiar_Color False, txtMaximo
End Sub

Private Sub txtMinimo_GotFocus()
    Seleccionar_Texto txtMinimo
    Cambiar_Color True, txtMinimo
End Sub

Private Sub txtMinimo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMinimo_LostFocus()
    txtMinimo.text = Format(txtMinimo.text, FMoneda)
    Cambiar_Color False, txtMinimo
End Sub

Private Sub txtModelo_GotFocus()
    Seleccionar_Texto txtModelo
    Cambiar_Color True, txtModelo
End Sub

Private Sub txtModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtModelo_LostFocus()
    Cambiar_Color False, txtModelo
End Sub
