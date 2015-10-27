VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistorico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte histórico"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHistorico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   14550
   Begin VB.OptionButton opAlmoneda 
      Appearance      =   0  'Flat
      Caption         =   "Almoneda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5640
      TabIndex        =   28
      Top             =   1800
      Width           =   1725
   End
   Begin VB.OptionButton opTodos 
      Appearance      =   0  'Flat
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7440
      TabIndex        =   19
      Top             =   1800
      Width           =   1095
   End
   Begin vbAcceleratorGrid6.vbalGrid grdHistorico 
      Height          =   5775
      Left            =   0
      TabIndex        =   21
      Top             =   2280
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   10186
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
      DisableIcons    =   -1  'True
      DefaultRowHeight=   17
   End
   Begin VB.OptionButton opPagados 
      Appearance      =   0  'Flat
      Caption         =   "Pagados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2280
      TabIndex        =   17
      Top             =   1800
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton opNoPagados 
      Appearance      =   0  'Flat
      Caption         =   "No pagados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3840
      TabIndex        =   18
      Top             =   1800
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenado por"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5640
      TabIndex        =   12
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton opFechaMovimiento 
         Appearance      =   0  'Flat
         Caption         =   "Fecha Movimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Top             =   960
         Width           =   2655
      End
      Begin VB.OptionButton opPrestamo 
         Appearance      =   0  'Flat
         Caption         =   "Préstamo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton opNombre 
         Appearance      =   0  'Flat
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3000
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton opFecha 
         Appearance      =   0  'Flat
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton opFolio 
         Appearance      =   0  'Flat
         Caption         =   "Contrato"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   480
         Width           =   1365
      End
   End
   Begin VB.TextBox txtImporte 
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
      Left            =   3000
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtFolioIni 
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
      Left            =   1200
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtFolioFin 
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
      Left            =   3000
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtDesde 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   255
      Width           =   1215
   End
   Begin VB.TextBox txtHasta 
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
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   375
      Left            =   8940
      TabIndex        =   20
      Top             =   1785
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Buscar"
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
      Picture         =   "frmHistorico.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   1
      Left            =   4965
      TabIndex        =   5
      Top             =   255
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      AlignCaption    =   4
      AlignPicture    =   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmHistorico.frx":0391
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   0
      Left            =   2445
      TabIndex        =   2
      Top             =   255
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      AlignCaption    =   4
      AlignPicture    =   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmHistorico.frx":04A6
   End
   Begin vbalIml6.vbalImageList lstIcons 
      Left            =   4800
      Top             =   1080
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   5740
      Images          =   "frmHistorico.frx":05BB
      Version         =   131072
      KeyCount        =   5
      Keys            =   "ÿÿÿÿ"
   End
   Begin MSComctlLib.ProgressBar BarraProgreso 
      Height          =   195
      Left            =   10560
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   12360
      TabIndex        =   25
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   " &Imprimir"
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
      Picture         =   "frmHistorico.frx":1C47
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   11265
      TabIndex        =   26
      Top             =   1800
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
      Picture         =   "frmHistorico.frx":2199
   End
   Begin DevPowerFlatBttn.FlatBttn cmdPerdida 
      Height          =   375
      Left            =   10080
      TabIndex        =   27
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Perdido"
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
      Picture         =   "frmHistorico.frx":26EB
      PictureDisabled =   "frmHistorico.frx":2C3D
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de contratos:"
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
      TabIndex        =   16
      Top             =   1800
      Width           =   2010
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Contratos mayores a:"
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
      TabIndex        =   10
      Top             =   1200
      Width           =   2385
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "A"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   720
      Width           =   150
   End
   Begin VB.Label Label3 
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
      TabIndex        =   6
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Desde:"
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
      TabIndex        =   0
      Top             =   240
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hasta:"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 24/04/2002
' Modulo frmHistorico - frmHistorico.frm
' Ultima Modificacion - 14/05/2002
''Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdImprimir_Click()
    If grdHistorico.Rows > 0 Then Exportar_Excel grdHistorico
End Sub

Private Sub cmdPerdida_Click()

    If grdHistorico.SelectedRow > 0 And grdHistorico.SelectedRow < grdHistorico.Rows Then
                                                        
        If grdHistorico.CellItemData(grdHistorico.SelectedRow, 16) = 0 And grdHistorico.CellItemData(grdHistorico.SelectedRow, 6) = 0 And grdHistorico.CellItemData(grdHistorico.SelectedRow, 2) = 0 Then
            
            If MsgBox("Desea marcar el contrato seleccionado como perdido ??", vbQuestion + vbYesNo + vbDefaultButton2, "Reporte histórico") = vbYes Then
                
                dbDatos.Execute "UPDATE empeno SET Perdida=1 WHERE ID=" & Val(grdHistorico.CellItemData(grdHistorico.SelectedRow, 1))
                grdHistorico.CellItemData(grdHistorico.SelectedRow, 16) = 1
                Colorea grdHistorico, grdHistorico.SelectedRow, RGB(244, 119, 66)
                grdHistorico.ClearSelection
            
            End If
            
        ElseIf grdHistorico.CellItemData(grdHistorico.SelectedRow, 16) = 1 And grdHistorico.CellItemData(grdHistorico.SelectedRow, 6) = 0 And grdHistorico.CellItemData(grdHistorico.SelectedRow, 2) = 0 Then
            
            If MsgBox("Desea eliminar la marca de perdido al contrato seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Reporte histórico") = vbYes Then
                
                dbDatos.Execute "UPDATE empeno SET Perdida=0 WHERE ID=" & Val(grdHistorico.CellItemData(grdHistorico.SelectedRow, 1))
                grdHistorico.CellItemData(grdHistorico.SelectedRow, 16) = 0
                Poner_Colores grdHistorico, grdHistorico.SelectedRow, grdHistorico.CellItemData(grdHistorico.SelectedRow, 18)
                grdHistorico.ClearSelection
            
            End If
            
        Else
            
            grdHistorico.ClearSelection
        End If
    
    Else
        
        MsgBox "Seleccione el contrato que desea marcar como perdido !!", vbInformation, "Reporte histórico"
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdBuscar_Click()
    Buscar_Historico
End Sub

Private Sub cmdMosFecha_Click(Index As Integer)
    
    If Index = 0 Then
        
        txtDesde.text = frmCalendario.Fecha(txtDesde.text)
    Else
        
        txtHasta.text = frmCalendario.Fecha(txtHasta.text)
    End If
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

'Inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
    Crear_Encabezados
    Screen.MousePointer = vbDefault
End Sub

'Creamos los encabezados del grid
Private Sub Crear_Encabezados()
   
    With grdHistorico
        
        .ImageList = lstIcons
        .AddColumn "K1", "Fecha", ecgHdrTextALignCentre, , 75, , , , , "DD/MM/YY", , CCLSortDate
        .AddColumn "K2", "Contrato", ecgHdrTextALignRight, , 68, , , , , , , CCLSortNumeric
        .AddColumn "K3", "Folio", ecgHdrTextALignRight, , 68, False, , , , , , CCLSortNumeric
        .AddColumn "K4", "Cliente", ecgHdrTextALignLeft, , 260, , , , , , , CCLSortString
        .AddColumn "K5", "Inic.", ecgHdrTextALignLeft, , 55, False, , , , , , CCLSortString
        .AddColumn "K6", "Préstamo", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Avalúo", ecgHdrTextALignRight, , 82, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K8", "Vence", ecgHdrTextALignCentre, , 71, , , , , "DD/MM/YY", , CCLSortDate
        .AddColumn "K9", "Origen/Folio", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "K10", "Destino/Folio", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "K11", "Pago", ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K12", "Intereses", ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K13", "Otros Cobros", ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K14", "Iva", ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K15", "Fec. Mov.", ecgHdrTextALignCentre, , 80, , , , , "DD/MM/YY", , CCLSortDate
        .AddColumn "K16", "Tasa", ecgHdrTextALignLeft, , 170, , , , , , , CCLSortString
        .AddColumn "K17", "Num. Bolsa", ecgHdrTextALignLeft, , 160, , , , , , , CCLSortString
        .AddColumn "K18", "Valuador", ecgHdrTextALignLeft, , 160, , , , , , , CCLSortString
        .AddColumn "K19", "Ubicación", ecgHdrTextALignLeft, , 160, , , , , , , CCLSortString
   
   End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtDesde_GotFocus()
    Seleccionar_Texto txtDesde
    Cambiar_Color True, txtDesde
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDesde_LostFocus()
    Cambiar_Color False, txtDesde
End Sub

Private Sub txtFolioFin_GotFocus()
    Seleccionar_Texto txtFolioFin
    Cambiar_Color True, txtFolioFin
End Sub

Private Sub txtFolioFin_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub txtFolioFin_LostFocus()
    Cambiar_Color False, txtFolioFin
End Sub

Private Sub txtFolioIni_GotFocus()
    Seleccionar_Texto txtFolioIni
    Cambiar_Color True, txtFolioIni
End Sub

Private Sub txtFolioIni_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub txtFolioIni_LostFocus()
    Cambiar_Color False, txtFolioIni
End Sub

Private Sub txtHasta_GotFocus()
    Seleccionar_Texto txtHasta
    Cambiar_Color True, txtHasta
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtHasta_LostFocus()
    Cambiar_Color False, txtHasta
End Sub

Private Sub txtImporte_GotFocus()
    Seleccionar_Texto txtImporte
    Cambiar_Color True, txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub txtImporte_LostFocus()
    Cambiar_Color False, txtImporte
End Sub

'Buscamos todas las coincidencias
Private Sub Buscar_Historico()
Dim rcEmpeño As New ADODB.Recordset
Dim strBusqueda As String, i As Long
Dim strCondicion As String

On Error GoTo Error
    
    strCondicion = strCriterio
    BarraProgreso.Value = 0
    BarraProgreso.Min = 0
    rcEmpeño.Open "SELECT COUNT(ID) AS Total FROM empeno " & strCondicion, dbDatos, adOpenForwardOnly, adLockReadOnly
    If rcEmpeño!Total > 0 Then
        BarraProgreso.Max = rcEmpeño!Total
    End If
    rcEmpeño.Close
    BarraProgreso.Visible = True
    
    strBusqueda = "SELECT DISTINCT empeno.ID,empeno.Cancelado,empeno.Fecha,empeno.NumContrato,empeno.Folio,empeno.Prestamo,empeno.PrestamoInicial,empeno.Serie,empeno.Vencimiento,empeno.Origen,empeno.FolioOrigen,empeno.Destino,empeno.Pago,empeno.Pagado,empeno.FechaMovimiento,empeno.TipoInteres,empeno.TipoTasa,empeno.FolioDestino,empeno.Cancelado,(empeno.Intereses+empeno.ImporteAlmacenaje+empeno.ImporteSeguro+empeno.ImporteMoratorios) AS Interes,empeno.ImportePerdida,empeno.ImporteIva,empeno.Perdida,empeno.Avaluo,empeno.NumBolsa,empeno.Caja,empeno.Cajon,empeno.Fila,empeno.Valuador,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente,clientes.Iniciales,empeno.Periodo FROM empeno LEFT JOIN clientes ON empeno.IDCliente=clientes.ID"
    rcEmpeño.Open strBusqueda & strCondicion & strOrder, dbDatos, adOpenForwardOnly, adLockReadOnly
       
    With rcEmpeño
        
        grdHistorico.Redraw = False
        grdHistorico.Clear
        While Not .EOF
            i = i + 1
            DoEvents
            grdHistorico.AddRow
            grdHistorico.CellText(grdHistorico.Rows, 1) = !Fecha
            grdHistorico.CellIcon(grdHistorico.Rows, 1) = 3
            grdHistorico.CellItemData(grdHistorico.Rows, 1) = !ID
            grdHistorico.CellTextAlign(grdHistorico.Rows, 1) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdHistorico.CellText(grdHistorico.Rows, 2) = !NumContrato
            grdHistorico.CellItemData(grdHistorico.Rows, 2) = !Cancelado
            grdHistorico.CellTextAlign(grdHistorico.Rows, 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdHistorico.CellText(grdHistorico.Rows, 3) = !Folio
            grdHistorico.CellTextAlign(grdHistorico.Rows, 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdHistorico.CellText(grdHistorico.Rows, 4) = !Cliente
            grdHistorico.CellTextAlign(grdHistorico.Rows, 4) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdHistorico.CellText(grdHistorico.Rows, 5) = !Iniciales
            grdHistorico.CellTextAlign(grdHistorico.Rows, 5) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdHistorico.CellText(grdHistorico.Rows, 6) = IIf(!Destino = D_DESEMPEÑO And !TipoInteres = "FIJA", !PrestamoInicial, !Prestamo)
            grdHistorico.CellItemData(grdHistorico.Rows, 6) = !Pagado
            grdHistorico.CellTextAlign(grdHistorico.Rows, 6) = DT_RIGHT
            grdHistorico.CellText(grdHistorico.Rows, 7) = !Avaluo
            grdHistorico.CellTextAlign(grdHistorico.Rows, 7) = DT_RIGHT
            grdHistorico.CellText(grdHistorico.Rows, 8) = !Vencimiento
            grdHistorico.CellTextAlign(grdHistorico.Rows, 8) = DT_CENTER Or DT_WORD_ELLIPSIS
            'Origen
            grdHistorico.CellText(grdHistorico.Rows, 9) = OD_Origen(!Origen) & IIf(!Cancelado = 1 And !Destino = 0, "/Cancelado", "/" & !FolioOrigen)
            grdHistorico.CellTextAlign(grdHistorico.Rows, 9) = DT_LEFT Or DT_WORD_ELLIPSIS
            'Destino
            grdHistorico.CellText(grdHistorico.Rows, 10) = OD_Origen(!Destino) & IIf(!Cancelado = 1 And !Destino > 0, "/Cancelado", IIf(Val(!Destino) = D_VENTA Or Val(!Destino) = OD_REFRENDO, "/" & !FolioDestino, ""))
            grdHistorico.CellTextAlign(grdHistorico.Rows, 10) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdHistorico.CellText(grdHistorico.Rows, 11) = !Pago
            grdHistorico.CellTextAlign(grdHistorico.Rows, 11) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdHistorico.CellText(grdHistorico.Rows, 12) = !Interes
            grdHistorico.CellTextAlign(grdHistorico.Rows, 12) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdHistorico.CellText(grdHistorico.Rows, 13) = !ImportePerdida
            grdHistorico.CellTextAlign(grdHistorico.Rows, 13) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdHistorico.CellText(grdHistorico.Rows, 14) = !ImporteIva
            grdHistorico.CellTextAlign(grdHistorico.Rows, 14) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdHistorico.CellText(grdHistorico.Rows, 15) = !FechaMovimiento
            grdHistorico.CellTextAlign(grdHistorico.Rows, 15) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdHistorico.CellText(grdHistorico.Rows, 16) = !TipoInteres & "-" & !TipoTasa & " " & !Periodo & " Dias"
            grdHistorico.CellItemData(grdHistorico.Rows, 16) = !Perdida
            grdHistorico.CellTextAlign(grdHistorico.Rows, 16) = DT_LEFT
            grdHistorico.CellText(grdHistorico.Rows, 17) = IIf(IsNull(!NumBolsa), "", !NumBolsa)
            grdHistorico.CellText(grdHistorico.Rows, 18) = !Valuador
            grdHistorico.CellItemData(grdHistorico.Rows, 18) = i
            grdHistorico.CellText(grdHistorico.Rows, 19) = IIf(IsNull(!caja) Or Trim(!caja) = "", "", "CAJA:" & !caja) & " " & IIf(IsNull(!Cajon) Or Trim((!Cajon)) = "", "", " CAJÓN:" & !Cajon) & " " & IIf(IsNull(!Fila) Or Trim(!Fila) = "", "", " FILA:" & !Fila)
            
            If !Perdida = 1 Then
            
                Colorea grdHistorico, grdHistorico.Rows, RGB(244, 119, 66)
            Else
                
                Poner_Colores grdHistorico, grdHistorico.Rows, i
            End If
            
            DetalleEmpenos !ID, !Serie
            BarraProgreso.Value = i
            
        .MoveNext
        Wend
        
    End With
    rcEmpeño.Close
    Set rcEmpeño = Nothing
    
    BarraProgreso.Visible = False
    
    If grdHistorico.Rows > 0 Then
        grdHistorico.AddRow
        Poner_Totales
    End If
    
    grdHistorico.Redraw = True
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcEmpeño = Nothing
End Sub

Private Function strCriterio() As String
Dim Criterio As String
   
    strCriterio = ""
   
    'el criterio de las fechas
    If Trim(txtDesde.text) <> "" And Trim(txtHasta.text) = "" Then Criterio = " (DATE_FORMAT(Fecha,'%Y%/%m%/%d') >='" & Format(txtDesde.text, "YYYY/MM/DD") & "'" & IIf(opPagados.Value Or opTodos.Value, " OR DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d')>='" & Format(txtDesde.text, "YYYY/MM/DD") & "')", ")")
    If Trim(txtDesde.text) = "" And Trim(txtHasta.text) <> "" Then Criterio = " (DATE_FORMAT(Fecha,'%Y%/%m%/%d') <='" & Format(txtHasta.text, "YYYY/MM/DD") & "'" & IIf(opPagados.Value Or opTodos.Value, " OR DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d')<='" & Format(txtHasta.text, "YYYY/MM/DD") & "')", ")")
    If Trim(txtDesde.text) <> "" And Trim(txtHasta.text) <> "" Then Criterio = " (DATE_FORMAT(Fecha,'%Y%/%m%/%d') BETWEEN '" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND '" & Format(txtHasta.text, "YYYY/MM/DD") & "')" & IIf(opPagados.Value Or opTodos.Value, " OR (DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d') BETWEEN '" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND '" & Format(txtHasta.text, "YYYY/MM/DD") & "')", "")
   
    If Trim(txtFolioIni.text) <> "" Or Trim(txtFolioFin.text) <> "" Then If Criterio <> "" Then Criterio = Criterio & " AND "
    If Trim(txtFolioIni.text) <> "" And Trim(txtFolioFin.text) = "" Then Criterio = Criterio & " NumContrato >= " & txtFolioIni.text
    If Trim(txtFolioIni.text) = "" And Trim(txtFolioFin.text) <> "" Then Criterio = Criterio & " NumContrato <= " & txtFolioFin.text
    If Trim(txtFolioIni.text) <> "" And Trim(txtFolioFin.text) <> "" Then Criterio = Criterio & " NumContrato >= " & txtFolioIni.text & " AND NumContrato <= " & txtFolioFin.text
      
    If Trim(txtImporte.text) <> "" Then
        
        If Criterio <> "" Then Criterio = Criterio & " AND "
        Criterio = Criterio & " Prestamo > " & txtImporte.text
    End If
      
    If Criterio <> "" And Not opTodos.Value Then Criterio = Criterio & " AND "
    If opPagados.Value Then
        Criterio = Criterio & " Pagado=1"
    ElseIf opNoPagados.Value Then
        Criterio = Criterio & " Pagado=0"
    ElseIf opAlmoneda.Value Then
        Criterio = Criterio & " Almoneda=1"
    ElseIf opTodos.Value Then
        Criterio = Criterio & ""
    End If
   
    strCriterio = IIf(Criterio <> "", " WHERE " & Criterio, "")
End Function

Private Function strOrder() As String
Dim Cadena As String

'''''    If opPagados.Value Then
'''''
'''''        Cadena = "Pagado"
'''''    ElseIf opNoPagados.Value Then
'''''
'''''        Cadena = "Pagado"
'''''    End If
   
    If opFecha.Value Then
      
        Cadena = "Fecha,NumContrato,Folio"
   
    ElseIf opFolio.Value Then
      
        Cadena = "NumContrato,Folio"

    ElseIf opNombre.Value Then
      
        Cadena = "Concat(clientes.Nombre,' ',clientes.Apellido)"
      
    ElseIf opPrestamo.Value Then
      
        Cadena = "Prestamo,NumContrato,Folio"
   
    ElseIf opFechaMovimiento.Value Then
      
        Cadena = "FechaMovimiento,NumContrato,Folio"
    
    End If
   
    strOrder = " ORDER BY " & Cadena
   
End Function

Private Sub Poner_Totales()
Dim crPrestamo As Double, crAvaluo As Double, crPago As Double, crIntereses As Double, crOtrosCobros As Double, crIva As Double
Dim Renglon As Long, columna As Integer, Total As Long
    
On Error GoTo Error
    
    'Hago la sumatoria de los totales (Prestamo, Pago, Saldos) desde el renglon 1 hasta el numero de renglones del GRID
    Total = 0
    crPrestamo = 0
    crAvaluo = 0
    crPago = 0
    crIntereses = 0
    crOtrosCobros = 0
    crIva = 0
    
    For Renglon = 1 To grdHistorico.Rows - 1

'        If Val(grdHistorico.CellItemData(Renglon, 4)) = 0 Then
'
'            crPrestamo = crPrestamo + IIf(Val(grdHistorico.CellItemData(Renglon, 2)) = 1, 0, CDbl(grdHistorico.CellText(Renglon, 6)))
'            crAvaluo = crAvaluo + CDbl(grdHistorico.CellText(Renglon, 7))
'            crPago = crPago + CDbl(grdHistorico.CellText(Renglon, 11))
'            crIntereses = crIntereses + CDbl(IIf(IsNull(grdHistorico.CellText(Renglon, 12)), 0, grdHistorico.CellText(Renglon, 12)))
'            crOtrosCobros = crOtrosCobros + CDbl(IIf(IsNull(grdHistorico.CellText(Renglon, 13)), 0, grdHistorico.CellText(Renglon, 13)))
'            crIva = crIva + CDbl(IIf(IsNull(grdHistorico.CellText(Renglon, 14)), 0, grdHistorico.CellText(Renglon, 14)))
'            Total = Total + IIf(grdHistorico.RowVisible(Renglon) And grdHistorico.CellItemData(Renglon, 2) = 0, 1, 0)
'
'        End If
        If Val(grdHistorico.CellItemData(Renglon, 4)) = 0 Then
        
            crPrestamo = crPrestamo + IIf(Val(grdHistorico.CellItemData(Renglon, 2)) = 1, 0, CDbl(grdHistorico.CellText(Renglon, 6)))
            crAvaluo = crAvaluo + IIf(Val(grdHistorico.CellItemData(Renglon, 2)) = 1, 0, CDbl(grdHistorico.CellText(Renglon, 7)))
            crPago = crPago + IIf(Val(grdHistorico.CellItemData(Renglon, 2)) = 1, 0, CDbl(grdHistorico.CellText(Renglon, 11)))
            crIntereses = crIntereses + IIf(Val(grdHistorico.CellItemData(Renglon, 2)) = 1, 0, CDbl(IIf(IsNull(grdHistorico.CellText(Renglon, 12)), 0, grdHistorico.CellText(Renglon, 12))))
            crOtrosCobros = crOtrosCobros + IIf(Val(grdHistorico.CellItemData(Renglon, 2)) = 1, 0, CDbl(IIf(IsNull(grdHistorico.CellText(Renglon, 13)), 0, grdHistorico.CellText(Renglon, 13))))
            crIva = crIva + IIf(Val(grdHistorico.CellItemData(Renglon, 2)) = 1, 0, CDbl(IIf(IsNull(grdHistorico.CellText(Renglon, 14)), 0, grdHistorico.CellText(Renglon, 14))))
            Total = Total + IIf(grdHistorico.RowVisible(Renglon) And grdHistorico.CellItemData(Renglon, 2) = 0, 1, 0)
        End If

    Next Renglon
            
    'En la ultima linea del GRID cargo los totales (Prestamo, Pago, Saldos) y cambio el color de la linea
    grdHistorico.CellText(grdHistorico.Rows, 2) = "No. " & Total
    grdHistorico.CellTextAlign(grdHistorico.Rows, 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdHistorico.CellText(grdHistorico.Rows, 6) = crPrestamo
    grdHistorico.CellTextAlign(grdHistorico.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdHistorico.CellText(grdHistorico.Rows, 7) = crAvaluo
    grdHistorico.CellTextAlign(grdHistorico.Rows, 7) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdHistorico.CellText(grdHistorico.Rows, 11) = crPago
    grdHistorico.CellTextAlign(grdHistorico.Rows, 11) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdHistorico.CellText(grdHistorico.Rows, 12) = crIntereses
    grdHistorico.CellTextAlign(grdHistorico.Rows, 12) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdHistorico.CellText(grdHistorico.Rows, 13) = crOtrosCobros
    grdHistorico.CellTextAlign(grdHistorico.Rows, 13) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdHistorico.CellText(grdHistorico.Rows, 14) = crIva
    grdHistorico.CellTextAlign(grdHistorico.Rows, 14) = DT_RIGHT Or DT_WORD_ELLIPSIS
                   
    For columna = 1 To grdHistorico.Columns
        
        grdHistorico.CellBackColor(grdHistorico.Rows, columna) = RGB(223, 208, 102)
    Next columna
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Sub DetalleEmpenos(ID As Long, Serie As Integer)
Dim rcConsulta As New ADODB.Recordset
Dim strPrenda As String, strSql As String

On Error GoTo Error
    
    If Serie = SERIE_B Then
        
        strSql = "SELECT d.IDEmpeno,d.MarcayModelo,d.Placas,d.Año,d.Color,d.SerieChasis,d.NumMotor,d.NumTarjetaCircu,empeno.Prestamo,empeno.Avaluo " _
                & "FROM detallesempenoautos d INNER JOIN empeno ON d.IDEmpeno=empeno.ID WHERE d.IDEmpeno=" & ID

    Else
        
        strSql = "SELECT d.IDEmpeno,d.Tipo,d.Cantidad,d.Articulo,d.Observaciones,d.Peso,d.Estado,kilatajes.Descripcion,d.Prestamo,d.Avaluo,d.Destino FROM detallesempeno d LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave WHERE d.IDEmpeno=" & ID
    End If
    
    rcConsulta.Open strSql, dbDatos, adOpenForwardOnly, adLockReadOnly
    While Not rcConsulta.EOF
        
        grdHistorico.AddRow
        If Serie = SERIE_B Then
            
            strPrenda = "MARCA Y MODELO: " & rcConsulta!MarcayModelo & ", PLACAS: " & rcConsulta!Placas & ", AÑO: " & rcConsulta!Año & ", COLOR: " & rcConsulta!Color & ", SERIE CHASIS: " & rcConsulta!SerieChasis & ", NUM. MOTOR: " & rcConsulta!NumMotor & ", TARJETA CIRC.: " & rcConsulta!NumTarjetaCircu
        Else
            
            strPrenda = rcConsulta!Cantidad & " " & rcConsulta!Articulo & " " & rcConsulta!Observaciones & " " & rcConsulta!Descripcion & IIf(rcConsulta!Tipo = 1, " " & Format(rcConsulta!Peso, "###.000") & " Grms. ", "") & IIf(IsNull(rcConsulta!Estado) Or Trim(rcConsulta!Estado) = "", "", " ESTADO: " & rcConsulta!Estado)
            grdHistorico.CellText(grdHistorico.Rows, 10) = OD_Origen(rcConsulta!Destino)
        End If
        
        
        grdHistorico.CellText(grdHistorico.Rows, 4) = strPrenda
        grdHistorico.CellItemData(grdHistorico.Rows, 4) = rcConsulta!IDEmpeno
    
        grdHistorico.CellText(grdHistorico.Rows, 6) = rcConsulta!Prestamo
        grdHistorico.CellTextAlign(grdHistorico.Rows, 6) = DT_RIGHT
    
        grdHistorico.CellText(grdHistorico.Rows, 7) = rcConsulta!Avaluo
        grdHistorico.CellTextAlign(grdHistorico.Rows, 7) = DT_RIGHT
                
        grdHistorico.RowVisible(grdHistorico.Rows) = False
        SombreaGrid grdHistorico, 239, 239, 239, 255, 255, 255, grdHistorico.Rows
    
    rcConsulta.MoveNext
    Wend
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Sub MuestraOculta(ID As Long, Opcion As Boolean)
Dim i As Long

    For i = 1 To grdHistorico.Rows

        If grdHistorico.CellItemData(i, 4) = ID Then
            
            grdHistorico.RowVisible(i) = Opcion
        End If

    Next i

End Sub

Private Sub grdHistorico_Click(ByVal lRow As Long, ByVal lCol As Long)

    If lCol = 0 Or lRow = 0 Then
        
        Exit Sub
    
    ElseIf lCol = 1 And lRow > 0 And grdHistorico.CellIcon(lRow, lCol) = 3 And grdHistorico.CellItemData(lRow, 4) = 0 Then
        
        grdHistorico.CellIcon(lRow, lCol) = 4
        MuestraOculta grdHistorico.CellItemData(lRow, 1), True
    
    ElseIf lCol = 1 And lRow > 0 And grdHistorico.CellItemData(lRow, 4) = 0 Then
        
        grdHistorico.CellIcon(lRow, lCol) = 3
        MuestraOculta grdHistorico.CellItemData(lRow, 1), False
    End If

End Sub

'''''Private Sub Exportar_Excel()
'''''Dim Excel As Object, i As Integer, Col As Integer, y As Integer, Str As String, detalles As Boolean, Pos As Long, Ban As Boolean
'''''
'''''    On Error GoTo Error
'''''
'''''    detalles = False
'''''    Screen.MousePointer = vbHourglass
'''''    DoEvents
'''''
'''''    'Creo la Referencia al Excel
'''''    Set Excel = CreateObject("Excel.application")
'''''
'''''    With Excel
'''''
'''''        'Agrego un Nuevo Libro
'''''        .Workbooks.Add
'''''
'''''        'Creo los Encabezados
'''''        For i = 1 To grdHistorico.Columns
'''''            DoEvents
'''''            .Cells(1, i).Formula = grdHistorico.ColumnHeader("K" & i)
'''''        Next i
'''''
'''''        Pos = 1
'''''
'''''        For i = 1 To grdHistorico.Rows
'''''            DoEvents
'''''
'''''            For y = 1 To grdHistorico.Columns
'''''
'''''                .Cells(Pos + 1, y).Formula = grdHistorico.CellText(i, y)
'''''
'''''            Next y
'''''
'''''            If Ban Then
'''''                Ban = False
'''''            Else
'''''                Pos = Pos + 1
'''''            End If
'''''
'''''        Next i
'''''
'''''        ' autoajustar las columnas
'''''        .Columns("A:A").EntireColumn.AutoFit
'''''        .Columns("B:B").EntireColumn.AutoFit
'''''        .Columns("C:C").EntireColumn.AutoFit
'''''        .Columns("D:D").EntireColumn.AutoFit
'''''        .Columns("E:E").EntireColumn.AutoFit
'''''        .Columns("F:F").EntireColumn.AutoFit
'''''        .Columns("G:G").EntireColumn.AutoFit
'''''        .Columns("H:H").EntireColumn.AutoFit
'''''        .Columns("I:I").EntireColumn.AutoFit
'''''        .Columns("J:J").EntireColumn.AutoFit
'''''        .Columns("K:K").EntireColumn.AutoFit
'''''        .Columns("L:L").EntireColumn.AutoFit
'''''        .Columns("M:M").EntireColumn.AutoFit
'''''
'''''        .Range("A1:M1").Select
'''''        With .Selection.Font
'''''            .Name = "Arial"
'''''            .FontStyle = "Negrita"
'''''            .Size = 10
'''''            .Strikethrough = False
'''''            .Superscript = False
'''''            .Subscript = False
'''''            .OutlineFont = False
'''''            .Shadow = False
'''''        End With
'''''
'''''        Str = "M" & grdHistorico.Rows + 1
'''''        '.ActiveSheet.Range("A1", str).HorizontalAlignment = xlHAlignLeft
'''''        .ActiveSheet.Range("A1", Str).HorizontalAlignment = -4131
'''''        .Selection.Interior.ColorIndex = 35
'''''
'''''        'Hago Visible la Referencia
'''''        .Visible = True
'''''
'''''    End With
'''''    Set Excel = Nothing
'''''    Screen.MousePointer = vbDefault
'''''    Exit Sub
'''''
'''''Error:
'''''    Set Excel = Nothing
'''''    Screen.MousePointer = vbDefault
'''''    Maneja_Error Err
'''''End Sub

Private Sub Exportar_Excel(Grid As vbalGrid)
Dim Excel As Object, i As Integer, Col As Integer, Y As Integer, str As String, Columnas As Integer

On Error GoTo Error

    Screen.MousePointer = vbHourglass
    DoEvents
    
    'Creo la Referencia al Excel
    Set Excel = CreateObject("Excel.application")
    
    With Excel
                
        'Agrego un Nuevo Libro
        .Workbooks.Add
        
        .Range("A1:" & Chr(65 + Grid.Columns) & "1").Select
        With .Selection.Font
            .Name = "Arial"
            .FontStyle = "Negrita"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
        End With
        
        .Range("A1:" & Chr(65 + Grid.Columns) & "1").Select
        .Range("A1", Trim(Chr(65 + Grid.Columns) & "1")).HorizontalAlignment = xlHAlignCenter
        .Selection.Font.Color = RGB(255, 255, 255)
        .Selection.Interior.ColorIndex = 41
                
        'Creo los Encabezados
        For i = 1 To Grid.Columns
            
            .Cells(1, i).Formula = UCase(Grid.ColumnHeader(i))
            .Columns("A:" & i).EntireColumn.AutoFit
        Next i
        
        'Imprimo los contenidos
        For i = 1 To Grid.Rows
        
            For Y = 1 To Grid.Columns
                
                If Grid.RowVisible(i) Then
                    
                    .Cells(i + 1, Y).Formula = Grid.CellText(i, Y)
                
                End If
                
            Next Y
        
        Next i

'''''        Columnas = 1
'''''        For i = 65 To (65 + (Grid.Columns - 1))
'''''            .Columns(Chr(i) & ":" & Chr(i)).ColumnWidth = 15
'''''            Columnas = Columnas + 1
'''''            If Columnas >= Grid.Columns Then Exit For
'''''        Next i
        
        .Range("A2:" & Chr(65 + (Grid.Columns)) & Grid.Rows).Select
        With .Selection.Font
            .Name = "Arial"
            .FontStyle = "Negrita"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
        End With
        
'''''        .Range("A9:" & Chr(65 + (Grid.Columns - 1)) & 9).Select
'''''        .Selection.WrapText = True
'''''        .Selection.Interior.ColorIndex = 41
'''''        .Selection.VerticalAlignment = xlCenter
'''''        .Selection.HorizontalAlignment = xlCenter
'''''        .Range("A9", Chr((65 + (Grid.Columns - 1))) & Grid.Rows + 9).Borders().LineStyle = xlContinuous
        
        
'''''        .Columns() & ":" & Chr(65 + (Me.Servicios + 13))).Select
'''''        With .Selection
'''''            .VerticalAlignment = xlCenter
'''''            .HorizontalAlignment = xlCenter
'''''            .WrapText = True
'''''            .ColumnWidth = 25.14
'''''        End With
        
'''''        .Range("A10" & ":" & Chr(65 + Grid.Columns) & (Grid.Rows + 9)).Select
'''''        .Selection.RowHeight = 36.75
        
        'Hago Visible la Referencia
        .Visible = True

    End With
    Set Excel = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
Error:
    Set Excel = Nothing
    Screen.MousePointer = vbDefault
End Sub
