VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.2#0"; "vbalSGrid6.ocx"
Begin VB.Form frmMovimientosAtipicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Movimientos atípicos"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimientosAtipicos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   13200
   Begin TabDlg.SSTab TabFechas 
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2143
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Mensual"
      TabPicture(0)   =   "frmMovimientosAtipicos.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAño"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmbMesReporte"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdDec"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdInc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtAgnoReporte"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Semestral"
      TabPicture(1)   =   "frmMovimientosAtipicos.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDesde"
      Tab(1).Control(1)=   "txtHasta"
      Tab(1).Control(2)=   "cmdMosFecha(1)"
      Tab(1).Control(3)=   "cmdMosFecha(0)"
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(5)=   "Label1"
      Tab(1).ControlCount=   6
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
         Left            =   -73320
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   495
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
         Left            =   -73320
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtAgnoReporte 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "2014"
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdInc 
         Caption         =   ">"
         Height          =   360
         Left            =   1320
         TabIndex        =   28
         Top             =   720
         Width           =   390
      End
      Begin VB.CommandButton cmdDec 
         Caption         =   "<"
         Height          =   360
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   390
      End
      Begin VB.ComboBox cmbMesReporte 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmMovimientosAtipicos.frx":0044
         Left            =   1800
         List            =   "frmMovimientosAtipicos.frx":006C
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   720
         Width           =   1695
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
         Height          =   300
         Index           =   1
         Left            =   -72075
         TabIndex        =   34
         Top             =   855
         Visible         =   0   'False
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
         Picture         =   "frmMovimientosAtipicos.frx":00D5
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
         Height          =   300
         Index           =   0
         Left            =   -72075
         TabIndex        =   35
         Top             =   495
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
         Picture         =   "frmMovimientosAtipicos.frx":01EA
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74550
         TabIndex        =   37
         Top             =   495
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74550
         TabIndex        =   36
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblAño 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label lblMes 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   30
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1815
      Left            =   8640
      TabIndex        =   21
      Top             =   8160
      Width           =   4455
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   1080
         Left            =   2640
         Picture         =   "frmMovimientosAtipicos.frx":02FF
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdGenerarXML 
         Caption         =   "Generar Aviso"
         Height          =   1080
         Left            =   360
         Picture         =   "frmMovimientosAtipicos.frx":0795
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1590
      End
      Begin DevPowerFlatBttn.FlatBttn FlaGenerarXML 
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         AlignCaption    =   4
         AlignPicture    =   2
         AutoSize        =   0   'False
         Caption         =   "   &Generar Aviso XML"
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
         Picture         =   "frmMovimientosAtipicos.frx":1497
      End
   End
   Begin VB.OptionButton opCheque 
      Appearance      =   0  'Flat
      Caption         =   "Préstamos realizados con cheque"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   8040
      TabIndex        =   16
      Top             =   10830
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   8040
      TabIndex        =   15
      Top             =   10680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.OptionButton opPrestamo 
      Appearance      =   0  'Flat
      Caption         =   "Préstamo mayores a:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9240
      TabIndex        =   13
      Top             =   600
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   9120
      TabIndex        =   11
      Top             =   600
      Width           =   2535
      Begin VB.TextBox txtPrestamo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   600
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   435
         TabIndex        =   14
         Top             =   308
         Width           =   105
      End
   End
   Begin VB.OptionButton opUdis 
      Appearance      =   0  'Flat
      Caption         =   "Préstamos mayores a:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   10
      Top             =   600
      Width           =   2295
   End
   Begin VB.OptionButton opSalarios 
      Appearance      =   0  'Flat
      Caption         =   "Préstamos mayores a:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   9
      Top             =   600
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   6480
      TabIndex        =   6
      Top             =   600
      Width           =   2535
      Begin VB.TextBox txtUdis 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   7
         Text            =   "0"
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UDIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   8
         Top             =   300
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   3840
      TabIndex        =   3
      Top             =   600
      Width           =   2535
      Begin VB.TextBox txtSalario 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   4
         Text            =   "0"
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salarios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   5
         Top             =   308
         Width           =   705
      End
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdClientes 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   5106
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdDetalle 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   4895
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      DisableIcons    =   -1  'True
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   375
      Left            =   11880
      TabIndex        =   1
      Top             =   720
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
      Picture         =   "frmMovimientosAtipicos.frx":31E6
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdAvisos 
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Top             =   8280
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2990
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      DisableIcons    =   -1  'True
   End
   Begin VB.Label lblLabel6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUTA:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   7920
      Width           =   630
   End
   Begin VB.Label lblRutaXML 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARPETA:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   19
      Top             =   7920
      Width           =   12330
   End
   Begin VB.Label lblLosAvisos 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMovimientosAtipicos.frx":356B
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   17
      Top             =   7320
      Width           =   12930
   End
End
Attribute VB_Name = "frmMovimientosAtipicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sFntUnread As New StdFont
Dim rcConsulta As New ADODB.Recordset
Dim rcConsulta2 As New ADODB.Recordset
Dim rcConsulta3 As New ADODB.Recordset
Dim rcConsultaAuto As ADODB.Recordset
Dim rcConsultaCero As New ADODB.Recordset
Dim rcConsultaDatGen As New ADODB.Recordset
Dim crSMinimo As Double, crUdi As Double
Dim Sql2 As String

Private Sub cmbMesReporte_Change()
    txtDesde.text = "01/" & Format(cmbMesReporte.ItemData(cmbMesReporte.ListIndex), "00") & "/" & Format(TxtAgnoReporte.text, "0000")
    txtHasta.text = Format(DateAdd("D", -1, Format("01/" & Format(cmbMesReporte.ItemData(cmbMesReporte.ListIndex) + 1, "00") & "/" & Format(TxtAgnoReporte.text, "0000"), "DD/MM/YYYY")), "DD/MM/YYYY")
End Sub

Private Sub cmbMesReporte_Click()
    txtDesde.text = "01/" & Format(cmbMesReporte.ItemData(cmbMesReporte.ListIndex), "00") & "/" & Format(TxtAgnoReporte.text, "0000")
    txtHasta.text = Format(DateAdd("D", -1, Format("01/" & Format(cmbMesReporte.ItemData(cmbMesReporte.ListIndex) + 1, "00") & "/" & Format(TxtAgnoReporte.text, "0000"), "DD/MM/YYYY")), "DD/MM/YYYY")
End Sub

Private Sub cmdBuscar_Click()
    Dim sqlAuto As String
    If opSalarios.Value = False And opUdis.Value = False And opPrestamo.Value = False And opCheque.Value = False Then
        
        MsgBox "Seleccione una opción de búsqueda...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    
    ElseIf opSalarios.Value And Trim(txtSalario.text) = "" Then
        
        MsgBox "Introduzca la cantidad de salarios mínimos...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
        
    ElseIf opUdis.Value And Trim(txtUdis.text) = "" Then
        
        MsgBox "Introduzca la cantidad de udis...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    
    ElseIf opPrestamo.Value And Trim(txtPrestamo.text) = "" Then
        
        MsgBox "Introduzca el importe del préstamo...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
        
    End If
    '***********************************************  SQL para el caso de que sea con autos ************************
                            
    grdClientes.Redraw = False
    grdClientes.Clear
                            
    grdDetalle.Redraw = True
    grdDetalle.Clear
                            
    Me.Refresh
                            
    Sql2 = "SELECT c.ID, CONCAT(c.nombre,' ',c.apellido) AS cliente, CONCAT(c.direccion,' COL. ',c.colonia) AS DireccionCompleta, c.tel,SUM(e.prestamo) AS TotPrestamo ,ta.Clave, e.IdTipoAlerta " & _
           "FROM clientes c INNER JOIN empeno e ON c.ID=e.IDCliente LEFT JOIN mld_prestamos_tipo_alertas ta ON e.IdTipoAlerta = ta.Id " & _
           "WHERE e.cancelado=0 AND e.origen=1 AND DATE(fecha)>='" & Format(txtDesde.text, "yyyy/MM/dd") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "yyyy/MM/dd") & "' " & _
           "GROUP BY e.idcliente ,ta.Clave "
                                        
    rcConsulta.Open Sql2, dbDatos, adOpenForwardOnly, adLockReadOnly
                    
                    
                    'MsgBox rcConsulta.RecordCount
    While Not rcConsulta.EOF
        
        If opCheque.Value Then
            
            GoTo Agregar
        
        ElseIf CDbl(rcConsulta!TotPrestamo) > CDbl(GetCriterio) Then
                  
           GoTo Agregar
           
        Else
             
            GoTo Siguiente
        End If
        
Agregar:
        grdClientes.AddRow
        grdClientes.CellText(grdClientes.Rows, 1) = rcConsulta!Cliente
        grdClientes.CellItemData(grdClientes.Rows, 1) = rcConsulta!ID
        grdClientes.CellFont(grdClientes.Rows, 1) = sFntUnread
        grdClientes.CellText(grdClientes.Rows, 2) = rcConsulta!DireccionCompleta
        grdClientes.CellFont(grdClientes.Rows, 2) = sFntUnread
        
        grdClientes.CellText(grdClientes.Rows, 3) = rcConsulta!Tel
        grdClientes.CellFont(grdClientes.Rows, 3) = sFntUnread
        
        grdClientes.CellText(grdClientes.Rows, 4) = rcConsulta!TotPrestamo
        grdClientes.CellFont(grdClientes.Rows, 4) = sFntUnread
        grdClientes.CellTextAlign(grdClientes.Rows, 4) = DT_RIGHT
            
        grdClientes.CellItemData(grdClientes.Rows, 5) = rcConsulta!IDTipoAlerta
        grdClientes.CellText(grdClientes.Rows, 5) = rcConsulta!Clave
        grdClientes.CellFont(grdClientes.Rows, 5) = sFntUnread
        grdClientes.CellTextAlign(grdClientes.Rows, 5) = DT_CENTER
            
            
Siguiente:
    rcConsulta.MoveNext
    Wend
    rcConsulta.Close
    grdClientes.Redraw = True
    
End Sub

Private Sub cmdDec_Click()
    TxtAgnoReporte.text = TxtAgnoReporte.text - 1
End Sub

Private Sub cmdGenerarXML_Click()
     GenerarArchivoXML
End Sub

Private Sub cmdInc_Click()
    TxtAgnoReporte.text = TxtAgnoReporte.text + 1
End Sub

Private Sub cmdMosFecha_Click(Index As Integer)
    
    If Index = 0 Then
        
        txtDesde.text = frmCalendario.Fecha(txtDesde.text)
        txtHasta.text = Format(DateAdd("D", -1, DateAdd("M", 6, CDate(txtDesde.text))), "DD/MM/YYYY")
        
    Else
        
        txtHasta.text = frmCalendario.Fecha(txtHasta.text)
    
    End If
    
End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub FlaGenerarXML_Click()
    
   GenerarArchivoXML
    
Exit Sub
    
   Dim sqlEncabezadoAuto As String
    If opSalarios.Value = False And opUdis.Value = False And opPrestamo.Value = False And opCheque.Value = False Then
        
        MsgBox "Seleccione una opción de búsqueda...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    
    ElseIf opSalarios.Value And Trim(txtSalario.text) = "" Then
        
        MsgBox "Introduzca la cantidad de salarios mínimos...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
        
    ElseIf opUdis.Value And Trim(txtUdis.text) = "" Then
        
        MsgBox "Introduzca la cantidad de udis...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    
    ElseIf opPrestamo.Value And Trim(txtPrestamo.text) = "" Then
        
        MsgBox "Introduzca el importe del préstamo...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
        
    ElseIf Trim(Regresa_Valor_BD("RutaArchivosXML")) = "" Then
        MsgBox "No se ha especificado la Ruta de Generación de Archivos XML...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    
    ElseIf Dir(Regresa_Valor_BD("RutaArchivosXML"), vbDirectory) = "" Then
        MsgBox "No se encontró la Ruta de Generación de Archivos XML...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    
    End If
   ' Dim objDOM As New MSXML2.DOMDocument30
   'Dim objNode As MSXML2.IXMLDOMNode
   'Dim objChildNode As MSXML2.IXMLDOMNode
   'Dim objGrandChildNode As MSXML2.IXMLDOMNode
   'Dim objAttribute As MSXML2.IXMLDOMAttribute
   'Dim objElement As MSXML2.IXMLDOMElement
   
   'Dim Doc As MSXML2.DOMDocument40
   'Dim Nod(2) As MSXML2.IXMLDOMNode
   Dim archivoXML As String
   Dim personaFisica As Boolean
   Dim folioalerta As Integer
   Dim cero As Boolean
   
   Dim vDependencia As String, vClaveIdent As String, vClaveOcup As String, vClavePais As String, vClaveAlerta As String, vClaveMoneda As String
   Dim vTipoOper As Integer, vInstMon As Integer
   Dim Regs As Integer
   Dim vNombreSucursal As String
   
    Dim SqlEncabezado, sqlTotal, SqlDetalle, sqlCeros1, sqlSucursalDatgen, SqlDetalleAuto As String
    SqlEncabezado = ""
    cero = False
    Regs = 0
'        SqlEncabezado = "SELECT `empeno`.`IdCliente`,`clientes`.`personafisica`,`sucursales`.`RFC` as sujetoobligado, `mld_prestamos_tipo_alertas`.`Clave` AS ClaveAlerta , `clientes`.`Nombre`, `clientes`.`ApellidoPaterno` , `clientes`.`ApellidoMaterno`, `clientes`.`FecNac` " & _
'    ", `clientes`.`Rfc` as rfccliente, `clientes`.`Curp` as Clientecurp, `mld_paises`.`Clave` as nacionalidad, `mld_paises_1`.`Clave` as nacimiento, `mld_actividades_economicas`.`Clave` actividadeconomica, `mld_tipo_identificaciones`.`Clave` as identificacliente, `mld_tipo_identificaciones`.`Dependencia`" & _
'    ", `clientes`.`NumeroIdentificacion`, `clientes`.`RL_Nombre`, `clientes`.`RL_ApellidoPaterno`, `clientes`.`RL_ApellidoMaterno`, `clientes`.`RazonSocial`,`clientes`.`FechaAltaRazonSocial`,`clientes`.`RL_Rfc`, `clientes`.`RL_Curp`" & _
'    ", `clientes`.`Colonia`, `clientes`.`Direccion`, `clientes`.`NoExterior`, `clientes`.`NoInterior`, `clientes`.`CP` as ClienteCP, `mld_paises`.`Clave`, `clientes`.`Tel`, `clientes`.`Email`" & _
'    ", `empeno`.`Fecha`, `sucursales`.`Cp` as sucursalCP, `sucursales`.`RazonSocial` as sucursalRazonSocial, `usuarios`.`Nombre`,`sucursales`.`NombreComercial`, `empeno`.`IDTipoOperacion`, `empeno`.`Vencimiento`" & _
'    ", `empeno`.`IDInstrumentoMonetario`, `empeno`.`IDTipoMoneda`, SUM(`empeno`.`Prestamo`) totprestamo, `empeno`.`DescTipoAlerta` from `basedatos`.`sucursales`,`basedatos`.`empeno` LEFT JOIN `basedatos`.`clientes`" & _
'        " ON (`empeno`.`IDCliente` = `clientes`.`ID`) LEFT JOIN `basedatos`.`mld_paises` ON (`clientes`.`IdPaisNacionalidad` = `mld_paises`.`Id`) LEFT JOIN `basedatos`.`mld_paises` AS `mld_paises_1` " & _
'        " ON (`clientes`.`IdPaisNacimiento` = `mld_paises_1`.`Id`) LEFT JOIN `basedatos`.`mld_actividades_economicas` ON (`clientes`.`IdOcupacion` = `mld_actividades_economicas`.`Id`) " & _
'    " LEFT JOIN `basedatos`.`mld_tipo_identificaciones` ON (`clientes`.`IdTipoIdent` = `mld_tipo_identificaciones`.`Id`) LEFT JOIN `basedatos`.`mld_prestamos_tipo_alertas` ON (`empeno`.`IdTipoAlerta` = `mld_prestamos_tipo_alertas`.`Id`) " & _
'    "  LEFT JOIN `basedatos`.`usuarios` ON (`empeno`.`IDUsuarioMov` = `usuarios`.`ID`) where `empeno`.`cancelado`=0 AND `empeno`.`origen`=1  and DATE(fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' GROUP BY `clientes`.`Rfc`;"

    vClaveAlerta = Trim(SacaValor("mld_prestamos_tipo_alertas", "Clave", " WHERE RegDefault=1"))
    vClavePais = Trim(SacaValor("mld_paises", "Clave", " WHERE RegDefault=1"))
    vClaveOcup = Trim(SacaValor("mld_actividades_economicas", "Clave", " WHERE RegDefault=1"))
    vClaveIdent = Trim(SacaValor("mld_tipo_identificaciones", "Clave", " WHERE RegDefault=1"))
    vDependencia = UCase(Trim(SacaValor("mld_tipo_identificaciones", "Dependencia", " WHERE RegDefault=1")))
    vTipoOper = Val(SacaValor("mld_prestamos_tipo_operacion", "Clave", " WHERE RegDefault=1"))
    vInstMon = Val(SacaValor("mld_instr_monetarios", "Clave", " WHERE RegDefault=1")) 'mld_instr_monetarios
    vClaveMoneda = Trim(SacaValor("mld_tipo_monedas", "Clave", " WHERE MonedaDefault=1"))
    
    
    rcConsultaCero.Open "select count(IDempeno)as cuenta from detallesempenoautos", dbDatos, adOpenForwardOnly, adLockReadOnly
    If rcConsultaCero.Fields("cuenta") > 0 Then
        sqlEncabezadoAuto = " union " + Chr(13) + "(SELECT `empeno`.`Id` as IdEmpeno,`empeno`.`IdCliente`,`empeno`.`Serie`,`clientes`.`personafisica`,`sucursales`.`RFC` as sujetoobligado,if(`mld_prestamos_tipo_alertas`.`Clave` is null,'100',`mld_prestamos_tipo_alertas`.`Clave`) AS ClaveAlerta , `clientes`.`Nombre` as Nombre,if(`clientes`.`ApellidoPaterno` is null or `clientes`.`ApellidoPaterno`='',`clientes`.`Apellido`,`clientes`.`ApellidoPaterno`) as ApellidoPaterno , `clientes`.`ApellidoMaterno` as ApellidoMaterno, `clientes`.`Apellido` as Apellido,`clientes`.`FecNac` , `clientes`.`Rfc` as rfccliente, `clientes`.`Curp` as Clientecurp, if(`mld_paises`.`Clave` is null,'MX',`mld_paises`.`Clave`) as nacionalidad,if(`mld_paises_1`.`Clave` is null,'MX',`mld_paises_1`.`Clave`) as nacimiento, if(`mld_actividades_economicas`.`Clave` is null,'9999999',`mld_actividades_economicas`.`Clave`) as actividadeconomica,if(`mld_tipo_identificaciones`.`Clave` is null,'1'," + Chr(13) & _
                            "`mld_tipo_identificaciones`.`Clave`) as identificacliente,if(`mld_tipo_identificaciones`.`Dependencia` is null,'INSTITUTO FEDERAL ELECTORAL',`mld_tipo_identificaciones`.`Dependencia`) as Dependencia, `clientes`.`NumeroIdentificacion`,if(`clientes`.`RL_Nombre` is null,'',`clientes`.`RL_Nombre`) As RL_Nombre,if(`clientes`.`RL_ApellidoPaterno` is null,'',`clientes`.`RL_ApellidoPaterno`) AS RL_ApellidoPaterno, if(`clientes`.`RL_ApellidoMaterno` is null,'',`clientes`.`RL_ApellidoMaterno`) AS RL_ApellidoMaterno, `clientes`.`RazonSocial`,`clientes`.`FechaAltaRazonSocial`,`clientes`.`RL_Rfc`, `clientes`.`RL_Curp`, `clientes`.`Colonia`, `clientes`.`Direccion`, `clientes`.`NoExterior`, `clientes`.`NoInterior`, `clientes`.`CP` as ClienteCP, if(`mld_paises`.`Clave` is null, 'MX',`mld_paises`.`Clave`) as Clave, `clientes`.`Tel`, `clientes`.`Email`, `empeno`.`Fecha`, `sucursales`.`Cp` as sucursalCP, " + Chr(13) & _
                            "`sucursales`.`RazonSocial` as sucursalRazonSocial, `usuarios`.`Nombre` as NombreUsuario,`sucursales`.`NombreComercial`,`empeno`.`IDTipoOperacion`,`empeno`.`Vencimiento`,if(`mld_instr_monetarios`.`Clave` is null,'1',`mld_instr_monetarios`.`Clave`) as IDInstrumentoMonetario,if(`mld_tipo_monedas`.`Clave` is null,'MXN',`mld_tipo_monedas`.`Clave`) as IdTipoMoneda,SUM(`empeno`.`Prestamo`) totprestamo,`empeno`.`DescTipoAlerta` from `basedatos`.`sucursales`,`basedatos`.`empeno` LEFT JOIN `basedatos`.`clientes` ON (`empeno`.`IDCliente` = `clientes`.`ID`) LEFT JOIN `basedatos`.`mld_paises` ON (`clientes`.`IdPaisNacionalidad` = `mld_paises`.`Id`) LEFT JOIN `basedatos`.`mld_paises` AS `mld_paises_1`  ON (`clientes`.`IdPaisNacimiento` = `mld_paises_1`.`Id`) LEFT JOIN `basedatos`.`mld_actividades_economicas` ON (`clientes`.`IdOcupacion` = `mld_actividades_economicas`.`Id`)  " + Chr(13) & _
                            "LEFT JOIN `basedatos`.`mld_tipo_identificaciones` ON (`clientes`.`IdTipoIdent` = `mld_tipo_identificaciones`.`Id`)  LEFT JOIN `basedatos`.`mld_prestamos_tipo_alertas` ON (`empeno`.`IdTipoAlerta` = `mld_prestamos_tipo_alertas`.`Id`) LEFT JOIN `basedatos`.`usuarios` ON (`empeno`.`IDUsuarioMov` = `usuarios`.`ID`) LEFT JOIN `basedatos`.`mld_instr_monetarios` ON (`empeno`.`IdInstrumentoMonetario` = `mld_instr_monetarios`.`Id`) LEFT JOIN `basedatos`.`mld_tipo_monedas` ON (`empeno`.`IdTipoMoneda` = `mld_tipo_monedas`.`Id`) where  empeno.serie=2 and `empeno`.`cancelado`=0 AND `empeno`.`origen`=1 and DATE(fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' GROUP BY `empeno`.`IDCliente`,`empeno`.`ID`)"
 
         vNombreSucursal = Trim(SacaValor("sucursales", "NombreComercial", " WHERE Activa=1"))
        
    Else
        sqlEncabezadoAuto = ""
    End If
    rcConsultaCero.Close
    SqlEncabezado = "(SELECT `empeno`.`Id` as IdEmpeno,`empeno`.`IdCliente`,`empeno`.`Serie`,`clientes`.`personafisica`,`sucursales`.`RFC` as sujetoobligado, if(`mld_prestamos_tipo_alertas`.`Clave` is null,'" & vClaveAlerta & "',`mld_prestamos_tipo_alertas`.`Clave`) AS ClaveAlerta , `clientes`.`Nombre` as Nombre, if(`clientes`.`ApellidoPaterno` is null or `clientes`.`ApellidoPaterno`='',`clientes`.`Apellido`,`clientes`.`ApellidoPaterno`) as ApellidoPaterno , " + Chr(13) & _
                    "`clientes`.`ApellidoMaterno` as ApellidoMaterno, `clientes`.`Apellido` as Apellido,`clientes`.`FecNac` , `clientes`.`Rfc` as rfccliente, `clientes`.`Curp` as Clientecurp, if(`mld_paises`.`Clave` is null,'" & Trim(SacaValor("mld_paises", "Clave", " WHERE RegDefault=1")) & "',`mld_paises`.`Clave`) as nacionalidad, if(`mld_paises_1`.`Clave` is null,'" & vClavePais & "',`mld_paises_1`.`Clave`) as nacimiento, " + Chr(13) & _
                    "if(`mld_actividades_economicas`.`Clave` is null,'" & vClaveOcup & "',`mld_actividades_economicas`.`Clave`) as actividadeconomica, if(`mld_tipo_identificaciones`.`Clave` is null,'" & vClaveIdent & "',`mld_tipo_identificaciones`.`Clave`) as identificacliente, if(`mld_tipo_identificaciones`.`Dependencia` is null,'" & vDependencia & "',`mld_tipo_identificaciones`.`Dependencia`) as Dependencia, `clientes`.`NumeroIdentificacion`, " + Chr(13) & _
                    "if(`clientes`.`RL_Nombre` is null,'',`clientes`.`RL_Nombre`) As RL_Nombre, if(`clientes`.`RL_ApellidoPaterno` is null,'',`clientes`.`RL_ApellidoPaterno`) AS RL_ApellidoPaterno, if(`clientes`.`RL_ApellidoMaterno` is null,'',`clientes`.`RL_ApellidoMaterno`) AS RL_ApellidoMaterno, `clientes`.`RazonSocial`,`clientes`.`FechaAltaRazonSocial`,`clientes`.`RL_Rfc`, `clientes`.`RL_Curp`, " + Chr(13) & _
                    "`clientes`.`Colonia`, `clientes`.`Direccion`, `clientes`.`NoExterior`, `clientes`.`NoInterior`, `clientes`.`CP` as ClienteCP, if(`mld_paises`.`Clave` is null, '" & vClavePais & "',`mld_paises`.`Clave`) as Clave, `clientes`.`Tel`, `clientes`.`Email`, `empeno`.`Fecha`, `sucursales`.`Cp` as sucursalCP," + Chr(13) & _
                    " `sucursales`.`RazonSocial` as sucursalRazonSocial, `usuarios`.`Nombre` as NombreUsuario,`sucursales`.`NombreComercial`, `empeno`.`IDTipoOperacion`, `empeno`.`Vencimiento`, if(`mld_instr_monetarios`.`Clave` is null,'" & vInstMon & "',`mld_instr_monetarios`.`Clave`) as IDInstrumentoMonetario, if(`mld_tipo_monedas`.`Clave` is null, '" & vClaveMoneda & "',`mld_tipo_monedas`.`Clave`) as IdTipoMoneda, SUM(`empeno`.`Prestamo`) totprestamo, `empeno`.`DescTipoAlerta` " + Chr(13) & _
                    "from `basedatos`.`sucursales`,`basedatos`.`empeno` LEFT JOIN `basedatos`.`detallesempeno` ON (`empeno`.`ID` = `detallesempeno`.`IDEmpeno`)  INNER JOIN `basedatos`.`tipo` ON `tipo`.`ID` = `detallesempeno`.`tipo` LEFT JOIN `basedatos`.`clientes`" + Chr(13) & _
                    " ON (`empeno`.`IDCliente` = `clientes`.`ID`) LEFT JOIN `basedatos`.`mld_paises` ON (`clientes`.`IdPaisNacionalidad` = `mld_paises`.`Id`) LEFT JOIN `basedatos`.`mld_paises` AS `mld_paises_1` " + Chr(13) & _
                    " ON (`clientes`.`IdPaisNacimiento` = `mld_paises_1`.`Id`) LEFT JOIN `basedatos`.`mld_actividades_economicas` ON (`clientes`.`IdOcupacion` = `mld_actividades_economicas`.`Id`) " + Chr(13) & _
                    " LEFT JOIN `basedatos`.`mld_tipo_identificaciones` ON (`clientes`.`IdTipoIdent` = `mld_tipo_identificaciones`.`Id`) LEFT JOIN `basedatos`.`mld_prestamos_tipo_alertas` ON (`empeno`.`IdTipoAlerta` = `mld_prestamos_tipo_alertas`.`Id`) " + Chr(13) & _
                    " LEFT JOIN `basedatos`.`usuarios` ON (`empeno`.`IDUsuarioMov` = `usuarios`.`ID`) LEFT JOIN `basedatos`.`mld_instr_monetarios` ON (`empeno`.`IdInstrumentoMonetario` = `mld_instr_monetarios`.`Id`) LEFT JOIN `basedatos`.`mld_tipo_monedas` ON (`empeno`.`IdTipoMoneda` = `mld_tipo_monedas`.`Id`) " + Chr(13) & _
                    "where  tipo.IdTipoGarantia <> 0 and `empeno`.`cancelado`=0 AND `empeno`.`origen`=1  and DATE(fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' GROUP BY `empeno`.`IDCliente`,`empeno`.`ID`)"
                    
                    
    'SqlEncabezado = "SELECT `empeno`.`Id` as IdEmpeno,`empeno`.`IdCliente`,`empeno`.`Serie`,`clientes`.`personafisica`,`sucursales`.`RFC` as sujetoobligado, if(`mld_prestamos_tipo_alertas`.`Clave` is null,'" & vClaveAlerta & "',`mld_prestamos_tipo_alertas`.`Clave`) AS ClaveAlerta , `clientes`.`Nombre` as Nombre, if(`clientes`.`ApellidoPaterno` is null or `clientes`.`ApellidoPaterno`='',`clientes`.`Apellido`,`clientes`.`ApellidoPaterno`) as ApellidoPaterno , " & _
                    "`clientes`.`ApellidoMaterno` as ApellidoMaterno, `clientes`.`Apellido` as Apellido,`clientes`.`FecNac` , `clientes`.`Rfc` as rfccliente, `clientes`.`Curp` as Clientecurp, if(`mld_paises`.`Clave` is null,'" & Trim(SacaValor("mld_paises", "Clave", " WHERE RegDefault=1")) & "',`mld_paises`.`Clave`) as nacionalidad, if(`mld_paises_1`.`Clave` is null,'" & vClavePais & "',`mld_paises_1`.`Clave`) as nacimiento, " & _
                    "if(`mld_actividades_economicas`.`Clave` is null,'" & vClaveOcup & "',`mld_actividades_economicas`.`Clave`) as actividadeconomica, if(`mld_tipo_identificaciones`.`Clave` is null,'" & vClaveIdent & "',`mld_tipo_identificaciones`.`Clave`) as identificacliente, if(`mld_tipo_identificaciones`.`Dependencia` is null,'" & vDependencia & "',`mld_tipo_identificaciones`.`Dependencia`) as Dependencia, `clientes`.`NumeroIdentificacion`, " & _
                    "if(`clientes`.`RL_Nombre` is null,'',`clientes`.`RL_Nombre`) As RL_Nombre, if(`clientes`.`RL_ApellidoPaterno` is null,'',`clientes`.`RL_ApellidoPaterno`) AS RL_ApellidoPaterno, if(`clientes`.`RL_ApellidoMaterno` is null,'',`clientes`.`RL_ApellidoMaterno`) AS RL_ApellidoMaterno, `clientes`.`RazonSocial`,`clientes`.`FechaAltaRazonSocial`,`clientes`.`RL_Rfc`, `clientes`.`RL_Curp`, " & _
                    "`clientes`.`Colonia`, `clientes`.`Direccion`, `clientes`.`NoExterior`, `clientes`.`NoInterior`, `clientes`.`CP` as ClienteCP, if(`mld_paises`.`Clave` is null, '" & vClavePais & "',`mld_paises`.`Clave`) as Clave, `clientes`.`Tel`, `clientes`.`Email`, `empeno`.`Fecha`, `sucursales`.`Cp` as sucursalCP," & _
                    " `sucursales`.`RazonSocial` as sucursalRazonSocial, `usuarios`.`Nombre` as NombreUsuario,`sucursales`.`NombreComercial`, `empeno`.`IDTipoOperacion`, `empeno`.`Vencimiento`, if(`mld_instr_monetarios`.`Clave` is null,'" & vInstMon & "',`mld_instr_monetarios`.`Clave`) as IDInstrumentoMonetario, if(`mld_tipo_monedas`.`Clave` is null, '" & vClaveMoneda & "',`mld_tipo_monedas`.`Clave`) as IdTipoMoneda, SUM(`empeno`.`Prestamo`) totprestamo, `empeno`.`DescTipoAlerta` " & _
                    "from `basedatos`.`sucursales`,`basedatos`.`empeno` LEFT JOIN `basedatos`.`detallesempeno` ON (`empeno`.`ID` = `detallesempeno`.`IDEmpeno`)  INNER JOIN `basedatos`.`tipo` ON `tipo`.`ID` = `detallesempeno`.`tipo` LEFT JOIN `basedatos`.`clientes`" & _
                    " ON (`empeno`.`IDCliente` = `clientes`.`ID`) LEFT JOIN `basedatos`.`mld_paises` ON (`clientes`.`IdPaisNacionalidad` = `mld_paises`.`Id`) LEFT JOIN `basedatos`.`mld_paises` AS `mld_paises_1` " & _
                    " ON (`clientes`.`IdPaisNacimiento` = `mld_paises_1`.`Id`) LEFT JOIN `basedatos`.`mld_actividades_economicas` ON (`clientes`.`IdOcupacion` = `mld_actividades_economicas`.`Id`) " & _
                    " LEFT JOIN `basedatos`.`mld_tipo_identificaciones` ON (`clientes`.`IdTipoIdent` = `mld_tipo_identificaciones`.`Id`) LEFT JOIN `basedatos`.`mld_prestamos_tipo_alertas` ON (`empeno`.`IdTipoAlerta` = `mld_prestamos_tipo_alertas`.`Id`) " & _
                    " LEFT JOIN `basedatos`.`usuarios` ON (`empeno`.`IDUsuarioMov` = `usuarios`.`ID`) LEFT JOIN `basedatos`.`mld_instr_monetarios` ON (`empeno`.`IdInstrumentoMonetario` = `mld_instr_monetarios`.`Id`) LEFT JOIN `basedatos`.`mld_tipo_monedas` ON (`empeno`.`IdTipoMoneda` = `mld_tipo_monedas`.`Id`) " & _
                    "where  tipo.IdTipoGarantia <> 0 and `empeno`.`cancelado`=0 AND `empeno`.`origen`=1  and DATE(fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' GROUP BY `empeno`.`IDCliente`;"
    
    
    
    sqlSucursalDatgen = "SELECT `mld_actividad_vulnerable`.`Clave` as ActividadVulnerable , `mld_giro_mercantil`.`Descripcion` as giromercantil From `basedatos`.`parametros` " & _
                        "INNER JOIN `basedatos`.`mld_actividad_vulnerable` ON (`parametros`.`IDActividadVulnerable` = `mld_actividad_vulnerable`.`Id`) INNER JOIN `basedatos`.`mld_giro_mercantil` ON (`parametros`.`IdTipoGiroMercantil` = `mld_giro_mercantil`.`Id`)"
                        
          sqlTotal = SqlEncabezado + Chr(13) + sqlEncabezadoAuto
    rcConsulta2.Open sqlTotal, dbDatos, adOpenForwardOnly, adLockReadOnly
    
    rcConsultaDatGen.Open sqlSucursalDatgen, dbDatos, adOpenForwardOnly, adLockReadOnly
    
    If rcConsulta2.RecordCount <= 0 Then
       cero = True
       sqlCeros1 = "SELECT `RazonSocial`,`Clave`, `NombreComercial`, `RFC`, `Direccion`, `Ciudad`, `Estado`, `Telefono` , `Cp` from `basedatos`.`sucursales`;"
       rcConsultaCero.Open sqlCeros1, dbDatos, adOpenForwardOnly, adLockReadOnly
    End If
    
    'RUTA DE ARCHIVO XML
    archivoXML = Regresa_Valor_BD("RutaArchivosXML") & "\" & CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00")) & ".xml"
   '******************* Se Crea el archivo ****************************
    'Set Doc = New MSXML2.DOMDocument       'Iniciar documento XML y nodo raíz
    'Doc.appendChild Doc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
  ' Set Nod(0) = Doc.createElement("Archivo").
   
   'Set Nod(1) = Doc.createElement("Informe")
   '**********************************************************************
   'Set Nod(2) = Doc.createElement("Mes_reportado")
   'Nod(2).text = CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00"))
   'Nod(1).appendChild Nod(2)
   'Set Nod(1) = Doc.createElement(CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "MM")))
   'Nod(0).appendChild Nod(1)
      
   
   
   'Doc.appendChild Nod(0)                 'Agregar el nodo <Informe> al documento XML
   'Doc.Save App.Path & "\LAVADO.xml"
   Open archivoXML For Output As #1
  
   Print #1, "<?xml version='1.0' encoding='UTF-8' ?>"
    Print #1, "<archivo xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation = 'http://www.uif.shcp.gob.mx/recepcion/mpc mpc.xsd' " & _
            " xmlns='http://www.uif.shcp.gob.mx/recepcion/mpc'> "
    Print #1, " <informe>"
    Print #1, "  <mes_reportado>" & CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00")) & "</mes_reportado>"
    Print #1, "  <sujeto_obligado>"
    If cero = False Then
        Print #1, "     <clave_sujeto_obligado>" & rcConsulta2.Fields("sujetoobligado") & "</clave_sujeto_obligado>"
        Print #1, "     <clave_actividad>" & rcConsultaDatGen.Fields("ActividadVulnerable") & "</clave_actividad>"
    Else
        Print #1, "     <clave_sujeto_obligado>" & rcConsultaCero.Fields("rfc") & "</clave_sujeto_obligado>"
        Print #1, "     <clave_actividad>MPC</clave_actividad>"
    End If
    
    Print #1, "  </sujeto_obligado>"
    
   
    Print #1, "  <aviso>"
    Print #1, "    <referencia_aviso>" & Regresa_Movimiento(True, "FolioAvisosLavado") & "</referencia_aviso>"
    Print #1, "    <prioridad>1</prioridad>"
    Print #1, "    <alerta>"
    
    If cero = False Then
        Print #1, "       <tipo_alerta>" & rcConsulta2.Fields("ClaveAlerta") & "</tipo_alerta:element>"
    Else
        Print #1, "       <tipo_alerta>100</tipo_alerta>"
    
    End If
    Print #1, "    </alerta>"
    
    'DATOS DEL AVISO
    
    
   While Not rcConsulta2.EOF
         If rcConsulta2.Fields("totprestamo") > GetCriterio Then
               
               Regs = Regs + 1
               cero = False
              
              
               Print #1, "   <persona_aviso>"
               Print #1, "     <tipo_persona>"
               '*************************** PERSONA FISICA ****************************
               If rcConsulta2.Fields("personafisica") = 1 Then
                    
                    Print #1, "         <persona_fisica>"
                    Print #1, "             <nombre>" & LimpiaCad(rcConsulta2.Fields("Nombre")) & "</nombre>"
                    Print #1, "             <apellido_paterno>" & LimpiaCad(rcConsulta2.Fields("ApellidoPaterno")) & "</apellido_paterno>"
                    Print #1, "             <apellido_materno>" & LimpiaCad(rcConsulta2.Fields("ApellidoMaterno")) & "</apellido_materno>"
                    Print #1, "             <fecha_nacimiento>" & Format(rcConsulta2.Fields("fecnac"), "yyyyMMdd") & "</fecha_nacimiento>"
                    Print #1, "             <rfc>" & rcConsulta2.Fields("rfccliente") & "</rfc>"
                    Print #1, "             <curp>" & IIf(IsNull(rcConsulta2.Fields("clientecurp")), "", rcConsulta2.Fields("clientecurp")) & "</curp>"
                    Print #1, "             <pais_nacionalidad>" & rcConsulta2.Fields("nacionalidad") & "</pais_nacionalidad>"
                    Print #1, "             <pais_nacimiento>" & rcConsulta2.Fields("nacimiento") & "</pais_nacimiento>"
                    Print #1, "             <actividad_economica>" & rcConsulta2.Fields("actividadeconomica") & "</actividad_economica>"
                    Print #1, "             <tipo_identificacion>" & Trim(rcConsulta2.Fields("identificacliente")) & "</tipo_identificacion>"
                    Print #1, "             <autoridad_identificacion>" & LimpiaCad(UCase(rcConsulta2.Fields("dependencia"))) & "'</autoridad_identificacion>"
                    Print #1, "             <numero_identificacion>" & rcConsulta2.Fields("NumeroIdentificacion") & "</numero_identificacion>"
                    Print #1, "         </persona_fisica>"
               Else
                    Print #1, "         <persona_moral>"
                    Print #1, "            <denominacion_razon>" & Trim(LimpiaCad(rcConsulta2.Fields("RazonSocial"))) & "</denominacion_razon>"
                    Print #1, "            <fecha_constitucion>" & Format(rcConsulta2.Fields("FechaAltaRazonSocial"), "yyyyMMdd") & "</fecha_constitucion>"
                    Print #1, "            <rfc>" & rcConsulta2.Fields("rfccliente") & "</rfc>"
                    Print #1, "            <pais_nacionalidad>" & Trim(rcConsulta2.Fields("nacionalidad")) & "</pais_nacionalidad>"
                    Print #1, "            <giro_mercantil>" & rcConsulta2.Fields("actividadeconomica") & "</giro_mercantil>"
                    
                    Print #1, "            <representante_apoderado>"
                    Print #1, "                 <nombre>" & LimpiaCad(rcConsulta2.Fields("RL_Nombre")) & "</nombre>"
                    Print #1, "                 <apellido_paterno>" & LimpiaCad(rcConsulta2.Fields("RL_ApellidoPaterno")) & "</apellido_paterno>"
                    Print #1, "                 <apellido_materno>" & LimpiaCad(rcConsulta2.Fields("RL_ApellidoMaterno")) & "</apellido_materno>"
                    Print #1, "                 <fecha_nacimiento>" & Format(rcConsulta2.Fields("FecNAc"), "yyyyMMdd") & "</fecha_nacimiento>"
                    Print #1, "                 <rfc>" & rcConsulta2.Fields("RFC") & "</rfc>"
                    Print #1, "                 <curp>" & rcConsulta2.Fields("Clientecurp") & "</curp>"
                    Print #1, "                 <tipo_identificacion>" & Trim(rcConsulta2.Fields("identificacliente")) & "</tipo_identificacion>"
                    'Print #1, "                 <identificacion_otro></identificacion_otro>"
                    Print #1, "                 <autoridad_identificacion>" & LimpiaCad(UCase(rcConsulta2.Fields("dependencia"))) & "</autoridad_identificacion>"
                    Print #1, "                 <numero_identificacion>" & rcConsulta2.Fields("NumeroIdentificacion") & "</numero_identificacion>"
                    Print #1, "            </representante_apoderado>"
                    Print #1, "         </persona_moral>"
                    
                    
               End If
              
              Print #1, "     </tipo_persona>"
              Print #1, "     <tipo_domicilio>"
              Print #1, "        <nacional>"
              Print #1, "           <colonia>" & LimpiaCad(rcConsulta2.Fields("Colonia")) & "</colonia>"
              Print #1, "           <calle>" & LimpiaCad(rcConsulta2.Fields("Direccion")) & "</calle>"
              Print #1, "           <numero_exterior>" & rcConsulta2.Fields("NoExterior") & "</numero_exterior>"
              Print #1, "           <numero_interior>" & rcConsulta2.Fields("NoInterior") & "</numero_interior>"
              Print #1, "           <codigo_postal>" & rcConsulta2.Fields("CLienteCP") & "</codigo_postal>"
              Print #1, "        </nacional>"
              Print #1, "     </tipo_domicilio>"
              Print #1, "     <telefono>"
              Print #1, "         <clave_pais>MX</clave_pais>"
              Print #1, "         <numero_telefono>" & rcConsulta2.Fields("Tel") & "</numero_telefono>"
              Print #1, "         <correo_electronico>" & rcConsulta2.Fields("Email") & "</correo_electronico>"
              Print #1, "     </telefono>"
              Print #1, "   </persona_aviso>"
                   
                   
              
               
               '************************************ OPERACIONES *****************************
                Print #1, "   <detalle_operaciones>"
                Print #1, "        <operaciones_realizadas>"
               
              If rcConsulta2.Fields("serie") = 1 Then
                
                   SqlDetalle = "Select e.IDCliente,e.Fecha,e.Vencimiento,de.IDTipoGarantia,de.Prestamo AS DPrestamo,de.Articulo,e.IDTipoMoneda,e.IdTipoAlerta,e.DescTipoAlerta,e.IDInstrumentoMonetario," + Chr(13) & _
                            "tg.Clave AS ClaveTipoGarantia,if(tm.Clave is null,'" & vClaveMoneda & "',tm.Clave) AS ClaveTipoMoneda,if(im.Clave is null,'" & vInstMon & "',im.Clave) AS ClaveInstMonetario,if(ta.Clave is null,'" & vClaveAlerta & "',ta.Clave) AS ClaveTipoAlerta,if(toper.Clave is null,'" & vTipoOper & "',toper.Clave) as ClaveTipoOperacion " + Chr(13) & _
                            "From detallesempeno AS de Inner Join empeno AS e ON (de.IDEmpeno = e.ID) Left Join mld_tipo_monedas AS tm ON e.IDTipoMoneda = tm.ID Left Join mld_instr_monetarios AS im ON e.IDInstrumentoMonetario = im.Id Left Join mld_prestamos_tipo_alertas AS ta ON e.IdTipoAlerta = ta.Id " + Chr(13) & _
                            "Left Join mld_prestamos_tipo_operacion AS toper ON e.IDTipoOperacion = toper.Id Left Join tipo ON de.Tipo = tipo.ID Left Join mld_prestamos_tipo_garantias AS tg ON tipo.IdTipoGarantia = tg.Id " + Chr(13) & _
                            "WHERE e.idCliente = " & rcConsulta2.Fields("idCliente") & " AND de.IdEmpeno=" & rcConsulta2.Fields("idEmpeno") & ";"
                            
              Else
                SqlDetalle = "select e.IDCliente,e.Fecha,e.Vencimiento,de.IDTipoGarantia,e.Prestamo AS DPrestamo,concat(de.MarcaYModelo,' ',cast(año as char)) as articulo_otro,e.IDTipoMoneda,e.IdTipoAlerta,e.DescTipoAlerta,e.IDInstrumentoMonetario," + Chr(13) & _
                            " tg.Clave AS ClaveTipoGarantia,if(tm.Clave is null,'" & vClaveMoneda & "',tm.Clave) AS ClaveTipoMoneda,if(im.Clave is null,'" & vInstMon & "',im.Clave) AS ClaveInstMonetario,if(ta.Clave is null,'" & vClaveAlerta & "',ta.Clave) AS ClaveTipoAlerta,if(toper.Clave is null,'" & vTipoOper & "',toper.Clave) as ClaveTipoOperacion  from empeno as e " + Chr(13) & _
                            " left join detallesempenoautos as de on de.IDEmpeno = e.ID Left Join mld_tipo_monedas AS tm ON e.IDTipoMoneda = tm.ID  Left Join mld_instr_monetarios AS im ON e.IDInstrumentoMonetario = im.Id " + Chr(13) & _
                            " Left Join mld_prestamos_tipo_alertas AS ta ON e.IdTipoAlerta = ta.Id Left Join mld_prestamos_tipo_operacion AS toper ON e.IDTipoOperacion = toper.Id  Left Join tipo ON de.Tipo = tipo.ID " + Chr(13) & _
                            " Left Join mld_prestamos_tipo_garantias AS tg ON tipo.IdTipoGarantia = tg.Id WHERE e.idCliente = " & rcConsulta2.Fields("idCliente") & " AND de.IdEmpeno=" & rcConsulta2.Fields("idEmpeno") & " and e.serie=2;"
              End If
               rcConsulta3.Open SqlDetalle, dbDatos, adOpenForwardOnly, adLockReadOnly
               
               While Not rcConsulta3.EOF
                    Print #1, "        <datos_operacion>"
                    Print #1, "           <fecha_operacion>" & Format(rcConsulta3.Fields("Fecha"), "yyyyMMdd") & "</fecha_operacion>"
                    Print #1, "           <codigo_postal>" & rcConsulta2.Fields("clienteCP") & "</codigo_postal>"
                    Print #1, "           <nombre_sucursal>" & LimpiaCad(rcConsulta2.Fields("NombreComercial")) & "</nombre_sucursal>"
                    Print #1, "           <tipo_operacion>" & Trim(rcConsulta3.Fields("ClaveTipoOperacion")) & "</tipo_operacion>"
                    
                    Print #1, "           <datos_garantia>"
                    Print #1, "               <tipo_garantia>" & Trim(rcConsulta3.Fields("ClaveTipoGarantia")) & "</tipo_garantia>"
                    Print #1, "               <datos_bien_mutuo>"
                    If rcConsulta2.Fields("serie") = 2 Then
                         Print #1, "                        <datos_otro>"
                    
                         Print #1, "                              <descripcion_garantia>" & "Auto " & (rcConsulta3.Fields("articulo_otro")) & "</descripcion_garantia>"
                         Print #1, "                        </datos_otro>"
                    End If
                    Print #1, "               </datos_bien_mutuo>"
                    Print #1, "               <tipo_persona>"
                    Print #1, "                   <persona_fisica>"
                    Print #1, "                         <nombre>" & LimpiaCad(rcConsulta2.Fields("Nombre")) & "</nombre>"
                    Print #1, "                         <apellido_paterno>" & LimpiaCad(rcConsulta2.Fields("ApellidoPaterno")) & "</apellido_paterno>"
                    Print #1, "                         <apellido_materno>" & LimpiaCad(rcConsulta2.Fields("ApellidoMaterno")) & "</apellido_materno>"
                    Print #1, "                         <fecha_nacimiento>" & Format(rcConsulta2.Fields("fecnac"), "yyyyMMdd") & "</fecha_nacimiento>"
                    Print #1, "                         <rfc>" & rcConsulta2.Fields("rfccliente") & "</rfc>"
                    Print #1, "                         <curp>" & IIf(IsNull(rcConsulta2.Fields("clientecurp")), "", rcConsulta2.Fields("clientecurp")) & "</curp>"
                    Print #1, "                   </persona_fisica>"
                    Print #1, "               </tipo_persona>"
                    Print #1, "            </datos_garantia>"
                    
                    Print #1, "            <datos_liquidacion>"
                    Print #1, "                 <fecha_pago>" & Format(rcConsulta3.Fields("vencimiento"), "yyyyMMdd") & "</fecha_pago>"
                    Print #1, "                 <instrumento_monetario>" & Trim(rcConsulta3.Fields("ClaveInstMonetario")) & "</instrumento_monetario>"
                    Print #1, "                 <moneda>" & Trim(rcConsulta3.Fields("ClaveTipoMoneda")) & "</moneda>"
                    Print #1, "                 <monto_operacion>" & rcConsulta3.Fields("DPrestamo") & "</monto_operacion>"
                    Print #1, "            </datos_liquidacion>"
                    Print #1, "       </datos_operacion>"
                   
                    
                    
                    
                    
                    
                    rcConsulta3.MoveNext
                   
               Wend
               
               rcConsulta3.Close
               
                Print #1, "        </operaciones_realizadas>"
                Print #1, "   </detalle_operaciones>"
         End If
         rcConsulta2.MoveNext
    Wend
    'If cero = True Then
    '   Print #1, "<xsd:element name='detalle_operaciones' type='mpc:detalle_operaciones_type' minOccurs='0' maxOccurs='0'></xsd:element>"
    'End If
    Print #1, "   </aviso>"
    Print #1, " </informe>"
    Print #1, "</archivo>"
    Close #1
    On Error Resume Next
    'rcConsulta3.Close
    rcConsulta2.Close
    rcConsultaDatGen.Close
    rcConsultaCero.Close
    
    MsgBox "Se genero el archivo" + Chr(13) + archivoXML '+ " Regs." & Regs
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    
    sFntUnread.Name = "Tahoma"
    sFntUnread.Size = 8
    sFntUnread.Bold = False
    
    
    
    TxtAgnoReporte.text = Year(Date)
    cmbMesReporte.ItemData(0) = 1
    cmbMesReporte.ItemData(1) = 2
    cmbMesReporte.ItemData(2) = 3
    cmbMesReporte.ItemData(3) = 4
    cmbMesReporte.ItemData(4) = 5
    cmbMesReporte.ItemData(5) = 6
    cmbMesReporte.ItemData(6) = 7
    cmbMesReporte.ItemData(7) = 8
    cmbMesReporte.ItemData(8) = 9
    cmbMesReporte.ItemData(9) = 10
    cmbMesReporte.ItemData(10) = 11
    cmbMesReporte.ItemData(11) = 12
    cmbMesReporte.ListIndex = Month(Date) - 1
    
    txtDesde.text = "01/" & Format(cmbMesReporte.ItemData(cmbMesReporte.ListIndex), "00") & "/" & Format(TxtAgnoReporte.text, "0000")
    txtHasta.text = Format(DateAdd("D", -1, Format("01/" & Format(cmbMesReporte.ItemData(cmbMesReporte.ListIndex) + 1, "00") & "/" & Format(TxtAgnoReporte.text, "0000"), "DD/MM/YYYY")), "DD/MM/YYYY")
    
    lblRutaXML.Caption = Trim(Regresa_Valor_BD("RutaArchivosXML"))
    
    'txtDesde.text = Format("01/" & Format(Month(Date), "00") & "/" & Format(Year(Date), "0000"), "DD/MM/YYYY")
    'txtHasta.text = Format(DateAdd("D", -1, Format("01/" & Format(Month(Date) + 1, "00") & "/" & Format(Year(Date), "0000"), "DD/MM/YYYY")), "DD/MM/YYYY")
    crSMinimo = SacaValor("parametros", "ImporteSalario")
    crUdi = SacaValor("parametros", "ImporteUdi")
    txtSalario.text = Regresa_Valor_BD("ImporteVSMPrestamos")
    CrearEncabezados
    CentrarForm Me, frmMDI
End Sub

Sub CrearEncabezados()
    
    With grdClientes
        .AlternateRowBackColor = RGB(216, 240, 254)
        .AddColumn "C1", "Cliente", ecgHdrTextALignLeft, , 250, , , , , , , CCLSortString
        .AddColumn "C2", "Dirección", ecgHdrTextALignLeft, , 100, False, , , , , , CCLSortString
        .AddColumn "C3", "Teléfono", ecgHdrTextALignLeft, , 100, False, , , , , , CCLSortString
        .AddColumn "C4", "Préstamo", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortString
        .AddColumn "C5", "Clave Alerta", ecgHdrTextALignCentre, , 130, , , , , , , CCLSortString
    End With
    
    With grdDetalle
        .AddColumn "C1", "Fecha", ecgHdrTextALignCentre, , 90, , , , , "DD/MMM/YYYY", , CCLSortString
        .AddColumn "C2", "No. Contrato", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "C3", "Préstamo", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortString
        .AddColumn "C4", "Tipo de Garantia", ecgHdrTextALignRight, , 130, , , , , FMoneda, , CCLSortString
        .AddColumn "C5", "Clave Alerta", ecgHdrTextALignCentre, , 90, , , , , , , CCLSortString
        .AddColumn "C6", "Descripción Alerta", ecgHdrTextALignCentre, , 400, , , , , , , CCLSortString
    End With
    
    With grdAvisos
        .AddColumn "C1", "Alerta", ecgHdrTextALignCentre, , 60, , , , , , , CCLSortString
        .AddColumn "C2", "Archivo", ecgHdrTextALignLeft, , 250, , , , , , , CCLSortString
    End With
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rcConsulta = Nothing
End Sub

Private Sub grdClientes_Click(ByVal lRow As Long, ByVal lCol As Long)
    
    If grdClientes.SelectedRow > 0 Then
        
        grdDetalle.Redraw = False
        grdDetalle.Clear
        
        'rcConsulta.Open "SELECT e.ID,e.fecha,e.numcontrato,e.prestamo,e.IdTipoAlerta,e.DescTipoAlerta,m.Clave,m.Descripcion " & _
        '                "FROM empeno e LEFT JOIN mld_prestamos_tipo_alertas AS m ON e.IdTipoAlerta = m.Id " & _
        '                "WHERE e.cancelado=0 AND e.origen=1 AND " & IIf(opCheque.Value, "e.cheque=1 AND ", "") & "e.idcliente=" & grdClientes.CellItemData(grdClientes.SelectedRow, 1) & " AND DATE(e.fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' ORDER BY DATE(e.fecha),e.numcontrato", dbDatos, adOpenForwardOnly, adLockReadOnly
                                       
        rcConsulta.Open "SELECT e.ID,e.fecha,e.numcontrato,e.prestamo,e.IdTipoAlerta,e.DescTipoAlerta,m.Clave,m.Descripcion,e.Serie, " & _
                        "(SELECT t.Descripcion " & _
                        "FROM detallesempeno de INNER JOIN tipo t ON de.Tipo = t.Id and t.idTipoGarantia <> 0 " & _
                        "WHERE  e.ID = de.IdEmpeno limit 1) AS DescripcionGarantia " & _
                        "FROM empeno e LEFT JOIN mld_prestamos_tipo_alertas AS m ON e.IdTipoAlerta = m.Id " & _
                        "WHERE e.IdTipoAlerta=" & grdClientes.CellItemData(grdClientes.SelectedRow, 5) & " AND e.cancelado=0 AND e.origen=1 AND " & IIf(opCheque.Value, "e.cheque=1 AND ", "") & "e.idcliente=" & grdClientes.CellItemData(grdClientes.SelectedRow, 1) & " AND DATE(e.fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' ORDER BY DATE(e.fecha),e.numcontrato", dbDatos, adOpenForwardOnly, adLockReadOnly
                        
        While Not rcConsulta.EOF
            
            grdDetalle.AddRow
            grdDetalle.CellDetails grdDetalle.Rows, 1, rcConsulta!Fecha, DT_CENTER, , , , sFntUnread, , , rcConsulta!ID
            grdDetalle.CellItemData(grdDetalle.Rows, 2) = rcConsulta!ID
            grdDetalle.CellDetails grdDetalle.Rows, 2, rcConsulta!NumContrato, DT_CENTER, , , , sFntUnread
            grdDetalle.CellDetails grdDetalle.Rows, 3, rcConsulta!Prestamo, DT_RIGHT, , , , sFntUnread
            grdDetalle.CellDetails grdDetalle.Rows, 4, IIf(rcConsulta!Serie = SERIE_B, "AUTO", rcConsulta!DescripcionGarantia), DT_CENTER, , , , sFntUnread
            grdDetalle.CellDetails grdDetalle.Rows, 5, rcConsulta!Clave, DT_CENTER, , , , sFntUnread
            grdDetalle.CellItemData(grdDetalle.Rows, 5) = rcConsulta!IDTipoAlerta
            grdDetalle.CellDetails grdDetalle.Rows, 6, UCase(IIf(rcConsulta!DescTipoAlerta = "", rcConsulta!Descripcion, rcConsulta!DescTipoAlerta)), DT_LEFT, , , , sFntUnread
        rcConsulta.MoveNext
        Wend
        rcConsulta.Close
        
        grdDetalle.Redraw = True
    End If
    
End Sub




Private Sub opCheque_Click()
    Frame1.Enabled = False
    txtSalario.text = "0"
    Frame2.Enabled = False
    txtUdis.text = "0"
    Frame3.Enabled = False
    txtPrestamo.text = "0.00"
    grdClientes.Clear
    grdDetalle.Clear
End Sub

Private Sub opPrestamo_Click()
    Frame3.Enabled = True
    txtPrestamo.text = "0.00"
    Frame2.Enabled = False
    txtUdis.text = "0"
    Frame1.Enabled = False
    txtSalario.text = "0"
    grdClientes.Clear
    grdDetalle.Clear
    txtPrestamo.SetFocus
End Sub

Private Sub opSalarios_Click()
    Frame1.Enabled = True
    txtSalario.text = "0": txtSalario.text = Regresa_Valor_BD("ImporteVSMPrestamos")
    Frame2.Enabled = False
    txtUdis.text = "0"
    Frame3.Enabled = False
    txtPrestamo.text = "0.00"
    grdClientes.Clear
    grdDetalle.Clear
    txtSalario.SetFocus
End Sub

Private Sub opUdis_Click()
    Frame2.Enabled = True
    txtUdis.text = "0"
    Frame1.Enabled = False
    txtSalario.text = "0"
    Frame3.Enabled = False
    txtPrestamo.text = "0.00"
    grdClientes.Clear
    grdDetalle.Clear
    txtUdis.SetFocus
End Sub

Private Sub TabFechas_Click(PreviousTab As Integer)

    If PreviousTab = 1 Then
        
    Else
        txtDesde.text = Format("01/" & Format(Month(Date), "00") & "/" & Format(Year(Date), "0000"), "DD/MM/YYYY")
        txtHasta.text = Format(DateAdd("D", -1, DateAdd("M", 6, CDate(txtDesde.text))), "DD/MM/YYYY")
    End If

End Sub

Private Sub TxtAgnoReporte_Change()
    txtDesde.text = "01/" & Format(cmbMesReporte.ItemData(cmbMesReporte.ListIndex), "00") & "/" & Format(TxtAgnoReporte.text, "0000")
    txtHasta.text = Format(DateAdd("D", -1, Format("01/" & Format(cmbMesReporte.ItemData(cmbMesReporte.ListIndex) + 1, "00") & "/" & Format(TxtAgnoReporte.text, "0000"), "DD/MM/YYYY")), "DD/MM/YYYY")
End Sub

Private Sub txtDesde_GotFocus()
    Seleccionar_Texto txtDesde
    Cambiar_Color True, txtDesde
End Sub

Private Sub txtDesde_LostFocus()
    Cambiar_Color False, txtDesde
End Sub

Private Sub txtHasta_GotFocus()
    Seleccionar_Texto txtHasta
    Cambiar_Color True, txtHasta
End Sub

Private Sub txtHasta_LostFocus()
    Cambiar_Color False, txtHasta
End Sub

Private Sub txtPrestamo_GotFocus()
    Seleccionar_Texto txtPrestamo
    Cambiar_Color True, txtPrestamo
End Sub

Private Sub txtPrestamo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrestamo_LostFocus()
    txtPrestamo.text = Format(txtPrestamo.text, FMoneda)
    Cambiar_Color False, txtPrestamo
End Sub

Private Sub txtSalario_GotFocus()
    Seleccionar_Texto txtSalario
    Cambiar_Color True, txtSalario
End Sub

Private Sub txtSalario_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtSalario_LostFocus()
    Cambiar_Color False, txtSalario
End Sub

Private Sub txtUdis_GotFocus()
    Seleccionar_Texto txtUdis
    Cambiar_Color True, txtUdis
End Sub

Private Sub txtUdis_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtUdis_LostFocus()
    Cambiar_Color False, txtUdis
End Sub

Function GetCriterio() As Double
    
    GetCriterio = 0
    If opSalarios.Value Then

        GetCriterio = crSMinimo * Val(txtSalario.text)

    ElseIf opUdis.Value Then

        GetCriterio = crUdi * Val(txtUdis.text)

    ElseIf opPrestamo.Value Then

        GetCriterio = Val(CDbl(txtPrestamo.text))

    End If
    
End Function

Private Sub GenerarArchivoXML()
    
    Dim sqlEncabezadoAuto As String
    Dim archivoXML As String, NombreXML As String
    Dim personaFisica As Boolean
    Dim folioalerta As Integer
    Dim cero As Boolean
   
    Dim vDependencia As String, vClaveIdent As String, vClaveOcup As String, vClavePais As String, vClaveAlerta As String, vClaveMoneda As String, vIDMoneda As Integer
    Dim vTipoOper As Integer, vInstMon As Integer
    Dim Regs As Integer
    Dim vNombreSucursal As String, vRFCSucursal As String, vParam_ActividadVulnerable As String, vParam_GiroMercantil As String, vCPSucursal As String
    Dim vFolioAviso As Integer
    
    Dim SqlEncabezado, sqlCoTitular, sqlTotal, SqlDetalle, sqlCeros1, sqlSucursalDatgen, SqlOperaciones, SqlDetalleAuto As String
    Dim RsEmpenos As New ADODB.Recordset, RsTipoAlerta As New ADODB.Recordset
    Dim bndAbierto As Boolean
    
    On Error GoTo Error
    
    '**************************** VALIDACIONES *******************************
    If opSalarios.Value = False And opUdis.Value = False And opPrestamo.Value = False And opCheque.Value = False Then
        MsgBox "Seleccione una opción de búsqueda...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    ElseIf opSalarios.Value And Trim(txtSalario.text) = "" Then
        MsgBox "Introduzca la cantidad de salarios mínimos...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    ElseIf opUdis.Value And Trim(txtUdis.text) = "" Then
        MsgBox "Introduzca la cantidad de udis...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    ElseIf opPrestamo.Value And Trim(txtPrestamo.text) = "" Then
        MsgBox "Introduzca el importe del préstamo...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    ElseIf Trim(Regresa_Valor_BD("RutaArchivosXML")) = "" Then
        MsgBox "No se ha especificado la Ruta de Generación de Archivos XML...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    ElseIf Dir(Regresa_Valor_BD("RutaArchivosXML"), vbDirectory) = "" Then
        MsgBox "No se encontró la Ruta de Generación de Archivos XML...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    End If
    
    '***************** CARGAR DATOS DEFAULT DE CONFIGURACION ******************
    vClaveAlerta = Trim(SacaValor("mld_prestamos_tipo_alertas", "Clave", " WHERE RegDefault=1"))
    vClavePais = Trim(SacaValor("mld_paises", "Clave", " WHERE RegDefault=1"))
    vClaveOcup = Trim(SacaValor("mld_actividades_economicas", "Clave", " WHERE RegDefault=1"))
    vClaveIdent = Trim(SacaValor("mld_tipo_identificaciones", "Clave", " WHERE RegDefault=1"))
    vDependencia = UCase(Trim(SacaValor("mld_tipo_identificaciones", "Dependencia", " WHERE RegDefault=1")))
    vTipoOper = Val(SacaValor("mld_prestamos_tipo_operacion", "Clave", " WHERE RegDefault=1"))
    vInstMon = Val(SacaValor("mld_instr_monetarios", "Clave", " WHERE RegDefault=1")) 'mld_instr_monetarios
    vClaveMoneda = Trim(SacaValor("mld_tipo_monedas", "Clave", " WHERE MonedaDefault=1"))
    vIDMoneda = Trim(SacaValor("mld_tipo_monedas", "ID", " WHERE MonedaDefault=1"))
    vRFCSucursal = Replace(Trim(SacaValor("sucursales", "RFC", " WHERE Activa=1")), "-", "", 1)
    vNombreSucursal = Trim(SacaValor("sucursales", "NombreComercial", " WHERE Activa=1"))
    vCPSucursal = Trim(SacaValor("sucursales", "Cp", " WHERE Activa=1"))
    vFolioAviso = 0
    bndAbierto = False
    
    sqlSucursalDatgen = "SELECT `mld_actividad_vulnerable`.`Clave` as ActividadVulnerable , `mld_giro_mercantil`.`Descripcion` as GiroMercantil " & _
                        "FROM parametros INNER JOIN mld_actividad_vulnerable ON (parametros.IDActividadVulnerable = mld_actividad_vulnerable.Id) INNER JOIN `basedatos`.`mld_giro_mercantil` ON (`parametros`.`IdTipoGiroMercantil` = `mld_giro_mercantil`.`Id`)"
                        
    rcConsultaDatGen.Open sqlSucursalDatgen, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not rcConsultaDatGen.EOF Then
        vParam_ActividadVulnerable = rcConsultaDatGen!ActividadVulnerable
        vParam_GiroMercantil = rcConsultaDatGen!GiroMercantil
    End If
    rcConsultaDatGen.Close: Set rcConsultaDatGen = Nothing
    
    If vParam_ActividadVulnerable = "" Or vParam_GiroMercantil = "" Then
        MsgBox "Especifique los datos de Configuraciòn del Modulo Antilavado para la Sucursal.", vbCritical, Me.Caption
        Exit Sub
    End If
    
    
    
    '*************************************************************************
    '********************  ESCRITURA DEL ARCHIVO XML  ************************
    '*************************************************************************
    Dim rsEmpeno As New ADODB.Recordset
    Dim RsDatosCliente As New ADODB.Recordset
    Dim RsDatosCoTitular As New ADODB.Recordset
    Dim RsOperaciones As New ADODB.Recordset
    Dim RsDetalle As New ADODB.Recordset
    
    '----- Si Existen Clientes a reportar Aviso ------
    If grdClientes.Rows > 0 Then
            
        '----- BARRER POR TIPO DE ALERTA -----
        RsTipoAlerta.Open "SELECT Id,Clave,Descripcion,ReqDesc FROM mld_prestamos_tipo_alertas ORDER BY Id", dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not RsTipoAlerta.EOF Then
            Do While Not RsTipoAlerta.EOF
                    
                '-------------------------------------------------------------
                bndAbierto = False
                
                '------ BUSCAR LOS CLIENTES CON EL TIPO DE ALERTA ------
                SqlEncabezado = "SELECT `e`.`Id` as IdEmpeno, SUM(`e`.`Prestamo`) as Prestamo,`e`.`Serie`,`e`.`IdCliente`,`e`.`IdCotitular`,if(`mld_prestamos_tipo_alertas`.`Clave` is null,'100',`mld_prestamos_tipo_alertas`.`Clave`) AS ClaveAlerta , `e`.`DescTipoAlerta`,`e`.`IDTipoOperacion`, `e`.`ClaveTipoOperacion`," & _
                                "`e`.`Vencimiento`, if(`mld_instr_monetarios`.`Clave` is null,'1',`mld_instr_monetarios`.`Clave`) as IDInstrumentoMonetario, `mld_instr_monetarios`.`Descripcion` As DescInstrumentoMonetario, `e`.`UltDigitosTarj` As DigitosTarjeta, if(`mld_tipo_monedas`.`Clave` is null, 'MXN',`mld_tipo_monedas`.`Clave`) as IdTipoMoneda," & _
                                "`e`.`Fecha`, `c`.`personafisica`, `c`.`Nombre` as Nombre,  if(`c`.`ApellidoPaterno` is null or `c`.`ApellidoPaterno`='',`c`.`Apellido`,`c`.`ApellidoPaterno`) as ApellidoPaterno ,  `c`.`ApellidoMaterno` as ApellidoMaterno, `c`.`Apellido` as Apellido, `c`.`FecNac` , `c`.`Rfc` as rfccliente, `c`.`Curp` as Clientecurp," & _
                                "if(`mld_paises`.`Clave` is null,'MX',`mld_paises`.`Clave`) as nacionalidad, if(`mld_paises_1`.`Clave` is null,'MX',`mld_paises_1`.`Clave`) as nacimiento, if(`mld_actividades_economicas`.`Clave` is null,'9999999',`mld_actividades_economicas`.`Clave`) as actividadeconomica, if(`mld_tipo_identificaciones`.`Clave` is null,'1',`mld_tipo_identificaciones`.`Clave`) as identificacliente," & _
                                "if(`mld_tipo_identificaciones`.`Dependencia` is null,'INSTITUTO FEDERAL ELECTORAL',`mld_tipo_identificaciones`.`Dependencia`) as Dependencia, `c`.`NumeroIdentificacion`, if(`c`.`RL_Nombre` is null,'',`c`.`RL_Nombre`) As RL_Nombre, if(`c`.`RL_ApellidoPaterno` is null,'',`c`.`RL_ApellidoPaterno`) AS RL_ApellidoPaterno," & _
                                "if(`c`.`RL_ApellidoMaterno` is null,'',`c`.`RL_ApellidoMaterno`) AS RL_ApellidoMaterno, `c`.`RazonSocial`,`c`.`FechaAltaRazonSocial`,`c`.`RL_Rfc`, `c`.`RL_Curp`, `c`.`Colonia`, `c`.`Direccion`, `c`.`NoExterior`, `c`.`NoInterior`, `c`.`CP` as ClienteCP, `c`.`Municipio` as Municipio, `c`.`Estado` as Estado, if(`mld_paises`.`Clave` is null, 'MX',`mld_paises`.`Clave`) as Clave, `c`.`Tel`, `c`.`Email` " & _
                                "FROM `basedatos`.`sucursales` AS s,`basedatos`.`empeno` AS e  LEFT JOIN `basedatos`.`clientes` AS c ON (`e`.`IDCliente` = `c`.`ID`) LEFT JOIN `basedatos`.`mld_paises` ON (`c`.`IdPaisNacionalidad` = `mld_paises`.`Id`) LEFT JOIN `basedatos`.`mld_paises` AS `mld_paises_1` ON (`c`.`IdPaisNacimiento` = `mld_paises_1`.`Id`) LEFT JOIN `basedatos`.`mld_actividades_economicas` ON (`c`.`IdOcupacion` = `mld_actividades_economicas`.`Id`) " & _
                                "LEFT JOIN `basedatos`.`mld_tipo_identificaciones` ON (`c`.`IdTipoIdent` = `mld_tipo_identificaciones`.`Id`) LEFT JOIN `basedatos`.`mld_prestamos_tipo_alertas` ON (`e`.`IdTipoAlerta` = `mld_prestamos_tipo_alertas`.`Id`) " & _
                                "LEFT JOIN `basedatos`.`usuarios` AS u ON (`e`.`IDUsuarioMov` = `u`.`ID`) LEFT JOIN `basedatos`.`mld_instr_monetarios` ON (`e`.`IdInstrumentoMonetario` = `mld_instr_monetarios`.`Id`) LEFT JOIN `basedatos`.`mld_tipo_monedas` ON (`e`.`IdTipoMoneda` = `mld_tipo_monedas`.`Id`) " & _
                                "WHERE e.IdTipoAlerta = " & RsTipoAlerta!ID & " and `e`.`cancelado`=0 AND `e`.`origen`=1 AND DATE(fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' GROUP BY `e`.`IDCliente` HAVING SUM(`e`.`Prestamo`) > " & CDbl(GetCriterio) & " ORDER BY e.IdTipoAlerta"
                    
                rsEmpeno.Open SqlEncabezado, dbDatos, adOpenForwardOnly, adLockOptimistic
                If Not rsEmpeno.EOF Then
                    '-------------------------------------------------------------
                    
                    NombreXML = "": archivoXML = ""
                    
                    'OBTENER EL FOLIO CONSECUTIVO DEL AVISO
                    vFolioAviso = RegresaFolioAvisosXML(Year(txtDesde.text))    'Regresa_Movimiento(True, "FolioAvisosLavado")
                    
                    NombreXML = CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00")) & "_" & CStr(RsTipoAlerta!Clave) & "_F" & vFolioAviso & ".xml"
                    archivoXML = Regresa_Valor_BD("RutaArchivosXML") & "\" & CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00")) & "_" & CStr(RsTipoAlerta!Clave) & "_F" & vFolioAviso & ".xml"
                    Open archivoXML For Output As #1
                    bndAbierto = True
                    Print #1, "<?xml version='1.0' encoding='UTF-8' ?>"
                    'Print #1, "<archivo xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation = 'http://www.uif.shcp.gob.mx/recepcion/mpc mpc.xsd' " & _
                    '          " xmlns='http://www.uif.shcp.gob.mx/recepcion/mpc'> "
                              
                    Print #1, "<archivo xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns='http://www.uif.shcp.gob.mx/recepcion/mpc' xsi:schemaLocation='http://www.uif.shcp.gob.mx/recepcion/mpc mpc.xsd'>"
                   
                    Print #1, " <informe>"
                    Print #1, "  <mes_reportado>" & CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00")) & "</mes_reportado>"
                    Print #1, "  <sujeto_obligado>"
                    Print #1, "     <clave_sujeto_obligado>" & vRFCSucursal & "</clave_sujeto_obligado>"
                    Print #1, "     <clave_actividad>" & vParam_ActividadVulnerable & "</clave_actividad>"
                    Print #1, "  </sujeto_obligado>"
                       
                    Print #1, "  <aviso>"
                    Print #1, "    <referencia_aviso>" & CStr(vFolioAviso) & "</referencia_aviso>"
                    
                    'FNC::: Buscar el archivo previo para obtener el Folio Anterior
                    
                    Print #1, "    <prioridad>1</prioridad>"
                    
                    Print #1, "    <alerta>"
                    Print #1, "       <tipo_alerta>" & RsTipoAlerta!Clave & "</tipo_alerta>"
                    If RsTipoAlerta!ReqDesc = 1 Then
                        Print #1, "       <descripcion_alerta>" & Trim(RsTipoAlerta!Clave) & "</descripcion_alerta>"
                    End If
                    Print #1, "    </alerta>"
                    
                    '-------------------------------------------------------------
                    '************** P E R S O N A S    A V I S O  ****************
                    '-------------------------------------------------------------
                            
                    Do While Not rsEmpeno.EOF
                    
                        Print #1, "   <persona_aviso>"
                        Print #1, "     <tipo_persona>"
                        If rsEmpeno.Fields("personafisica") = 1 Then
                        
                             '*************************** PERSONA FISICA ****************************
                             Print #1, "         <persona_fisica>"
                             Print #1, "             <nombre>" & LimpiaCad(rsEmpeno.Fields("Nombre")) & "</nombre>"
                             Print #1, "             <apellido_paterno>" & LimpiaCad(rsEmpeno.Fields("ApellidoPaterno")) & "</apellido_paterno>"
                             Print #1, "             <apellido_materno>" & LimpiaCad(rsEmpeno.Fields("ApellidoMaterno")) & "</apellido_materno>"
                             Print #1, "             <fecha_nacimiento>" & Format(rsEmpeno.Fields("fecnac"), "yyyyMMdd") & "</fecha_nacimiento>"
                             Print #1, "             <rfc>" & rsEmpeno.Fields("rfccliente") & "</rfc>"
                             Print #1, "             <curp>" & IIf(IsNull(rsEmpeno.Fields("clientecurp")), "", rsEmpeno.Fields("clientecurp")) & "</curp>"
                             Print #1, "             <pais_nacionalidad>" & Trim(rsEmpeno.Fields("nacionalidad")) & "</pais_nacionalidad>"
                             Print #1, "             <pais_nacimiento>" & Trim(rsEmpeno.Fields("nacimiento")) & "</pais_nacimiento>"
                             Print #1, "             <actividad_economica>" & rsEmpeno.Fields("actividadeconomica") & "</actividad_economica>"
                             Print #1, "             <tipo_identificacion>" & Trim(rsEmpeno.Fields("identificacliente")) & "</tipo_identificacion>"
                             Print #1, "             <autoridad_identificacion>" & LimpiaCad(UCase(rsEmpeno.Fields("dependencia"))) & "</autoridad_identificacion>"
                             Print #1, "             <numero_identificacion>" & rsEmpeno.Fields("NumeroIdentificacion") & "</numero_identificacion>"
                             Print #1, "         </persona_fisica>"
                        Else
                            
                            '*************************** PERSONA MORAL ****************************
                             Print #1, "         <persona_moral>"
                             Print #1, "            <denominacion_razon>" & Trim(LimpiaCad(rsEmpeno.Fields("RazonSocial"))) & "</denominacion_razon>"
                             Print #1, "            <fecha_constitucion>" & Format(rsEmpeno.Fields("FechaAltaRazonSocial"), "yyyyMMdd") & "</fecha_constitucion>"
                             Print #1, "            <rfc>" & rsEmpeno.Fields("rfccliente") & "</rfc>"
                             Print #1, "            <pais_nacionalidad>" & Trim(rsEmpeno.Fields("nacionalidad")) & "</pais_nacionalidad>"
                             Print #1, "            <giro_mercantil>" & rsEmpeno.Fields("actividadeconomica") & "</giro_mercantil>"
                             
                             Print #1, "            <representante_apoderado>"
                             Print #1, "                 <nombre>" & LimpiaCad(rsEmpeno.Fields("RL_Nombre")) & "</nombre>"
                             Print #1, "                 <apellido_paterno>" & LimpiaCad(rsEmpeno.Fields("RL_ApellidoPaterno")) & "</apellido_paterno>"
                             Print #1, "                 <apellido_materno>" & LimpiaCad(rsEmpeno.Fields("RL_ApellidoMaterno")) & "</apellido_materno>"
                             Print #1, "                 <fecha_nacimiento>" & Format(rsEmpeno.Fields("FecNAc"), "yyyyMMdd") & "</fecha_nacimiento>"
                             Print #1, "                 <rfc>" & rsEmpeno.Fields("RFC") & "</rfc>"
                             Print #1, "                 <curp>" & rsEmpeno.Fields("Clientecurp") & "</curp>"
                             Print #1, "                 <tipo_identificacion>" & Trim(rsEmpeno.Fields("identificacliente")) & "</tipo_identificacion>"
                             'Print #1, "                 <identificacion_otro></identificacion_otro>"
                             Print #1, "                 <autoridad_identificacion>" & LimpiaCad(UCase(rsEmpeno.Fields("dependencia"))) & "</autoridad_identificacion>"
                             Print #1, "                 <numero_identificacion>" & rsEmpeno.Fields("NumeroIdentificacion") & "</numero_identificacion>"
                             Print #1, "            </representante_apoderado>"
                             Print #1, "         </persona_moral>"
                             
                        End If
                        Print #1, "     </tipo_persona>"
                        
                        '*******************  DOMICILIO  *********************
                        Print #1, "     <tipo_domicilio>"
                        If Trim(rsEmpeno!Nacionalidad) = "MX" Then
                            '**********  NACIONAL ************
                            Print #1, "        <nacional>"
                            Print #1, "           <colonia>" & LimpiaCad(rsEmpeno.Fields("Colonia")) & "</colonia>"
                            Print #1, "           <calle>" & LimpiaCad(rsEmpeno.Fields("Direccion")) & "</calle>"
                            Print #1, "           <numero_exterior>" & rsEmpeno.Fields("NoExterior") & "</numero_exterior>"
                            If rsEmpeno!NoInterior <> "" Then Print #1, "           <numero_interior>" & rsEmpeno.Fields("NoInterior") & "</numero_interior>"
                            Print #1, "           <codigo_postal>" & rsEmpeno.Fields("CLienteCP") & "</codigo_postal>"
                            Print #1, "        </nacional>"
                        Else
                            '*********  EXTRANJERO  *********
                            Print #1, "        <extranjero>"
                            Print #1, "           <pais>" & LimpiaCad(rsEmpeno.Fields("Clave")) & "</pais>"
                            Print #1, "           <estado_provincia>" & LimpiaCad(rsEmpeno.Fields("Estado")) & "</estado_provincia>"
                            Print #1, "           <ciudad_poblacion>" & LimpiaCad(rsEmpeno.Fields("Municipio")) & "</ciudad_poblacion>"
                            
                            Print #1, "           <colonia>" & LimpiaCad(rsEmpeno.Fields("Colonia")) & "</colonia>"
                            Print #1, "           <calle>" & LimpiaCad(rsEmpeno.Fields("Direccion")) & "</calle>"
                            Print #1, "           <numero_exterior>" & rsEmpeno.Fields("NoExterior") & "</numero_exterior>"
                            If rsEmpeno!NoInterior <> "" Then Print #1, "           <numero_interior>" & rsEmpeno.Fields("NoInterior") & "</numero_interior>"
                            Print #1, "           <codigo_postal>" & rsEmpeno.Fields("CLienteCP") & "</codigo_postal>"
                            Print #1, "        </extranjero>"
                        End If
                        Print #1, "     </tipo_domicilio>"
                        
                        '********************  TELEFONO  *********************
                        Print #1, "     <telefono>"
                        Print #1, "         <clave_pais>" & Trim(rsEmpeno.Fields("Clave")) & "</clave_pais>"
                        If rsEmpeno!Tel <> "" Then Print #1, "         <numero_telefono>" & Replace(rsEmpeno.Fields("Tel"), "-", "", 1) & "</numero_telefono>" Else Print #1, "         <numero_telefono>9991111111</numero_telefono>"
'                        If rsEmpeno!Email <> "" Then Print #1, "         <correo_electronico>" & UCase(rsEmpeno.Fields("Email")) & "</correo_electronico>"
                        Print #1, "     </telefono>"
                        Print #1, "   </persona_aviso>"
                        
                        '*****************************************************
                        rsEmpeno.MoveNext
                    Loop
                        
                        
                    '-------------------------------------------------------------
                    '***************** B E N E F I C I A R I O S *****************
                    '-------------------------------------------------------------
                    rsEmpeno.MoveFirst
                    Do While Not rsEmpeno.EOF
                    
                        '**************  DATOS DEL BENEFICIARIO **************
                        If rsEmpeno!IDCotitular <> 0 Then
                        
                            sqlCoTitular = "SELECT `c`.`personafisica`,`c`.`Nombre` as Nombre, if(`c`.`ApellidoPaterno` is null or `c`.`ApellidoPaterno`='',`c`.`Apellido`,`c`.`ApellidoPaterno`) as ApellidoPaterno , `c`.`ApellidoMaterno` as ApellidoMaterno, `c`.`Apellido` as Apellido,`c`.`FecNac` , `c`.`Rfc` as rfccliente, `c`.`Curp` as Clientecurp, " & _
                                           "if(`mld_paises`.`Clave` is null,'MX',`mld_paises`.`Clave`) as nacionalidad, if(`mld_paises_1`.`Clave` is null,'MX',`mld_paises_1`.`Clave`) as nacimiento, if(`mld_actividades_economicas`.`Clave` is null,'9999999',`mld_actividades_economicas`.`Clave`) as actividadeconomica, " & _
                                           "if(`mld_tipo_identificaciones`.`Clave` is null,'1',`mld_tipo_identificaciones`.`Clave`) as identificacliente, if(`mld_tipo_identificaciones`.`Dependencia` is null,'INSTITUTO FEDERAL ELECTORAL',`mld_tipo_identificaciones`.`Dependencia`) as Dependencia, `c`.`NumeroIdentificacion`, " & _
                                           "if(`c`.`RL_Nombre` is null,'',`c`.`RL_Nombre`) As RL_Nombre, if(`c`.`RL_ApellidoPaterno` is null,'',`c`.`RL_ApellidoPaterno`) AS RL_ApellidoPaterno, if(`c`.`RL_ApellidoMaterno` is null,'',`c`.`RL_ApellidoMaterno`) AS RL_ApellidoMaterno, " & _
                                           "`c`.`RazonSocial`,`c`.`FechaAltaRazonSocial`,`c`.`RL_Rfc`, `c`.`RL_Curp`, `c`.`Colonia`, `c`.`Direccion`, `c`.`NoExterior`, `c`.`NoInterior`, `c`.`CP` as ClienteCP, `c`.`Municipio` as Municipio,`c`.`Estado` as Estado, if(`mld_paises`.`Clave` is null, 'MX',`mld_paises`.`Clave`) as Clave, `c`.`Tel`, `c`.`Email` " & _
                                           "FROM `basedatos`.`clientes` AS c LEFT JOIN `basedatos`.`mld_paises` ON (`c`.`IdPaisNacionalidad` = `mld_paises`.`Id`) LEFT JOIN `basedatos`.`mld_paises` AS `mld_paises_1` ON (`c`.`IdPaisNacimiento` = `mld_paises_1`.`Id`) " & _
                                           "LEFT JOIN `basedatos`.`mld_actividades_economicas` ON (`c`.`IdOcupacion` = `mld_actividades_economicas`.`Id`) LEFT JOIN `basedatos`.`mld_tipo_identificaciones` ON (`c`.`IdTipoIdent` = `mld_tipo_identificaciones`.`Id`) " & _
                                           "WHERE c.Id = " & rsEmpeno!IDCotitular
                            
                            RsDatosCoTitular.Open sqlCoTitular, dbDatos, adOpenForwardOnly, adLockOptimistic
                            If Not RsDatosCoTitular.EOF Then
                                Print #1, "   <dueno_beneficiario>"
                                Print #1, "     <tipo_persona>"
                        
                                If RsDatosCoTitular.Fields("personafisica") = 1 Then
                        
                                     '*************************** PERSONA FISICA ****************************
                                     Print #1, "         <persona_fisica>"
                                     Print #1, "             <nombre>" & LimpiaCad(RsDatosCoTitular.Fields("Nombre")) & "</nombre>"
                                     Print #1, "             <apellido_paterno>" & LimpiaCad(RsDatosCoTitular.Fields("ApellidoPaterno")) & "</apellido_paterno>"
                                     Print #1, "             <apellido_materno>" & LimpiaCad(RsDatosCoTitular.Fields("ApellidoMaterno")) & "</apellido_materno>"
                                     Print #1, "             <fecha_nacimiento>" & Format(RsDatosCoTitular.Fields("fecnac"), "yyyyMMdd") & "</fecha_nacimiento>"
                                     Print #1, "             <rfc>" & RsDatosCoTitular.Fields("rfccliente") & "</rfc>"
                                     Print #1, "             <curp>" & IIf(IsNull(RsDatosCoTitular.Fields("clientecurp")), "", RsDatosCoTitular.Fields("clientecurp")) & "</curp>"
                                     Print #1, "             <pais_nacionalidad>" & Trim(RsDatosCoTitular.Fields("nacionalidad")) & "</pais_nacionalidad>"
                                     Print #1, "             <pais_nacimiento>" & Trim(RsDatosCoTitular.Fields("nacimiento")) & "</pais_nacimiento>"
                                     Print #1, "             <actividad_economica>" & RsDatosCoTitular.Fields("actividadeconomica") & "</actividad_economica>"
                                     Print #1, "         </persona_fisica>"
                                Else
                                    
                                    '*************************** PERSONA MORAL ****************************
                                     Print #1, "         <persona_moral>"
                                     Print #1, "            <denominacion_razon>" & Trim(LimpiaCad(RsDatosCoTitular.Fields("RazonSocial"))) & "</denominacion_razon>"
                                     Print #1, "            <fecha_constitucion>" & Format(RsDatosCoTitular.Fields("FechaAltaRazonSocial"), "yyyyMMdd") & "</fecha_constitucion>"
                                     Print #1, "            <rfc>" & RsDatosCoTitular.Fields("rfccliente") & "</rfc>"
                                     Print #1, "            <pais_nacionalidad>" & Trim(RsDatosCoTitular.Fields("nacionalidad")) & "</pais_nacionalidad>"
                                     Print #1, "            <giro_mercantil>" & RsDatosCoTitular.Fields("actividadeconomica") & "</giro_mercantil>"
                                     Print #1, "         </persona_moral>"
                                     
                                End If
                                Print #1, "     </tipo_persona>"
                                
                                '*********  DOMICILIO COTITULAR  *************
                                Print #1, "     <tipo_domicilio>"
                                If Trim(RsDatosCoTitular!Nacionalidad) = "MX" Then
                                    '**********  NACIONAL ************
                                    Print #1, "        <nacional>"
                                    Print #1, "           <colonia>" & LimpiaCad(RsDatosCoTitular.Fields("Colonia")) & "</colonia>"
                                    Print #1, "           <calle>" & LimpiaCad(RsDatosCoTitular.Fields("Direccion")) & "</calle>"
                                    Print #1, "           <numero_exterior>" & RsDatosCoTitular.Fields("NoExterior") & "</numero_exterior>"
                                    If RsDatosCoTitular!NoInterior <> "" Then Print #1, "           <numero_interior>" & RsDatosCoTitular.Fields("NoInterior") & "</numero_interior>"
                                    Print #1, "           <codigo_postal>" & RsDatosCoTitular.Fields("CLienteCP") & "</codigo_postal>"
                                    Print #1, "        </nacional>"
                                Else
                                    '*********  EXTRANJERO  *********
                                    Print #1, "        <extranjero>"
                                    Print #1, "           <pais>" & LimpiaCad(RsDatosCoTitular.Fields("Clave")) & "</pais>"
                                    Print #1, "           <estado_provincia>" & LimpiaCad(RsDatosCoTitular.Fields("Estado")) & "</estado_provincia>"
                                    Print #1, "           <ciudad_poblacion>" & LimpiaCad(RsDatosCoTitular.Fields("Municipio")) & "</ciudad_poblacion>"
                                    
                                    Print #1, "           <colonia>" & LimpiaCad(RsDatosCoTitular.Fields("Colonia")) & "</colonia>"
                                    Print #1, "           <calle>" & LimpiaCad(RsDatosCoTitular.Fields("Direccion")) & "</calle>"
                                    Print #1, "           <numero_exterior>" & RsDatosCoTitular.Fields("NoExterior") & "</numero_exterior>"
                                    If RsDatosCoTitular!NoInterior <> "" Then Print #1, "           <numero_interior>" & RsDatosCoTitular.Fields("NoInterior") & "</numero_interior>"
                                    Print #1, "           <codigo_postal>" & RsDatosCoTitular.Fields("CLienteCP") & "</codigo_postal>"
                                    Print #1, "        </extranjero>"
                                End If
                                Print #1, "     </tipo_domicilio>"
                                
                                '********************  TELEFONO  *********************
                                Print #1, "     <telefono>"
                                Print #1, "         <clave_pais>" & Trim(RsDatosCoTitular.Fields("Clave")) & "</clave_pais>"
                                If RsDatosCoTitular!Tel <> "" Then Print #1, "         <numero_telefono>" & Replace(RsDatosCoTitular.Fields("Tel"), "-", "", 1) & "</numero_telefono>" Else Print #1, "         <numero_telefono>9991111111</numero_telefono>"
'                                If RsDatosCoTitular!Email <> "" Then Print #1, "         <correo_electronico>" & UCase(RsDatosCoTitular.Fields("Email")) & "</correo_electronico>"
                                Print #1, "     </telefono>"
                                Print #1, "   </dueno_beneficiario>"
                                        
                            End If
                            RsDatosCoTitular.Close
                            Set RsDatosCoTitular = Nothing
                    
                        End If
                        '*****************************************************
                        rsEmpeno.MoveNext
                    Loop
                        
                    '-----------------------------------------------------------
                        
                    '************  DETALLE DE LAS OPERACIONES  ***********
                    SqlOperaciones = "SELECT `e`.`Id` as IdEmpeno,`e`.`Prestamo` as Prestamo, X.TPrestamo ,`e`.`Serie`,`e`.`Fecha`,`e`.`IdCliente`, `c`.`Nombre` as Nombre,  if(`c`.`ApellidoPaterno` is null or `c`.`ApellidoPaterno`='',`c`.`Apellido`,`c`.`ApellidoPaterno`) as ApellidoPaterno ,  `c`.`ApellidoMaterno` as ApellidoMaterno, `c`.`Apellido` as Apellido, `c`.`FecNac` , `c`.`Rfc` as rfccliente, `c`.`Curp` as Clientecurp, if(toper.Clave is null,'" & vTipoOper & "',toper.Clave) as ClaveTipoOperacion " & _
                                    "FROM `basedatos`.`empeno` AS e  LEFT JOIN `basedatos`.`clientes` AS c ON (`e`.`IDCliente` = `c`.`ID`) " & _
                                    "LEFT JOIN (SELECT `e`.`IdCliente`, SUM(`e`.`Prestamo`) as TPrestamo FROM `basedatos`.`empeno` AS e  WHERE e.IdTipoAlerta = " & RsTipoAlerta!ID & " and `e`.`cancelado`=0 AND `e`.`origen`=1 and DATE(fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' GROUP BY `e`.`IDCliente` HAVING SUM(`e`.`Prestamo`) > " & CDbl(GetCriterio) & " ORDER BY e.IdTipoAlerta) as X ON X.IdCliente = e.IdCliente " & _
                                    "LEFT JOIN mld_prestamos_tipo_operacion AS toper ON e.IDTipoOperacion = toper.Id WHERE e.IdTipoAlerta = " & RsTipoAlerta!ID & " and `e`.`cancelado`=0 AND `e`.`origen`=1 AND DATE(fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' " & _
                                    "GROUP BY `e`.`IDCliente`,`e`.`ID` HAVING X.TPrestamo > " & CDbl(GetCriterio) & " ORDER BY e.IdTipoAlerta;"
                                    
                    RsOperaciones.Open SqlOperaciones, dbDatos, adOpenForwardOnly, adLockOptimistic
                    If Not RsOperaciones.EOF Then
                        'rsEmpeno.MoveFirst
                        
                        'Abrir Etiqueta de Operaciones
                        Print #1, "   <detalle_operaciones>"
                        Print #1, "        <operaciones_realizadas>"
                                
                        Print #1, "           <datos_operacion>"
                        Print #1, "              <fecha_operacion>" & Format(RsOperaciones.Fields("Fecha"), "yyyyMMdd") & "</fecha_operacion>"
                        Print #1, "              <codigo_postal>" & vCPSucursal & "</codigo_postal>"
                        Print #1, "              <nombre_sucursal>" & LimpiaCad(Trim(vNombreSucursal)) & "</nombre_sucursal>"
                        Print #1, "              <tipo_operacion>" & Trim(RsOperaciones.Fields("ClaveTipoOperacion")) & "</tipo_operacion>"
                                            
                        'Print #1, "              <datos_garantia>"
                                            
                        Do While Not RsOperaciones.EOF
                            
                            If RsOperaciones.Fields("Serie") = SERIE_A Then
                    
                                SqlDetalle = "Select e.IDCliente,e.Fecha,e.Vencimiento,e.FechaMovimiento,e.Origen,e.Destino,de.IDTipoGarantia,de.Prestamo AS DPrestamo,de.Articulo,e.IDTipoMoneda,e.IdTipoAlerta,e.DescTipoAlerta,e.IDInstrumentoMonetario," + Chr(13) & _
                                             "tg.Clave AS ClaveTipoGarantia,if(tm.Clave is null,'" & vClaveMoneda & "',tm.Clave) AS ClaveTipoMoneda,if(im.Clave is null,'" & vInstMon & "',im.Clave) AS ClaveInstMonetario,if(ta.Clave is null,'" & vClaveAlerta & "',ta.Clave) AS ClaveTipoAlerta,if(toper.Clave is null,'" & vTipoOper & "',toper.Clave) as ClaveTipoOperacion " + Chr(13) & _
                                             "From detallesempeno AS de Inner Join empeno AS e ON (de.IDEmpeno = e.ID) Left Join mld_tipo_monedas AS tm ON e.IDTipoMoneda = tm.ID Left Join mld_instr_monetarios AS im ON e.IDInstrumentoMonetario = im.Id Left Join mld_prestamos_tipo_alertas AS ta ON e.IdTipoAlerta = ta.Id " + Chr(13) & _
                                             "Left Join mld_prestamos_tipo_operacion AS toper ON e.IDTipoOperacion = toper.Id Left Join tipo ON de.Tipo = tipo.ID Left Join mld_prestamos_tipo_garantias AS tg ON tipo.IdTipoGarantia = tg.Id " + Chr(13) & _
                                             "WHERE e.idCliente = " & RsOperaciones.Fields("idCliente") & " AND de.IdEmpeno=" & RsOperaciones.Fields("idEmpeno") & ";"
                                          
                            Else
                                
                                SqlDetalle = "select e.IDCliente,e.Fecha,e.Vencimiento,e.FechaMovimiento,e.Origen,e.Destino,de.IDTipoGarantia,e.Prestamo AS DPrestamo,concat(de.MarcaYModelo,' ',cast(año as char)) as articulo_otro,e.IDTipoMoneda,e.IdTipoAlerta,e.DescTipoAlerta,e.IDInstrumentoMonetario," + Chr(13) & _
                                             " " & Val(SacaValor("mld_prestamos_tipo_garantias", "Clave", " WHERE Descripcion LIKE '%Vehículo terrestre%'")) & " AS ClaveTipoGarantia,if(tm.Clave is null,'" & vClaveMoneda & "',tm.Clave) AS ClaveTipoMoneda,if(im.Clave is null,'" & vInstMon & "',im.Clave) AS ClaveInstMonetario,if(ta.Clave is null,'" & vClaveAlerta & "',ta.Clave) AS ClaveTipoAlerta,if(toper.Clave is null,'" & vTipoOper & "',toper.Clave) as ClaveTipoOperacion  from empeno as e " + Chr(13) & _
                                             " left join detallesempenoautos as de on de.IDEmpeno = e.ID Left Join mld_tipo_monedas AS tm ON e.IDTipoMoneda = tm.ID  Left Join mld_instr_monetarios AS im ON e.IDInstrumentoMonetario = im.Id " + Chr(13) & _
                                             " Left Join mld_prestamos_tipo_alertas AS ta ON e.IdTipoAlerta = ta.Id Left Join mld_prestamos_tipo_operacion AS toper ON e.IDTipoOperacion = toper.Id  Left Join tipo ON de.Tipo = tipo.ID " + Chr(13) & _
                                             " Left Join mld_prestamos_tipo_garantias AS tg ON de.IdTipoGarantia = tg.Id WHERE e.idCliente = " & RsOperaciones.Fields("idCliente") & " AND de.IdEmpeno=" & RsOperaciones.Fields("idEmpeno") & " and e.serie=" & SERIE_B & ";"
                            End If
                            RsDetalle.Open SqlDetalle, dbDatos, adOpenForwardOnly, adLockReadOnly
                            If Not RsDetalle.EOF Then
                                
                                '-----------------------------------------
                                Do While Not RsDetalle.EOF
                                    Print #1, "              <datos_garantia>"
                                    Print #1, "                 <tipo_garantia>" & Trim(RsDetalle.Fields("ClaveTipoGarantia")) & "</tipo_garantia>"
                                    Print #1, "              </datos_garantia>"
                                    RsDetalle.MoveNext
                                Loop
                                                            
                            End If
                            RsDetalle.Close
                            Set RsDetalle = Nothing
                            '-------------------------------------------------------------
                            '-------------------------------------------------------------
                            RsOperaciones.MoveNext
                        
                        Loop
                        
                        
                        '********************  DATOS DEL BIEN MUTUO  **********************
                        RsOperaciones.MoveFirst
                        Do While Not RsOperaciones.EOF

                        If RsOperaciones.Fields("Serie") = SERIE_A Then

                                SqlDetalle = "Select e.IDCliente,e.Fecha,e.Vencimiento,e.FechaMovimiento,e.Origen,e.Destino,de.IDTipoGarantia,de.Prestamo AS DPrestamo,de.Articulo,e.IDTipoMoneda,e.IdTipoAlerta,e.DescTipoAlerta,e.IDInstrumentoMonetario," + Chr(13) & _
                                             "tg.Clave AS ClaveTipoGarantia,if(tm.Clave is null,'" & vClaveMoneda & "',tm.Clave) AS ClaveTipoMoneda,if(im.Clave is null,'" & vInstMon & "',im.Clave) AS ClaveInstMonetario,if(ta.Clave is null,'" & vClaveAlerta & "',ta.Clave) AS ClaveTipoAlerta,if(toper.Clave is null,'" & vTipoOper & "',toper.Clave) as ClaveTipoOperacion " + Chr(13) & _
                                             "From detallesempeno AS de Inner Join empeno AS e ON (de.IDEmpeno = e.ID) Left Join mld_tipo_monedas AS tm ON e.IDTipoMoneda = tm.ID Left Join mld_instr_monetarios AS im ON e.IDInstrumentoMonetario = im.Id Left Join mld_prestamos_tipo_alertas AS ta ON e.IdTipoAlerta = ta.Id " + Chr(13) & _
                                             "Left Join mld_prestamos_tipo_operacion AS toper ON e.IDTipoOperacion = toper.Id Left Join tipo ON de.Tipo = tipo.ID Left Join mld_prestamos_tipo_garantias AS tg ON tipo.IdTipoGarantia = tg.Id " + Chr(13) & _
                                             "WHERE e.idCliente = " & RsOperaciones.Fields("idCliente") & " AND de.IdEmpeno=" & RsOperaciones.Fields("idEmpeno") & ";"

                            Else
                                SqlDetalle = "select e.IDCliente,e.Fecha,e.Vencimiento,e.FechaMovimiento,e.Origen,e.Destino,de.IDTipoGarantia,e.Prestamo AS DPrestamo,concat(de.MarcaYModelo,' ',cast(año as char)) as articulo_otro,e.IDTipoMoneda,e.IdTipoAlerta,e.DescTipoAlerta,e.IDInstrumentoMonetario," + Chr(13) & _
                                             " tg.Clave AS ClaveTipoGarantia,if(tm.Clave is null,'" & vClaveMoneda & "',tm.Clave) AS ClaveTipoMoneda,if(im.Clave is null,'" & vInstMon & "',im.Clave) AS ClaveInstMonetario,if(ta.Clave is null,'" & vClaveAlerta & "',ta.Clave) AS ClaveTipoAlerta,if(toper.Clave is null,'" & vTipoOper & "',toper.Clave) as ClaveTipoOperacion  from empeno as e " + Chr(13) & _
                                             " left join detallesempenoautos as de on de.IDEmpeno = e.ID Left Join mld_tipo_monedas AS tm ON e.IDTipoMoneda = tm.ID  Left Join mld_instr_monetarios AS im ON e.IDInstrumentoMonetario = im.Id " + Chr(13) & _
                                             " Left Join mld_prestamos_tipo_alertas AS ta ON e.IdTipoAlerta = ta.Id Left Join mld_prestamos_tipo_operacion AS toper ON e.IDTipoOperacion = toper.Id  Left Join tipo ON de.Tipo = tipo.ID " + Chr(13) & _
                                             " Left Join mld_prestamos_tipo_garantias AS tg ON de.IdTipoGarantia = tg.Id WHERE e.idCliente = " & RsOperaciones.Fields("idCliente") & " AND de.IdEmpeno=" & RsOperaciones.Fields("idEmpeno") & " and e.serie=" & SERIE_B & ";"
                            End If
'                            RsDetalle.Open SqlDetalle, dbDatos, adOpenForwardOnly, adLockReadOnly
'                            If Not RsDetalle.EOF Then
'
'                                RsDetalle.MoveFirst
'                                Do While Not RsDetalle.EOF
'
'                                    Print #1, "                 <datos_bien_mutuo>"
'                                    If RsOperaciones.Fields("serie") = SERIE_B Then
'                                         Print #1, "                        <datos_otro>"
'                                         Print #1, "                              <descripcion_garantia>" & LimpiaCad(Trim("AUTOMOVIL " & (RsDetalle.Fields("articulo_otro")))) & "</descripcion_garantia>"
'                                         Print #1, "                        </datos_otro>"
'                                    Else
'                                         Print #1, "                        <datos_otro>"
'                                         Print #1, "                              <descripcion_garantia>" & LimpiaCad(Trim((RsDetalle.Fields("articulo")))) & "</descripcion_garantia>"
'                                         Print #1, "                        </datos_otro>"
'                                    End If
'                                    Print #1, "                 </datos_bien_mutuo>"
'
'                                    RsDetalle.MoveNext
'                                Loop
'                                '-----------------------------------------
'
'                            End If
'                            RsDetalle.Close
'                            Set RsDetalle = Nothing
                            '-------------------------------------------------------------
                            '-------------------------------------------------------------
                            RsOperaciones.MoveNext

                        Loop

                    End If
                    
                    '**************  DATOS DE LIQUIDACION  ****************
                    RsOperaciones.MoveFirst
                    If Not RsOperaciones.EOF Then
                    
                        Do While Not RsOperaciones.EOF
                        
                        If RsOperaciones.Fields("Serie") = SERIE_A Then
                    
                                SqlDetalle = "Select e.IDCliente,e.Fecha,e.Vencimiento,e.FechaMovimiento,e.Origen,e.Destino,de.IDTipoGarantia,de.Prestamo AS DPrestamo,de.Articulo,e.IDTipoMoneda,e.IdTipoAlerta,e.DescTipoAlerta,e.IDInstrumentoMonetario," + Chr(13) & _
                                             "tg.Clave AS ClaveTipoGarantia,if(tm.Clave is null,'" & vClaveMoneda & "',tm.Clave) AS ClaveTipoMoneda,if(im.Clave is null,'" & vInstMon & "',im.Clave) AS ClaveInstMonetario,if(ta.Clave is null,'" & vClaveAlerta & "',ta.Clave) AS ClaveTipoAlerta,if(toper.Clave is null,'" & vTipoOper & "',toper.Clave) as ClaveTipoOperacion " + Chr(13) & _
                                             "From detallesempeno AS de Inner Join empeno AS e ON (de.IDEmpeno = e.ID) Left Join mld_tipo_monedas AS tm ON e.IDTipoMoneda = tm.ID Left Join mld_instr_monetarios AS im ON e.IDInstrumentoMonetario = im.Id Left Join mld_prestamos_tipo_alertas AS ta ON e.IdTipoAlerta = ta.Id " + Chr(13) & _
                                             "Left Join mld_prestamos_tipo_operacion AS toper ON e.IDTipoOperacion = toper.Id Left Join tipo ON de.Tipo = tipo.ID Left Join mld_prestamos_tipo_garantias AS tg ON tipo.IdTipoGarantia = tg.Id " + Chr(13) & _
                                             "WHERE e.idCliente = " & RsOperaciones.Fields("idCliente") & " AND de.IdEmpeno=" & RsOperaciones.Fields("idEmpeno") & ";"
                                          
                            Else
                                SqlDetalle = "select e.IDCliente,e.Fecha,e.Vencimiento,e.FechaMovimiento,e.Origen,e.Destino,de.IDTipoGarantia,e.Prestamo AS DPrestamo,concat(de.MarcaYModelo,' ',cast(año as char)) as articulo_otro,e.IDTipoMoneda,e.IdTipoAlerta,e.DescTipoAlerta,e.IDInstrumentoMonetario," + Chr(13) & _
                                             " tg.Clave AS ClaveTipoGarantia,if(tm.Clave is null,'" & vClaveMoneda & "',tm.Clave) AS ClaveTipoMoneda,if(im.Clave is null,'" & vInstMon & "',im.Clave) AS ClaveInstMonetario,if(ta.Clave is null,'" & vClaveAlerta & "',ta.Clave) AS ClaveTipoAlerta,if(toper.Clave is null,'" & vTipoOper & "',toper.Clave) as ClaveTipoOperacion from empeno as e " + Chr(13) & _
                                             " left join detallesempenoautos as de on de.IDEmpeno = e.ID Left Join mld_tipo_monedas AS tm ON e.IDTipoMoneda = tm.ID  Left Join mld_instr_monetarios AS im ON e.IDInstrumentoMonetario = im.Id " + Chr(13) & _
                                             " Left Join mld_prestamos_tipo_alertas AS ta ON e.IdTipoAlerta = ta.Id Left Join mld_prestamos_tipo_operacion AS toper ON e.IDTipoOperacion = toper.Id  Left Join tipo ON de.Tipo = tipo.ID " + Chr(13) & _
                                             " Left Join mld_prestamos_tipo_garantias AS tg ON de.IdTipoGarantia = tg.Id WHERE e.idCliente = " & RsOperaciones.Fields("idCliente") & " AND de.IdEmpeno=" & RsOperaciones.Fields("idEmpeno") & " and e.serie=" & SERIE_B & ";"
                            End If
                            RsDetalle.Open SqlDetalle, dbDatos, adOpenForwardOnly, adLockReadOnly
                            If Not RsDetalle.EOF Then
                                
                                RsDetalle.MoveFirst
                                Do While Not RsDetalle.EOF
                                
                                    Print #1, "              <datos_liquidacion>"
                                    Print #1, "                 <fecha_pago>" & Format(IIf(IsNull(RsDetalle.Fields("FechaMovimiento")), RsDetalle.Fields("Fecha"), RsDetalle.Fields("FechaMovimiento")), "yyyyMMdd") & "</fecha_pago>"
                                    Print #1, "                 <instrumento_monetario>" & Trim(RsDetalle.Fields("ClaveInstMonetario")) & "</instrumento_monetario>"
                                    Print #1, "                 <moneda>" & (RsDetalle.Fields("IDTipoMoneda")) & "</moneda>"
                                    Print #1, "                 <monto_operacion>" & Format(RsDetalle.Fields("DPrestamo"), "##########.00") & "</monto_operacion>"
                                    Print #1, "              </datos_liquidacion>"
                                    
                                    RsDetalle.MoveNext
                                Loop
                                
                            End If
                            RsDetalle.Close
                            Set RsDetalle = Nothing
                            '-------------------------------------------------------------
                            '-------------------------------------------------------------
                            RsOperaciones.MoveNext
                        
                        
                        Loop

                    End If
                    
                    
                    RsOperaciones.Close
                    Set RsOperaciones = Nothing
                    '*****************************************************
                            
                    'Cerrar Etiqueta Operaciones
                    Print #1, "           </datos_operacion>"
                    Print #1, "        </operaciones_realizadas>"
                    Print #1, "   </detalle_operaciones>"
                                
                            
                    Print #1, "   </aviso>"
                    Print #1, " </informe>"
                    Print #1, "</archivo>"
                    Close #1
                    
                    
                    'INCREMENTAR EL FOLIO DE AVISO
                    vFolioAviso = RegresaFolioAvisosXML(Year(txtDesde.text), True)
                    
                    
                    RegistraArchivoAviso RsTipoAlerta!Clave, NombreXML
                    
                End If
                rsEmpeno.Close
                Set rsEmpeno = Nothing
                
                RsTipoAlerta.MoveNext
                
            Loop
        End If
        
        MsgBox "Se generaron satisfactoriamente los Avisos XML.", vbInformation, Me.Caption
        
    Else
        
        '---------------------------------------------------------------------
        'INCREMENTAR EL FOLIO DE AVISO
        vFolioAviso = RegresaFolioAvisosXML(Year(txtDesde.text))
        '----------------------  ARCHIVO XML EN CERO  ------------------------
        NombreXML = "": archivoXML = ""
        NombreXML = CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00")) & "_" & CStr(SacaValor("mld_prestamos_tipo_alertas", "Clave", " WHERE RegDefault=1")) & "_F" & vFolioAviso & ".xml"
        archivoXML = Regresa_Valor_BD("RutaArchivosXML") & "\" & CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00")) & "_" & CStr(100) & "_F" & vFolioAviso & ".xml"
        Open archivoXML For Output As #1
        Print #1, "<?xml version='1.0' encoding='UTF-8' ?>"
        'Print #1, "<archivo xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation = 'http://www.uif.shcp.gob.mx/recepcion/mpc mpc.xsd' " & _
        '          " xmlns='http://www.uif.shcp.gob.mx/recepcion/mpc'> "
                  
        Print #1, "<archivo xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns='http://www.uif.shcp.gob.mx/recepcion/mpc' xsi:schemaLocation='http://www.uif.shcp.gob.mx/recepcion/mpc mpc.xsd'>"
                  
        Print #1, " <informe>"
        Print #1, "  <mes_reportado>" & CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00")) & "</mes_reportado>"
        Print #1, "  <sujeto_obligado>"
        Print #1, "     <clave_sujeto_obligado>" & vRFCSucursal & "</clave_sujeto_obligado>"
        Print #1, "     <clave_actividad>" & vParam_ActividadVulnerable & "</clave_actividad>"
        Print #1, "  </sujeto_obligado>"
        Print #1, " </informe>"
        Print #1, "</archivo>"
        Close #1
        '---------------------------------------------------------------------
        'INCREMENTAR EL FOLIO DE AVISO
        vFolioAviso = RegresaFolioAvisosXML(Year(txtDesde.text), True)
        '---------------------------------------------------------------------
        
        RegistraArchivoAviso "CERO", NombreXML
        
        MsgBox "Se genero el archivo" + Chr(13) + archivoXML, vbInformation, Me.Caption
        
    End If
    
    bndAbierto = False
    
Exit Sub

Error:
    Maneja_Error Err
    If bndAbierto = True Then Close #1
    grdAvisos.Clear False
    Exit Sub
End Sub

Private Function FolioArchivoPrevio() As String
        
    'If Dir(Regresa_Valor_BD("RutaArchivosXML") & "\" & CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00")) & "_" & CStr(RsTipoAlerta!Clave) & "_F*.xml", vbArchive) <> "" Then
    '    FolioArchivoPrevio = 1234
    'End If
        
End Function


Private Sub GenerarArchivoXML_PRE()
   Dim sqlEncabezadoAuto As String
    If opSalarios.Value = False And opUdis.Value = False And opPrestamo.Value = False And opCheque.Value = False Then
        
        MsgBox "Seleccione una opción de búsqueda...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    
    ElseIf opSalarios.Value And Trim(txtSalario.text) = "" Then
        
        MsgBox "Introduzca la cantidad de salarios mínimos...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
        
    ElseIf opUdis.Value And Trim(txtUdis.text) = "" Then
        
        MsgBox "Introduzca la cantidad de udis...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    
    ElseIf opPrestamo.Value And Trim(txtPrestamo.text) = "" Then
        
        MsgBox "Introduzca el importe del préstamo...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
        
    ElseIf Trim(Regresa_Valor_BD("RutaArchivosXML")) = "" Then
        MsgBox "No se ha especificado la Ruta de Generación de Archivos XML...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    
    ElseIf Dir(Regresa_Valor_BD("RutaArchivosXML"), vbDirectory) = "" Then
        MsgBox "No se encontró la Ruta de Generación de Archivos XML...!!", vbInformation, "Reporte movimientos atípicos"
        Exit Sub
    
    End If
   ' Dim objDOM As New MSXML2.DOMDocument30
   'Dim objNode As MSXML2.IXMLDOMNode
   'Dim objChildNode As MSXML2.IXMLDOMNode
   'Dim objGrandChildNode As MSXML2.IXMLDOMNode
   'Dim objAttribute As MSXML2.IXMLDOMAttribute
   'Dim objElement As MSXML2.IXMLDOMElement
   
   'Dim Doc As MSXML2.DOMDocument40
   'Dim Nod(2) As MSXML2.IXMLDOMNode
   Dim archivoXML As String
   Dim personaFisica As Boolean
   Dim folioalerta As Integer
   Dim cero As Boolean
   
   Dim vDependencia As String, vClaveIdent As String, vClaveOcup As String, vClavePais As String, vClaveAlerta As String, vClaveMoneda As String
   Dim vTipoOper As Integer, vInstMon As Integer
   Dim Regs As Integer
   Dim vNombreSucursal As String
   
    Dim SqlEncabezado, sqlTotal, SqlDetalle, sqlCeros1, sqlSucursalDatgen, SqlDetalleAuto As String
    SqlEncabezado = ""
    cero = False
    Regs = 0
'        SqlEncabezado = "SELECT `empeno`.`IdCliente`,`clientes`.`personafisica`,`sucursales`.`RFC` as sujetoobligado, `mld_prestamos_tipo_alertas`.`Clave` AS ClaveAlerta , `clientes`.`Nombre`, `clientes`.`ApellidoPaterno` , `clientes`.`ApellidoMaterno`, `clientes`.`FecNac` " & _
'    ", `clientes`.`Rfc` as rfccliente, `clientes`.`Curp` as Clientecurp, `mld_paises`.`Clave` as nacionalidad, `mld_paises_1`.`Clave` as nacimiento, `mld_actividades_economicas`.`Clave` actividadeconomica, `mld_tipo_identificaciones`.`Clave` as identificacliente, `mld_tipo_identificaciones`.`Dependencia`" & _
'    ", `clientes`.`NumeroIdentificacion`, `clientes`.`RL_Nombre`, `clientes`.`RL_ApellidoPaterno`, `clientes`.`RL_ApellidoMaterno`, `clientes`.`RazonSocial`,`clientes`.`FechaAltaRazonSocial`,`clientes`.`RL_Rfc`, `clientes`.`RL_Curp`" & _
'    ", `clientes`.`Colonia`, `clientes`.`Direccion`, `clientes`.`NoExterior`, `clientes`.`NoInterior`, `clientes`.`CP` as ClienteCP, `mld_paises`.`Clave`, `clientes`.`Tel`, `clientes`.`Email`" & _
'    ", `empeno`.`Fecha`, `sucursales`.`Cp` as sucursalCP, `sucursales`.`RazonSocial` as sucursalRazonSocial, `usuarios`.`Nombre`,`sucursales`.`NombreComercial`, `empeno`.`IDTipoOperacion`, `empeno`.`Vencimiento`" & _
'    ", `empeno`.`IDInstrumentoMonetario`, `empeno`.`IDTipoMoneda`, SUM(`empeno`.`Prestamo`) totprestamo, `empeno`.`DescTipoAlerta` from `basedatos`.`sucursales`,`basedatos`.`empeno` LEFT JOIN `basedatos`.`clientes`" & _
'        " ON (`empeno`.`IDCliente` = `clientes`.`ID`) LEFT JOIN `basedatos`.`mld_paises` ON (`clientes`.`IdPaisNacionalidad` = `mld_paises`.`Id`) LEFT JOIN `basedatos`.`mld_paises` AS `mld_paises_1` " & _
'        " ON (`clientes`.`IdPaisNacimiento` = `mld_paises_1`.`Id`) LEFT JOIN `basedatos`.`mld_actividades_economicas` ON (`clientes`.`IdOcupacion` = `mld_actividades_economicas`.`Id`) " & _
'    " LEFT JOIN `basedatos`.`mld_tipo_identificaciones` ON (`clientes`.`IdTipoIdent` = `mld_tipo_identificaciones`.`Id`) LEFT JOIN `basedatos`.`mld_prestamos_tipo_alertas` ON (`empeno`.`IdTipoAlerta` = `mld_prestamos_tipo_alertas`.`Id`) " & _
'    "  LEFT JOIN `basedatos`.`usuarios` ON (`empeno`.`IDUsuarioMov` = `usuarios`.`ID`) where `empeno`.`cancelado`=0 AND `empeno`.`origen`=1  and DATE(fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' GROUP BY `clientes`.`Rfc`;"

    vClaveAlerta = Trim(SacaValor("mld_prestamos_tipo_alertas", "Clave", " WHERE RegDefault=1"))
    vClavePais = Trim(SacaValor("mld_paises", "Clave", " WHERE RegDefault=1"))
    vClaveOcup = Trim(SacaValor("mld_actividades_economicas", "Clave", " WHERE RegDefault=1"))
    vClaveIdent = Trim(SacaValor("mld_tipo_identificaciones", "Clave", " WHERE RegDefault=1"))
    vDependencia = UCase(Trim(SacaValor("mld_tipo_identificaciones", "Dependencia", " WHERE RegDefault=1")))
    vTipoOper = Val(SacaValor("mld_prestamos_tipo_operacion", "Clave", " WHERE RegDefault=1"))
    vInstMon = Val(SacaValor("mld_instr_monetarios", "Clave", " WHERE RegDefault=1")) 'mld_instr_monetarios
    vClaveMoneda = Trim(SacaValor("mld_tipo_monedas", "Clave", " WHERE MonedaDefault=1"))
    
    
    rcConsultaCero.Open "select count(IDempeno)as cuenta from detallesempenoautos", dbDatos, adOpenForwardOnly, adLockReadOnly
    If rcConsultaCero.Fields("cuenta") > 0 Then
        sqlEncabezadoAuto = " union " + Chr(13) + "(SELECT `empeno`.`Id` as IdEmpeno,`empeno`.`IdCliente`,`empeno`.`Serie`,`clientes`.`personafisica`,`sucursales`.`RFC` as sujetoobligado,if(`mld_prestamos_tipo_alertas`.`Clave` is null,'100',`mld_prestamos_tipo_alertas`.`Clave`) AS ClaveAlerta , `clientes`.`Nombre` as Nombre,if(`clientes`.`ApellidoPaterno` is null or `clientes`.`ApellidoPaterno`='',`clientes`.`Apellido`,`clientes`.`ApellidoPaterno`) as ApellidoPaterno , `clientes`.`ApellidoMaterno` as ApellidoMaterno, `clientes`.`Apellido` as Apellido,`clientes`.`FecNac` , `clientes`.`Rfc` as rfccliente, `clientes`.`Curp` as Clientecurp, if(`mld_paises`.`Clave` is null,'MX',`mld_paises`.`Clave`) as nacionalidad,if(`mld_paises_1`.`Clave` is null,'MX',`mld_paises_1`.`Clave`) as nacimiento, if(`mld_actividades_economicas`.`Clave` is null,'9999999',`mld_actividades_economicas`.`Clave`) as actividadeconomica,if(`mld_tipo_identificaciones`.`Clave` is null,'1'," + Chr(13) & _
                            "`mld_tipo_identificaciones`.`Clave`) as identificacliente,if(`mld_tipo_identificaciones`.`Dependencia` is null,'INSTITUTO FEDERAL ELECTORAL',`mld_tipo_identificaciones`.`Dependencia`) as Dependencia, `clientes`.`NumeroIdentificacion`,if(`clientes`.`RL_Nombre` is null,'',`clientes`.`RL_Nombre`) As RL_Nombre,if(`clientes`.`RL_ApellidoPaterno` is null,'',`clientes`.`RL_ApellidoPaterno`) AS RL_ApellidoPaterno, if(`clientes`.`RL_ApellidoMaterno` is null,'',`clientes`.`RL_ApellidoMaterno`) AS RL_ApellidoMaterno, `clientes`.`RazonSocial`,`clientes`.`FechaAltaRazonSocial`,`clientes`.`RL_Rfc`, `clientes`.`RL_Curp`, `clientes`.`Colonia`, `clientes`.`Direccion`, `clientes`.`NoExterior`, `clientes`.`NoInterior`, `clientes`.`CP` as ClienteCP, if(`mld_paises`.`Clave` is null, 'MX',`mld_paises`.`Clave`) as Clave, `clientes`.`Tel`, `clientes`.`Email`, `empeno`.`Fecha`, `sucursales`.`Cp` as sucursalCP, " + Chr(13) & _
                            "`sucursales`.`RazonSocial` as sucursalRazonSocial, `usuarios`.`Nombre` as NombreUsuario,`sucursales`.`NombreComercial`,`empeno`.`IDTipoOperacion`,`empeno`.`Vencimiento`,if(`mld_instr_monetarios`.`Clave` is null,'1',`mld_instr_monetarios`.`Clave`) as IDInstrumentoMonetario,if(`mld_tipo_monedas`.`Clave` is null,'MXN',`mld_tipo_monedas`.`Clave`) as IdTipoMoneda,SUM(`empeno`.`Prestamo`) totprestamo,`empeno`.`DescTipoAlerta` from `basedatos`.`sucursales`,`basedatos`.`empeno` LEFT JOIN `basedatos`.`clientes` ON (`empeno`.`IDCliente` = `clientes`.`ID`) LEFT JOIN `basedatos`.`mld_paises` ON (`clientes`.`IdPaisNacionalidad` = `mld_paises`.`Id`) LEFT JOIN `basedatos`.`mld_paises` AS `mld_paises_1`  ON (`clientes`.`IdPaisNacimiento` = `mld_paises_1`.`Id`) LEFT JOIN `basedatos`.`mld_actividades_economicas` ON (`clientes`.`IdOcupacion` = `mld_actividades_economicas`.`Id`)  " + Chr(13) & _
                            "LEFT JOIN `basedatos`.`mld_tipo_identificaciones` ON (`clientes`.`IdTipoIdent` = `mld_tipo_identificaciones`.`Id`)  LEFT JOIN `basedatos`.`mld_prestamos_tipo_alertas` ON (`empeno`.`IdTipoAlerta` = `mld_prestamos_tipo_alertas`.`Id`) LEFT JOIN `basedatos`.`usuarios` ON (`empeno`.`IDUsuarioMov` = `usuarios`.`ID`) LEFT JOIN `basedatos`.`mld_instr_monetarios` ON (`empeno`.`IdInstrumentoMonetario` = `mld_instr_monetarios`.`Id`) LEFT JOIN `basedatos`.`mld_tipo_monedas` ON (`empeno`.`IdTipoMoneda` = `mld_tipo_monedas`.`Id`) where  empeno.serie=2 and `empeno`.`cancelado`=0 AND `empeno`.`origen`=1 and DATE(fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' GROUP BY `empeno`.`IDCliente`,`empeno`.`ID`)"
 
         vNombreSucursal = Trim(SacaValor("sucursales", "NombreComercial", " WHERE Activa=1"))
        
    Else
        sqlEncabezadoAuto = ""
    End If
    rcConsultaCero.Close
    SqlEncabezado = "(SELECT `empeno`.`Id` as IdEmpeno,`empeno`.`IdCliente`,`empeno`.`Serie`,`clientes`.`personafisica`,`sucursales`.`RFC` as sujetoobligado, if(`mld_prestamos_tipo_alertas`.`Clave` is null,'" & vClaveAlerta & "',`mld_prestamos_tipo_alertas`.`Clave`) AS ClaveAlerta , `clientes`.`Nombre` as Nombre, if(`clientes`.`ApellidoPaterno` is null or `clientes`.`ApellidoPaterno`='',`clientes`.`Apellido`,`clientes`.`ApellidoPaterno`) as ApellidoPaterno , " + Chr(13) & _
                    "`clientes`.`ApellidoMaterno` as ApellidoMaterno, `clientes`.`Apellido` as Apellido,`clientes`.`FecNac` , `clientes`.`Rfc` as rfccliente, `clientes`.`Curp` as Clientecurp, if(`mld_paises`.`Clave` is null,'" & Trim(SacaValor("mld_paises", "Clave", " WHERE RegDefault=1")) & "',`mld_paises`.`Clave`) as nacionalidad, if(`mld_paises_1`.`Clave` is null,'" & vClavePais & "',`mld_paises_1`.`Clave`) as nacimiento, " + Chr(13) & _
                    "if(`mld_actividades_economicas`.`Clave` is null,'" & vClaveOcup & "',`mld_actividades_economicas`.`Clave`) as actividadeconomica, if(`mld_tipo_identificaciones`.`Clave` is null,'" & vClaveIdent & "',`mld_tipo_identificaciones`.`Clave`) as identificacliente, if(`mld_tipo_identificaciones`.`Dependencia` is null,'" & vDependencia & "',`mld_tipo_identificaciones`.`Dependencia`) as Dependencia, `clientes`.`NumeroIdentificacion`, " + Chr(13) & _
                    "if(`clientes`.`RL_Nombre` is null,'',`clientes`.`RL_Nombre`) As RL_Nombre, if(`clientes`.`RL_ApellidoPaterno` is null,'',`clientes`.`RL_ApellidoPaterno`) AS RL_ApellidoPaterno, if(`clientes`.`RL_ApellidoMaterno` is null,'',`clientes`.`RL_ApellidoMaterno`) AS RL_ApellidoMaterno, `clientes`.`RazonSocial`,`clientes`.`FechaAltaRazonSocial`,`clientes`.`RL_Rfc`, `clientes`.`RL_Curp`, " + Chr(13) & _
                    "`clientes`.`Colonia`, `clientes`.`Direccion`, `clientes`.`NoExterior`, `clientes`.`NoInterior`, `clientes`.`CP` as ClienteCP, if(`mld_paises`.`Clave` is null, '" & vClavePais & "',`mld_paises`.`Clave`) as Clave, `clientes`.`Tel`, `clientes`.`Email`, `empeno`.`Fecha`, `sucursales`.`Cp` as sucursalCP," + Chr(13) & _
                    " `sucursales`.`RazonSocial` as sucursalRazonSocial, `usuarios`.`Nombre` as NombreUsuario,`sucursales`.`NombreComercial`, `empeno`.`IDTipoOperacion`, `empeno`.`Vencimiento`, if(`mld_instr_monetarios`.`Clave` is null,'" & vInstMon & "',`mld_instr_monetarios`.`Clave`) as IDInstrumentoMonetario, if(`mld_tipo_monedas`.`Clave` is null, '" & vClaveMoneda & "',`mld_tipo_monedas`.`Clave`) as IdTipoMoneda, SUM(`empeno`.`Prestamo`) totprestamo, `empeno`.`DescTipoAlerta` " + Chr(13) & _
                    "from `basedatos`.`sucursales`,`basedatos`.`empeno` LEFT JOIN `basedatos`.`detallesempeno` ON (`empeno`.`ID` = `detallesempeno`.`IDEmpeno`)  INNER JOIN `basedatos`.`tipo` ON `tipo`.`ID` = `detallesempeno`.`tipo` LEFT JOIN `basedatos`.`clientes`" + Chr(13) & _
                    " ON (`empeno`.`IDCliente` = `clientes`.`ID`) LEFT JOIN `basedatos`.`mld_paises` ON (`clientes`.`IdPaisNacionalidad` = `mld_paises`.`Id`) LEFT JOIN `basedatos`.`mld_paises` AS `mld_paises_1` " + Chr(13) & _
                    " ON (`clientes`.`IdPaisNacimiento` = `mld_paises_1`.`Id`) LEFT JOIN `basedatos`.`mld_actividades_economicas` ON (`clientes`.`IdOcupacion` = `mld_actividades_economicas`.`Id`) " + Chr(13) & _
                    " LEFT JOIN `basedatos`.`mld_tipo_identificaciones` ON (`clientes`.`IdTipoIdent` = `mld_tipo_identificaciones`.`Id`) LEFT JOIN `basedatos`.`mld_prestamos_tipo_alertas` ON (`empeno`.`IdTipoAlerta` = `mld_prestamos_tipo_alertas`.`Id`) " + Chr(13) & _
                    " LEFT JOIN `basedatos`.`usuarios` ON (`empeno`.`IDUsuarioMov` = `usuarios`.`ID`) LEFT JOIN `basedatos`.`mld_instr_monetarios` ON (`empeno`.`IdInstrumentoMonetario` = `mld_instr_monetarios`.`Id`) LEFT JOIN `basedatos`.`mld_tipo_monedas` ON (`empeno`.`IdTipoMoneda` = `mld_tipo_monedas`.`Id`) " + Chr(13) & _
                    "where  tipo.IdTipoGarantia <> 0 and `empeno`.`cancelado`=0 AND `empeno`.`origen`=1  and DATE(fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' GROUP BY `empeno`.`IDCliente`,`empeno`.`ID`)"
                    
                    
    'SqlEncabezado = "SELECT `empeno`.`Id` as IdEmpeno,`empeno`.`IdCliente`,`empeno`.`Serie`,`clientes`.`personafisica`,`sucursales`.`RFC` as sujetoobligado, if(`mld_prestamos_tipo_alertas`.`Clave` is null,'" & vClaveAlerta & "',`mld_prestamos_tipo_alertas`.`Clave`) AS ClaveAlerta , `clientes`.`Nombre` as Nombre, if(`clientes`.`ApellidoPaterno` is null or `clientes`.`ApellidoPaterno`='',`clientes`.`Apellido`,`clientes`.`ApellidoPaterno`) as ApellidoPaterno , " & _
                    "`clientes`.`ApellidoMaterno` as ApellidoMaterno, `clientes`.`Apellido` as Apellido,`clientes`.`FecNac` , `clientes`.`Rfc` as rfccliente, `clientes`.`Curp` as Clientecurp, if(`mld_paises`.`Clave` is null,'" & Trim(SacaValor("mld_paises", "Clave", " WHERE RegDefault=1")) & "',`mld_paises`.`Clave`) as nacionalidad, if(`mld_paises_1`.`Clave` is null,'" & vClavePais & "',`mld_paises_1`.`Clave`) as nacimiento, " & _
                    "if(`mld_actividades_economicas`.`Clave` is null,'" & vClaveOcup & "',`mld_actividades_economicas`.`Clave`) as actividadeconomica, if(`mld_tipo_identificaciones`.`Clave` is null,'" & vClaveIdent & "',`mld_tipo_identificaciones`.`Clave`) as identificacliente, if(`mld_tipo_identificaciones`.`Dependencia` is null,'" & vDependencia & "',`mld_tipo_identificaciones`.`Dependencia`) as Dependencia, `clientes`.`NumeroIdentificacion`, " & _
                    "if(`clientes`.`RL_Nombre` is null,'',`clientes`.`RL_Nombre`) As RL_Nombre, if(`clientes`.`RL_ApellidoPaterno` is null,'',`clientes`.`RL_ApellidoPaterno`) AS RL_ApellidoPaterno, if(`clientes`.`RL_ApellidoMaterno` is null,'',`clientes`.`RL_ApellidoMaterno`) AS RL_ApellidoMaterno, `clientes`.`RazonSocial`,`clientes`.`FechaAltaRazonSocial`,`clientes`.`RL_Rfc`, `clientes`.`RL_Curp`, " & _
                    "`clientes`.`Colonia`, `clientes`.`Direccion`, `clientes`.`NoExterior`, `clientes`.`NoInterior`, `clientes`.`CP` as ClienteCP, if(`mld_paises`.`Clave` is null, '" & vClavePais & "',`mld_paises`.`Clave`) as Clave, `clientes`.`Tel`, `clientes`.`Email`, `empeno`.`Fecha`, `sucursales`.`Cp` as sucursalCP," & _
                    " `sucursales`.`RazonSocial` as sucursalRazonSocial, `usuarios`.`Nombre` as NombreUsuario,`sucursales`.`NombreComercial`, `empeno`.`IDTipoOperacion`, `empeno`.`Vencimiento`, if(`mld_instr_monetarios`.`Clave` is null,'" & vInstMon & "',`mld_instr_monetarios`.`Clave`) as IDInstrumentoMonetario, if(`mld_tipo_monedas`.`Clave` is null, '" & vClaveMoneda & "',`mld_tipo_monedas`.`Clave`) as IdTipoMoneda, SUM(`empeno`.`Prestamo`) totprestamo, `empeno`.`DescTipoAlerta` " & _
                    "from `basedatos`.`sucursales`,`basedatos`.`empeno` LEFT JOIN `basedatos`.`detallesempeno` ON (`empeno`.`ID` = `detallesempeno`.`IDEmpeno`)  INNER JOIN `basedatos`.`tipo` ON `tipo`.`ID` = `detallesempeno`.`tipo` LEFT JOIN `basedatos`.`clientes`" & _
                    " ON (`empeno`.`IDCliente` = `clientes`.`ID`) LEFT JOIN `basedatos`.`mld_paises` ON (`clientes`.`IdPaisNacionalidad` = `mld_paises`.`Id`) LEFT JOIN `basedatos`.`mld_paises` AS `mld_paises_1` " & _
                    " ON (`clientes`.`IdPaisNacimiento` = `mld_paises_1`.`Id`) LEFT JOIN `basedatos`.`mld_actividades_economicas` ON (`clientes`.`IdOcupacion` = `mld_actividades_economicas`.`Id`) " & _
                    " LEFT JOIN `basedatos`.`mld_tipo_identificaciones` ON (`clientes`.`IdTipoIdent` = `mld_tipo_identificaciones`.`Id`) LEFT JOIN `basedatos`.`mld_prestamos_tipo_alertas` ON (`empeno`.`IdTipoAlerta` = `mld_prestamos_tipo_alertas`.`Id`) " & _
                    " LEFT JOIN `basedatos`.`usuarios` ON (`empeno`.`IDUsuarioMov` = `usuarios`.`ID`) LEFT JOIN `basedatos`.`mld_instr_monetarios` ON (`empeno`.`IdInstrumentoMonetario` = `mld_instr_monetarios`.`Id`) LEFT JOIN `basedatos`.`mld_tipo_monedas` ON (`empeno`.`IdTipoMoneda` = `mld_tipo_monedas`.`Id`) " & _
                    "where  tipo.IdTipoGarantia <> 0 and `empeno`.`cancelado`=0 AND `empeno`.`origen`=1  and DATE(fecha)>='" & Format(txtDesde.text, "YYYY/MM/DD") & "' AND DATE(fecha)<='" & Format(txtHasta.text, "YYYY/MM/DD") & "' GROUP BY `empeno`.`IDCliente`;"
    
    
    
    sqlSucursalDatgen = "SELECT `mld_actividad_vulnerable`.`Clave` as ActividadVulnerable , `mld_giro_mercantil`.`Descripcion` as giromercantil From `basedatos`.`parametros` " & _
                        "INNER JOIN `basedatos`.`mld_actividad_vulnerable` ON (`parametros`.`IDActividadVulnerable` = `mld_actividad_vulnerable`.`Id`) INNER JOIN `basedatos`.`mld_giro_mercantil` ON (`parametros`.`IdTipoGiroMercantil` = `mld_giro_mercantil`.`Id`)"
                        
          sqlTotal = SqlEncabezado + Chr(13) + sqlEncabezadoAuto
    rcConsulta2.Open sqlTotal, dbDatos, adOpenForwardOnly, adLockReadOnly
    
    rcConsultaDatGen.Open sqlSucursalDatgen, dbDatos, adOpenForwardOnly, adLockReadOnly
    
    If rcConsulta2.RecordCount <= 0 Then
       cero = True
       sqlCeros1 = "SELECT `RazonSocial`,`Clave`, `NombreComercial`, `RFC`, `Direccion`, `Ciudad`, `Estado`, `Telefono` , `Cp` from `basedatos`.`sucursales`;"
       rcConsultaCero.Open sqlCeros1, dbDatos, adOpenForwardOnly, adLockReadOnly
    End If
    
    'RUTA DE ARCHIVO XML
    archivoXML = Regresa_Valor_BD("RutaArchivosXML") & "\" & CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00")) & ".xml"
   '******************* Se Crea el archivo ****************************
    'Set Doc = New MSXML2.DOMDocument       'Iniciar documento XML y nodo raíz
    'Doc.appendChild Doc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
  ' Set Nod(0) = Doc.createElement("Archivo").
   
   'Set Nod(1) = Doc.createElement("Informe")
   '**********************************************************************
   'Set Nod(2) = Doc.createElement("Mes_reportado")
   'Nod(2).text = CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00"))
   'Nod(1).appendChild Nod(2)
   'Set Nod(1) = Doc.createElement(CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "MM")))
   'Nod(0).appendChild Nod(1)
      
   
   
   'Doc.appendChild Nod(0)                 'Agregar el nodo <Informe> al documento XML
   'Doc.Save App.Path & "\LAVADO.xml"
   Open archivoXML For Output As #1
  
   Print #1, "<?xml version='1.0' encoding='UTF-8' ?>"
    Print #1, "<archivo xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation = 'http://www.uif.shcp.gob.mx/recepcion/mpc mpc.xsd' " & _
            " xmlns='http://www.uif.shcp.gob.mx/recepcion/mpc'> "
    Print #1, " <informe>"
    Print #1, "  <mes_reportado>" & CStr(Year(txtHasta.text) & Strings.Format(Month(txtHasta.text), "00")) & "</mes_reportado>"
    Print #1, "  <sujeto_obligado>"
    If cero = False Then
        Print #1, "     <clave_sujeto_obligado>" & rcConsulta2.Fields("sujetoobligado") & "</clave_sujeto_obligado>"
        Print #1, "     <clave_actividad>" & rcConsultaDatGen.Fields("ActividadVulnerable") & "</clave_actividad>"
    Else
        Print #1, "     <clave_sujeto_obligado>" & rcConsultaCero.Fields("rfc") & "</clave_sujeto_obligado>"
        Print #1, "     <clave_actividad>MPC</clave_actividad>"
    End If
    
    Print #1, "  </sujeto_obligado>"
    
   
    Print #1, "  <aviso>"
    Print #1, "    <referencia_aviso>" & Regresa_Movimiento(True, "FolioAvisosLavado") & "</referencia_aviso>"
    Print #1, "    <prioridad>1</prioridad>"
    Print #1, "    <alerta>"
    
    If cero = False Then
        Print #1, "       <tipo_alerta>" & rcConsulta2.Fields("ClaveAlerta") & "</tipo_alerta:element>"
    Else
        Print #1, "       <tipo_alerta>100</tipo_alerta>"
    
    End If
    Print #1, "    </alerta>"
    
    'DATOS DEL AVISO
    
    
   While Not rcConsulta2.EOF
         If rcConsulta2.Fields("totprestamo") > GetCriterio Then
               
               Regs = Regs + 1
               cero = False
              
              
               Print #1, "   <persona_aviso>"
               Print #1, "     <tipo_persona>"
               '*************************** PERSONA FISICA ****************************
               If rcConsulta2.Fields("personafisica") = 1 Then
                    
                    Print #1, "         <persona_fisica>"
                    Print #1, "             <nombre>" & LimpiaCad(rcConsulta2.Fields("Nombre")) & "</nombre>"
                    Print #1, "             <apellido_paterno>" & LimpiaCad(rcConsulta2.Fields("ApellidoPaterno")) & "</apellido_paterno>"
                    Print #1, "             <apellido_materno>" & LimpiaCad(rcConsulta2.Fields("ApellidoMaterno")) & "</apellido_materno>"
                    Print #1, "             <fecha_nacimiento>" & Format(rcConsulta2.Fields("fecnac"), "yyyyMMdd") & "</fecha_nacimiento>"
                    Print #1, "             <rfc>" & rcConsulta2.Fields("rfccliente") & "</rfc>"
                    Print #1, "             <curp>" & IIf(IsNull(rcConsulta2.Fields("clientecurp")), "", rcConsulta2.Fields("clientecurp")) & "</curp>"
                    Print #1, "             <pais_nacionalidad>" & rcConsulta2.Fields("nacionalidad") & "</pais_nacionalidad>"
                    Print #1, "             <pais_nacimiento>" & rcConsulta2.Fields("nacimiento") & "</pais_nacimiento>"
                    Print #1, "             <actividad_economica>" & rcConsulta2.Fields("actividadeconomica") & "</actividad_economica>"
                    Print #1, "             <tipo_identificacion>" & Trim(rcConsulta2.Fields("identificacliente")) & "</tipo_identificacion>"
                    Print #1, "             <autoridad_identificacion>" & LimpiaCad(UCase(rcConsulta2.Fields("dependencia"))) & "'</autoridad_identificacion>"
                    Print #1, "             <numero_identificacion>" & rcConsulta2.Fields("NumeroIdentificacion") & "</numero_identificacion>"
                    Print #1, "         </persona_fisica>"
               Else
                    Print #1, "         <persona_moral>"
                    Print #1, "            <denominacion_razon>" & Trim(LimpiaCad(rcConsulta2.Fields("RazonSocial"))) & "</denominacion_razon>"
                    Print #1, "            <fecha_constitucion>" & Format(rcConsulta2.Fields("FechaAltaRazonSocial"), "yyyyMMdd") & "</fecha_constitucion>"
                    Print #1, "            <rfc>" & rcConsulta2.Fields("rfccliente") & "</rfc>"
                    Print #1, "            <pais_nacionalidad>" & Trim(rcConsulta2.Fields("nacionalidad")) & "</pais_nacionalidad>"
                    Print #1, "            <giro_mercantil>" & rcConsulta2.Fields("actividadeconomica") & "</giro_mercantil>"
                    
                    Print #1, "            <representante_apoderado>"
                    Print #1, "                 <nombre>" & LimpiaCad(rcConsulta2.Fields("RL_Nombre")) & "</nombre>"
                    Print #1, "                 <apellido_paterno>" & LimpiaCad(rcConsulta2.Fields("RL_ApellidoPaterno")) & "</apellido_paterno>"
                    Print #1, "                 <apellido_materno>" & LimpiaCad(rcConsulta2.Fields("RL_ApellidoMaterno")) & "</apellido_materno>"
                    Print #1, "                 <fecha_nacimiento>" & Format(rcConsulta2.Fields("FecNAc"), "yyyyMMdd") & "</fecha_nacimiento>"
                    Print #1, "                 <rfc>" & rcConsulta2.Fields("RFC") & "</rfc>"
                    Print #1, "                 <curp>" & rcConsulta2.Fields("Clientecurp") & "</curp>"
                    Print #1, "                 <tipo_identificacion>" & Trim(rcConsulta2.Fields("identificacliente")) & "</tipo_identificacion>"
                    'Print #1, "                 <identificacion_otro></identificacion_otro>"
                    Print #1, "                 <autoridad_identificacion>" & LimpiaCad(UCase(rcConsulta2.Fields("dependencia"))) & "</autoridad_identificacion>"
                    Print #1, "                 <numero_identificacion>" & rcConsulta2.Fields("NumeroIdentificacion") & "</numero_identificacion>"
                    Print #1, "            </representante_apoderado>"
                    Print #1, "         </persona_moral>"
                    
                    
               End If
              
              Print #1, "     </tipo_persona>"
              Print #1, "     <tipo_domicilio>"
              Print #1, "        <nacional>"
              Print #1, "           <colonia>" & LimpiaCad(rcConsulta2.Fields("Colonia")) & "</colonia>"
              Print #1, "           <calle>" & LimpiaCad(rcConsulta2.Fields("Direccion")) & "</calle>"
              Print #1, "           <numero_exterior>" & rcConsulta2.Fields("NoExterior") & "</numero_exterior>"
              Print #1, "           <numero_interior>" & rcConsulta2.Fields("NoInterior") & "</numero_interior>"
              Print #1, "           <codigo_postal>" & rcConsulta2.Fields("CLienteCP") & "</codigo_postal>"
              Print #1, "        </nacional>"
              Print #1, "     </tipo_domicilio>"
              Print #1, "     <telefono>"
              Print #1, "         <clave_pais>MX</clave_pais>"
              Print #1, "         <numero_telefono>" & rcConsulta2.Fields("Tel") & "</numero_telefono>"
              Print #1, "         <correo_electronico>" & rcConsulta2.Fields("Email") & "</correo_electronico>"
              Print #1, "     </telefono>"
              Print #1, "   </persona_aviso>"
                   
                   
              
               
               '************************************ OPERACIONES *****************************
                Print #1, "   <detalle_operaciones>"
                Print #1, "        <operaciones_realizadas>"
               
              If rcConsulta2.Fields("serie") = 1 Then
                
                   SqlDetalle = "Select e.IDCliente,e.Fecha,e.Vencimiento,de.IDTipoGarantia,de.Prestamo AS DPrestamo,de.Articulo,e.IDTipoMoneda,e.IdTipoAlerta,e.DescTipoAlerta,e.IDInstrumentoMonetario," + Chr(13) & _
                            "tg.Clave AS ClaveTipoGarantia,if(tm.Clave is null,'" & vClaveMoneda & "',tm.Clave) AS ClaveTipoMoneda,if(im.Clave is null,'" & vInstMon & "',im.Clave) AS ClaveInstMonetario,if(ta.Clave is null,'" & vClaveAlerta & "',ta.Clave) AS ClaveTipoAlerta,if(toper.Clave is null,'" & vTipoOper & "',toper.Clave) as ClaveTipoOperacion " + Chr(13) & _
                            "From detallesempeno AS de Inner Join empeno AS e ON (de.IDEmpeno = e.ID) Left Join mld_tipo_monedas AS tm ON e.IDTipoMoneda = tm.ID Left Join mld_instr_monetarios AS im ON e.IDInstrumentoMonetario = im.Id Left Join mld_prestamos_tipo_alertas AS ta ON e.IdTipoAlerta = ta.Id " + Chr(13) & _
                            "Left Join mld_prestamos_tipo_operacion AS toper ON e.IDTipoOperacion = toper.Id Left Join tipo ON de.Tipo = tipo.ID Left Join mld_prestamos_tipo_garantias AS tg ON tipo.IdTipoGarantia = tg.Id " + Chr(13) & _
                            "WHERE e.idCliente = " & rcConsulta2.Fields("idCliente") & " AND de.IdEmpeno=" & rcConsulta2.Fields("idEmpeno") & ";"
                            
              Else
                SqlDetalle = "select e.IDCliente,e.Fecha,e.Vencimiento,de.IDTipoGarantia,e.Prestamo AS DPrestamo,concat(de.MarcaYModelo,' ',cast(año as char)) as articulo_otro,e.IDTipoMoneda,e.IdTipoAlerta,e.DescTipoAlerta,e.IDInstrumentoMonetario," + Chr(13) & _
                            " tg.Clave AS ClaveTipoGarantia,if(tm.Clave is null,'" & vClaveMoneda & "',tm.Clave) AS ClaveTipoMoneda,if(im.Clave is null,'" & vInstMon & "',im.Clave) AS ClaveInstMonetario,if(ta.Clave is null,'" & vClaveAlerta & "',ta.Clave) AS ClaveTipoAlerta,if(toper.Clave is null,'" & vTipoOper & "',toper.Clave) as ClaveTipoOperacion  from empeno as e " + Chr(13) & _
                            " left join detallesempenoautos as de on de.IDEmpeno = e.ID Left Join mld_tipo_monedas AS tm ON e.IDTipoMoneda = tm.ID  Left Join mld_instr_monetarios AS im ON e.IDInstrumentoMonetario = im.Id " + Chr(13) & _
                            " Left Join mld_prestamos_tipo_alertas AS ta ON e.IdTipoAlerta = ta.Id Left Join mld_prestamos_tipo_operacion AS toper ON e.IDTipoOperacion = toper.Id  Left Join tipo ON de.Tipo = tipo.ID " + Chr(13) & _
                            " Left Join mld_prestamos_tipo_garantias AS tg ON tipo.IdTipoGarantia = tg.Id WHERE e.idCliente = " & rcConsulta2.Fields("idCliente") & " AND de.IdEmpeno=" & rcConsulta2.Fields("idEmpeno") & " and e.serie=2;"
              End If
               rcConsulta3.Open SqlDetalle, dbDatos, adOpenForwardOnly, adLockReadOnly
               
               While Not rcConsulta3.EOF
                    Print #1, "        <datos_operacion>"
                    Print #1, "           <fecha_operacion>" & Format(rcConsulta3.Fields("Fecha"), "yyyyMMdd") & "</fecha_operacion>"
                    Print #1, "           <codigo_postal>" & rcConsulta2.Fields("clienteCP") & "</codigo_postal>"
                    Print #1, "           <nombre_sucursal>" & LimpiaCad(rcConsulta2.Fields("NombreComercial")) & "</nombre_sucursal>"
                    Print #1, "           <tipo_operacion>" & Trim(rcConsulta3.Fields("ClaveTipoOperacion")) & "</tipo_operacion>"
                    
                    Print #1, "           <datos_garantia>"
                    Print #1, "               <tipo_garantia>" & Trim(rcConsulta3.Fields("ClaveTipoGarantia")) & "</tipo_garantia>"
                    Print #1, "               <datos_bien_mutuo>"
                    If rcConsulta2.Fields("serie") = 2 Then
                         Print #1, "                        <datos_otro>"
                    
                         Print #1, "                              <descripcion_garantia>" & "Auto " & (rcConsulta3.Fields("articulo_otro")) & "</descripcion_garantia>"
                         Print #1, "                        </datos_otro>"
                    End If
                    Print #1, "               </datos_bien_mutuo>"
                    Print #1, "               <tipo_persona>"
                    Print #1, "                   <persona_fisica>"
                    Print #1, "                         <nombre>" & LimpiaCad(rcConsulta2.Fields("Nombre")) & "</nombre>"
                    Print #1, "                         <apellido_paterno>" & LimpiaCad(rcConsulta2.Fields("ApellidoPaterno")) & "</apellido_paterno>"
                    Print #1, "                         <apellido_materno>" & LimpiaCad(rcConsulta2.Fields("ApellidoMaterno")) & "</apellido_materno>"
                    Print #1, "                         <fecha_nacimiento>" & Format(rcConsulta2.Fields("fecnac"), "yyyyMMdd") & "</fecha_nacimiento>"
                    Print #1, "                         <rfc>" & rcConsulta2.Fields("rfccliente") & "</rfc>"
                    Print #1, "                         <curp>" & IIf(IsNull(rcConsulta2.Fields("clientecurp")), "", rcConsulta2.Fields("clientecurp")) & "</curp>"
                    Print #1, "                   </persona_fisica>"
                    Print #1, "               </tipo_persona>"
                    Print #1, "            </datos_garantia>"
                    
                    Print #1, "            <datos_liquidacion>"
                    Print #1, "                 <fecha_pago>" & Format(rcConsulta3.Fields("vencimiento"), "yyyyMMdd") & "</fecha_pago>"
                    Print #1, "                 <instrumento_monetario>" & Trim(rcConsulta3.Fields("ClaveInstMonetario")) & "</instrumento_monetario>"
                    Print #1, "                 <moneda>" & Trim(rcConsulta3.Fields("ClaveTipoMoneda")) & "</moneda>"
                    Print #1, "                 <monto_operacion>" & rcConsulta3.Fields("DPrestamo") & "</monto_operacion>"
                    Print #1, "            </datos_liquidacion>"
                    Print #1, "       </datos_operacion>"
                   
                    
                    
                    
                    
                    
                    rcConsulta3.MoveNext
                   
               Wend
               
               rcConsulta3.Close
               
                Print #1, "        </operaciones_realizadas>"
                Print #1, "   </detalle_operaciones>"
         End If
         rcConsulta2.MoveNext
    Wend
    'If cero = True Then
    '   Print #1, "<xsd:element name='detalle_operaciones' type='mpc:detalle_operaciones_type' minOccurs='0' maxOccurs='0'></xsd:element>"
    'End If
    Print #1, "   </aviso>"
    Print #1, " </informe>"
    Print #1, "</archivo>"
    Close #1
    On Error Resume Next
    'rcConsulta3.Close
    rcConsulta2.Close
    rcConsultaDatGen.Close
    rcConsultaCero.Close
    
    MsgBox "Se genero el archivo" + Chr(13) + archivoXML '+ " Regs." & Regs

End Sub


Private Function RegresaFolioAvisosXML(ByVal Agno As Integer, Optional Foliar As Boolean = False) As Integer
    Dim Rs As New ADODB.Recordset
    Dim Movimiento As Long
    
    Rs.Open "SELECT * FROM mld_folio_avisos WHERE AnoAviso=" & Agno, dbDatos, adOpenKeyset, adLockOptimistic
    If Not Rs.EOF Then
        Movimiento = Rs!FolioAviso
        If Foliar Then
            Movimiento = Movimiento + 1
            dbDatos.Execute "UPDATE mld_folio_avisos SET FolioAviso= " & Movimiento & " WHERE AnoAviso=" & Agno & ";"
        End If
        RegresaFolioAvisosXML = Movimiento
    Else
        Movimiento = 1
        dbDatos.Execute "INSERT INTO mld_folio_avisos (AnoAviso,FolioAviso) VALUES (" & Agno & "," & Movimiento & ");"
        RegresaFolioAvisosXML = Movimiento
    End If
    Rs.Close
    Set Rs = Nothing
    
End Function


Private Sub RegistraArchivoAviso(ByVal Clave As String, ByVal Archivo As String)

    With grdAvisos
        .AddRow
        .CellText(.Rows, 1) = Trim(Clave)
        .CellTextAlign(.Rows, 1) = DT_CENTER
        .CellText(.Rows, 2) = Trim(Archivo)
    End With
    grdAvisos.Redraw = True

End Sub
