VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmAutorizaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorizaciones"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutorizaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   7020
   Begin vbAcceleratorGrid6.vbalGrid grdAutorizaciones 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   6324
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      HighlightBackColor=   -2147483633
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
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   5850
      TabIndex        =   1
      Top             =   3675
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
      Picture         =   "frmAutorizaciones.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3675
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Imprimir"
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
      Picture         =   "frmAutorizaciones.frx":055E
   End
End
Attribute VB_Name = "frmAutorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fecha1 As Date, Fecha2 As Date

Public Property Let FechaIni(Valor As Date)
    Fecha1 = Valor
End Property

Public Property Get FechaIni() As Date
    FechaIni = Fecha1
End Property

Public Property Let FechaFin(Valor As Date)
    Fecha2 = Valor
End Property

Public Property Get FechaFin() As Date
    FechaFin = Fecha2
End Property

Private Sub cmdImprimir_Click()
    With frmMDI.Cr
    
        .Reset
        .DiscardSavedData = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\Autorizaciones.rpt"
        .SelectionFormula = "{autorizaciones.Fecha}>=date('" & Format(FechaIni, "YYYY/MM/DD") & "')" & " AND {autorizaciones.Fecha}<=date('" & Format(FechaFin, "YYYY/MM/DD") & "')"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .WindowShowPrintSetupBtn = True
        .WindowTitle = "Reporte Autorizaciones"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    
    End With
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Function ver(Fecha1 As Date, Fecha2 As Date)
    FechaIni = Fecha1
    FechaFin = Fecha2
    CreaEncabezados
    CargaDatos
End Function

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    CentrarForm Me, frmMDI
End Sub

Function CreaEncabezados()
    With grdAutorizaciones
        .ImageList = frmMDI.img
        .AddColumn "K1", "Código Autorización", ecgHdrTextALignLeft, , 170, False, , , , , , CCLSortString
        .AddColumn "K2", "Fecha", ecgHdrTextALignLeft, , 140, , , , , "DD/MMM/YYYY HH:MM:SSAM/PM", , CCLSortDate
        .AddColumn "K3", "Tipo Aut.", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
        .AddColumn "K4", "Contrato", ecgHdrTextALignCentre, , 60, , , , , , , CCLSortNumeric
        .AddColumn "K5", "Cliente", ecgHdrTextALignLeft, , 170, , , , , , , CCLSortString
    End With
End Function

Sub CargaDatos()
Dim rcConsulta As New ADODB.Recordset

On Error GoTo error

    rcConsulta.Open "SELECT autorizaciones.Status,autorizaciones.ID,autorizaciones.Fecha,autorizaciones.Opcion,autorizaciones.Codigo,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente,empeno.NumContrato FROM autorizaciones INNER JOIN empeno ON autorizaciones.ID=empeno.IDAutorizacion LEFT JOIN clientes ON empeno.IDCliente=clientes.ID WHERE DATE_FORMAT(autorizaciones.Fecha,'%Y%/%m%/%d') BETWEEN '" & Format(FechaIni, "YYYY/MM/DD") & "' AND '" & Format(FechaFin, "YYYY/MM/DD") & "' ORDER BY autorizaciones.Fecha", dbDatos, adOpenForwardOnly, adLockReadOnly
    With grdAutorizaciones
        
        .Redraw = False
        While Not rcConsulta.EOF
        
            .AddRow
            .CellText(.Rows, 1) = rcConsulta!Codigo
            .CellIcon(.Rows, 1) = IIf(rcConsulta!Status = 1, frmMDI.img.ItemIndex(1), -1)
            .CellItemData(.Rows, 1) = rcConsulta!ID
            .CellTextAlign(.Rows, 1) = DT_RIGHT
            .CellText(.Rows, 2) = rcConsulta!Fecha
            .CellText(.Rows, 3) = IIf(rcConsulta!Opcion = 0, "Limite 1", "Límite 2")
            .CellText(.Rows, 4) = rcConsulta!NumContrato
            .CellTextAlign(.Rows, 4) = DT_CENTER
            .CellText(.Rows, 5) = rcConsulta!Cliente
        rcConsulta.MoveNext
        Wend
        .Redraw = True
    End With
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Private Sub grdAutorizaciones_Click(ByVal lRow As Long, ByVal lCol As Long)
    If lCol = 1 And grdAutorizaciones.SelectedRow > 0 Then
        If grdAutorizaciones.CellIcon(lRow, lCol) = -1 Then
            grdAutorizaciones.CellIcon(lRow, lCol) = frmMDI.img.ItemIndex(1)
            dbDatos.Execute "UPDATE autorizaciones SET Status=1 WHERE ID=" & grdAutorizaciones.CellItemData(lRow, 1)
        Else
            grdAutorizaciones.CellIcon(lRow, lCol) = -1
            dbDatos.Execute "UPDATE autorizaciones SET Status=0 WHERE ID=" & grdAutorizaciones.CellItemData(lRow, 1)
        End If
    End If
End Sub
