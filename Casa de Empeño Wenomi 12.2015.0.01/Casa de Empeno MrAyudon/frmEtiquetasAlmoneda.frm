VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{0BFA85A1-F9B8-11CF-8939-444553540000}#1.0#0"; "barcode.ocx"
Begin VB.Form frmEtiquetasAlmoneda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Etiquetas Inventario"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEtiquetasAlmoneda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   13125
   Begin VB.Frame Frame1 
      Caption         =   "Tipo Entrada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   90
      TabIndex        =   4
      Top             =   6780
      Width           =   3975
      Begin VB.ComboBox cmbTipoEntrada 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "frmEtiquetasAlmoneda.frx":000C
         Left            =   90
         List            =   "frmEtiquetasAlmoneda.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   3645
      End
   End
   Begin vbAcceleratorGrid6.vbalGrid grdArticulos 
      Height          =   6780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   11959
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
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
   Begin BARCODELib.Barcode bcCodigo 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   4710
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
      _ExtentY        =   1931
      _StockProps     =   25
      Text            =   "12345678901212"
      TypeName        =   "EAN 13"
      Text            =   "12345678901212"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Borderwidth     =   5
      Borderheight    =   8
      NotchHeightInPercent=   15
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCancelar 
      Height          =   375
      Left            =   9585
      TabIndex        =   3
      Top             =   7035
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Cancelar"
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
      Picture         =   "frmEtiquetasAlmoneda.frx":0039
   End
   Begin DevPowerFlatBttn.FlatBttn cmdTodos 
      Height          =   375
      Left            =   8460
      TabIndex        =   5
      Top             =   7050
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "      &Todos"
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
      Picture         =   "frmEtiquetasAlmoneda.frx":00AF
      PictureDisabled =   "frmEtiquetasAlmoneda.frx":0419
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   11940
      TabIndex        =   6
      Top             =   7020
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmEtiquetasAlmoneda.frx":0573
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   10755
      TabIndex        =   7
      Top             =   7035
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
      Picture         =   "frmEtiquetasAlmoneda.frx":0AC5
   End
End
Attribute VB_Name = "frmEtiquetasAlmoneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fl() As cFlatControl
Dim FechaInicial As String, FechaFinal As String

Public Property Let FechaIni(Fecha As String)
    FechaInicial = Fecha
End Property

Public Property Get FechaIni() As String
    FechaIni = FechaInicial
End Property

Public Property Let FechaFin(Fecha As String)
    FechaFinal = Fecha
End Property

Public Property Get FechaFin() As String
    FechaFin = FechaFinal
End Property

Sub CreaEncabezado()

    With grdArticulos
        .ImageList = frmMDI.img
        .AddColumn "K1", "Código", ecgHdrTextALignLeft, , 110, , , , , , , CCLSortString
        .AddColumn "K2", "Descripción", ecgHdrTextALignLeft, , 290, , , , , , , CCLSortString
        .AddColumn "K3", "Existencia", ecgHdrTextALignRight, , 60, , , , , , , CCLSortNumeric
        .AddColumn "K4", "Peso", ecgHdrTextALignRight, , 60, , , , , "###,###0.000", , CCLSortNumeric
        .AddColumn "K5", "Costo", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Precio", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Precio V.", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K8", "Tipo", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
    End With

End Sub

Sub CargaDatos()
Dim rcBusqueda As New ADODB.Recordset
Dim Iva As Double, TipoEntrada As Integer

On Error GoTo error
    
    Iva = Regresa_Valor_BD("IVAVentas") / 100
    
    Select Case cmbTipoEntrada.ListIndex
    Case 0
        
        TipoEntrada = ENTRADADOTACION
    Case 1
        
        TipoEntrada = ENTRADAALMONEDA
    Case 2
        
        TipoEntrada = ENTRADACOMPRA
    End Select
    
    
    rcBusqueda.Open "SELECT d.codigo,d.Descripcion,d.Cantidad,d.Peso," _
                    & "d.Costo,d.Precio AS PrecioVenta,d.PrecioVitrina,d.Kilates,tipo.Descripcion AS TipoDescripcion FROM detallesentradainventario d INNER JOIN tipo ON d.Tipo=tipo.ID INNER JOIN entradainventario ON d.IDEntrada=entradainventario.ID " _
                    & "WHERE DATE_FORMAT(entradainventario.Fecha,'%Y%/%m%/%d')>='" & Format(CDate(Me.FechaIni), "YYYY/MM/DD") & "' AND DATE_FORMAT(entradainventario.Fecha,'%Y%/%m%/%d')<='" & Format(CDate(Me.FechaFin), "YYYY/MM/DD") & "' AND d.Cantidad>0 AND d.TipoEntrada=" & TipoEntrada, dbDatos, adOpenForwardOnly, adLockOptimistic
    grdArticulos.Clear
    If Not rcBusqueda.BOF And Not rcBusqueda.EOF Then
        
        rcBusqueda.MoveFirst
        With grdArticulos
            
            .Redraw = False
            While Not rcBusqueda.EOF
                .AddRow
                .CellIcon(.Rows, 1) = frmMDI.img.ItemIndex(2)
                .CellText(.Rows, 1) = rcBusqueda!Codigo
                .CellText(.Rows, 2) = rcBusqueda!Descripcion
                .CellText(.Rows, 3) = rcBusqueda!Cantidad
                .CellTextAlign(.Rows, 3) = DT_RIGHT
                .CellText(.Rows, 4) = rcBusqueda!Peso
                .CellTextAlign(.Rows, 4) = DT_RIGHT
                .CellText(.Rows, 5) = rcBusqueda!Costo
                .CellTextAlign(.Rows, 5) = DT_RIGHT
                .CellText(.Rows, 6) = rcBusqueda!PrecioVenta
                .CellItemData(.Rows, 6) = rcBusqueda!Kilates
                .CellTextAlign(.Rows, 6) = DT_RIGHT
                .CellText(.Rows, 7) = rcBusqueda!PrecioVitrina * (1 + Iva)
                .CellTextAlign(.Rows, 7) = DT_RIGHT
                .CellText(.Rows, 8) = rcBusqueda!TipoDescripcion
            rcBusqueda.MoveNext
            Wend
            .Redraw = True
        rcBusqueda.Close
        Set rcBusqueda = Nothing
        End With
    
    End If
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcBusqueda = Nothing
End Sub

Private Sub cmbTipoEntrada_Click()
    CargaDatos
End Sub

Private Sub cmbTipoEntrada_GotFocus()
    Cambiar_Color True, cmbTipoEntrada
End Sub

Private Sub cmbTipoEntrada_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoEntrada_LostFocus()
    Cambiar_Color False, cmbTipoEntrada
End Sub

Private Sub cmdCancelar_Click()
Dim i As Long

    For i = 1 To grdArticulos.Rows
        grdArticulos.CellIcon(i, 1) = frmMDI.img.ItemIndex(2)
    Next i
End Sub

Private Sub cmdImprimir_Click()
Dim i As Long

    If MsgBox("Desea imprimir las etiquetas de las prendas seleccionadas ??", vbQuestion + vbYesNo + vbDefaultButton1, "Etiquetas Almoneda") = vbYes Then
        
        Screen.MousePointer = vbHourglass
        For i = 1 To grdArticulos.Rows
            
            If grdArticulos.CellIcon(i, 1) = frmMDI.img.ItemIndex(1) Then
                
                DoEvents
                bcCodigo.text = ""
                bcCodigo.text = Mid(grdArticulos.CellText(i, 1), 1, 12)
                Sleep 1100
                Imprimir grdArticulos.CellText(i, 1), grdArticulos.CellText(i, 4), grdArticulos.CellText(i, 7), grdArticulos.CellItemData(i, 6), Val(grdArticulos.CellText(i, 3)), Trim(grdArticulos.CellText(i, 2))
            
            End If
        
        Next i
        QuitarImpresos
        Screen.MousePointer = vbDefault
    
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTodos_Click()
Dim i As Long

    For i = 1 To grdArticulos.Rows
        grdArticulos.CellIcon(i, 1) = frmMDI.img.ItemIndex(1)
    Next i
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    CreaEncabezado
    cmbTipoEntrada.ListIndex = 0
    Poner_Flat fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat fl
End Sub

Private Sub grdArticulos_Click(ByVal lRow As Long, ByVal lCol As Long)
    If lCol = 1 And lRow > 0 Then grdArticulos.CellIcon(lRow, lCol) = IIf(grdArticulos.CellIcon(lRow, lCol) = frmMDI.img.ItemIndex(2), frmMDI.img.ItemIndex(1), frmMDI.img.ItemIndex(2))
End Sub

Sub Imprimir(Codigo As String, Peso As Double, Precio As Double, Kilates As Integer, Cantidad As Integer, strPrenda As String)
Dim Impresora As Printer
Dim i As Integer

On Error GoTo error

    DoEvents
    Sleep 500
    bcCodigo.text = Left(Codigo, 12)
    
    Set Impresora = Printer
    With Impresora

        For i = 1 To Cantidad
        
            bcCodigo.text = Left(Codigo, 12)
            .ScaleMode = vbMillimeters
            .Font = "Arial"
            .FontSize = 6.5
    
            'Imprimo el peso
            .CurrentX = Regresa_Valor("ETIQUETAS", "PesoX", 0)
            .CurrentY = Regresa_Valor("ETIQUETAS", "PesoY", 0)
            Impresora.Print Format(Peso, "##,###0.00") & " Grs."
    
            'Imprimo el Kilataje
            .CurrentX = Regresa_Valor("ETIQUETAS", "KilatesX", 0)
            .CurrentY = Regresa_Valor("ETIQUETAS", "KilatesY", 0)
            Impresora.Print SacaKilates(Kilates)
            
            'Imprimo la prenda
            .CurrentX = Regresa_Valor("ETIQUETAS", "PrendaX", 0)
            .CurrentY = Regresa_Valor("ETIQUETAS", "PrendaY", 0)
            Impresora.Print strPrenda
                
            'Imprimo el precio
            .CurrentX = Regresa_Valor("ETIQUETAS", "PrecioX", 0)
            .CurrentY = Regresa_Valor("ETIQUETAS", "PrecioY", 0)
            Impresora.Print Format(Precio, FMoneda)
            
            'Imprimo Leyenda
            .CurrentX = Regresa_Valor("ETIQUETAS", "LeyendaX", 0)
            .CurrentY = Regresa_Valor("ETIQUETAS", "LeyendaY", 0)
            Impresora.Print "Artículo usado sin garantía"
        
            Sleep 500
            'Imprimo el Código de Barras
            .PaintPicture bcCodigo.Picture, Regresa_Valor("ETIQUETAS", "CodigoX", 0), Regresa_Valor("ETIQUETAS", "CodigoY", 0), Regresa_Valor("ETIQUETAS", "Anchocodigo", 0), Regresa_Valor("ETIQUETAS", "Altocodigo", 0)
        
        .EndDoc
        Next i

    End With
    Exit Sub
    
error:
    Maneja_Error Err
End Sub

Function QuitarImpresos()
Dim i As Long

    For i = grdArticulos.Rows To 1 Step -1
        If grdArticulos.CellIcon(i, 1) = frmMDI.img.ItemIndex(1) Then
            grdArticulos.RemoveRow i
        End If
    Next i
End Function
