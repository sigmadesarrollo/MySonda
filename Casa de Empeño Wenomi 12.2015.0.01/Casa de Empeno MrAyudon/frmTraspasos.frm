VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmTraspasos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso de Inventario"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTraspasos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   12360
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5715
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   12195
      Begin VB.TextBox txtDestino 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   540
         Width           =   3525
      End
      Begin VB.TextBox txtCodigo 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1155
         Width           =   1695
      End
      Begin vbAcceleratorGrid6.vbalGrid grdTraspasos 
         Height          =   4140
         Left            =   120
         TabIndex        =   3
         Top             =   1515
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   7303
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
      Begin DevPowerFlatBttn.FlatBttn cmbDestino 
         Height          =   240
         Left            =   3660
         TabIndex        =   4
         Top             =   540
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   423
         AutoSize        =   0   'False
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   1
         TextColor       =   -2147483630
         Object.ToolTipText     =   ""
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMuestraArticulos 
         Height          =   240
         Left            =   1830
         TabIndex        =   5
         Top             =   1170
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   423
         AutoSize        =   0   'False
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   1
         TextColor       =   -2147483630
         Object.ToolTipText     =   ""
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   4920
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   8550
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblFolio3 
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   5880
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblFecha3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01/Ene/2006"
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
         Left            =   9510
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblTotalTraspasos 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   870
         Left            =   840
         TabIndex        =   6
         Top             =   5400
         Visible         =   0   'False
         Width           =   4170
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   11130
      TabIndex        =   13
      Top             =   5775
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
      Picture         =   "frmTraspasos.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   9930
      TabIndex        =   14
      Top             =   5775
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
      Picture         =   "frmTraspasos.frx":009D
   End
End
Attribute VB_Name = "frmTraspasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbDestino_Click()
frmMostrarSucursales.ver Me, txtDestino, True
End Sub

Private Sub cmdAceptar_Click()
If grdTraspasos.Rows > 0 And txtDestino.Tag <> "" Then
    Traspaso
    grdTraspasos.Clear
    txtDestino.text = ""
    txtDestino.Tag = ""
    txtCodigo.text = ""
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdMuestraArticulos_Click()
frmMuestraarticulos.ver Me, txtCodigo, True, 1
End Sub

Private Sub Form_Load()
Inicializar
End Sub

Sub Inicializar()
Screen.MousePointer = vbHourglass
CentrarForm Me, frmMDI
Limpiar "Consignación"
Frame3.BorderStyle = 0
Crear_Encabezados
Poner_Flat Fl, Me.Controls, Me
Screen.MousePointer = vbDefault
End Sub

Public Sub Crear_Encabezados()
With Me.grdTraspasos
    .AddColumn "K1", "Código", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
    .AddColumn "K2", "Artículo", ecgHdrTextALignLeft, , 270, , , , , , , CCLSortString
    .AddColumn "K3", "Peso", ecgHdrTextALignRight, , 60, , , , , "0.000", , CCLSortNumeric
    .AddColumn "K4", "Kilates", ecgHdrTextALignRight, , 70, , , , , , , CCLSortString
    .AddColumn "K5", "Avalúo", ecgHdrTextALignRight, , 100, , , , , "###,###,###.00", , CCLSortNumeric
    .AddColumn "K6", "Préstamo", ecgHdrTextALignRight, , 100, , , , , "###,###,###.00", , CCLSortNumeric
    .AddColumn "K7", "P. Venta", ecgHdrTextALignRight, , 100, , , , , "###,###,###.00", , CCLSortNumeric
End With
End Sub

Private Sub Limpiar(Contededor As String)
  Dim ctrl As Control
  
  For Each ctrl In Controls
    On Error Resume Next
    If ctrl.Container.Caption = Contededor Then
      If TypeOf ctrl Is TextBox And ctrl.Name <> "NotaRef" And ctrl.Name <> "FechaCap" And ctrl.Name <> "VencimientoCap" Then ctrl.text = ""
      If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = ""
      If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
      On Error Resume Next
      ctrl.Tag = ""
    End If
  Next
End Sub

Private Sub Traspaso()
Dim Folio As Long, Movimiento As Long, Cont As Integer, IDTraspaso As Double, Cuenta As String, Hora As String
Dim rcArticulos As ADODB.Recordset
Dim rcSucursal As ADODB.Recordset

On Error GoTo error

If MsgBox("Estan correctos los datos ??", vbYesNo + vbQuestion, "Traspaso de Inventario") = vbYes Then
    
    'Tomo el Folio
    Folio = Regresa_Movimiento(False, "FolioTraspasos")
    Regresa_Movimiento True, "FolioTraspasos"
    
    'Tomo la Hora
    Hora = Time
    
    'Tabla de traspasos
    dbDatos.Execute "insert into traspasos (Fecha,Folio,IDUsuario,IDSucursal,SucursalDestino) values ('" & Format(Date, "YYYY/MM/DD") & "'," & Folio & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & Val(txtDestino.Tag) & ")"
    Set rcArticulos = dbDatos.Execute("Select max(ID) as IDD from traspasos")
    IDTraspaso = rcArticulos!idd
    
    With grdTraspasos
        For Cont = 1 To .Rows
            
            Set rcArticulos = dbDatos.Execute("select * from detallesentradainventario where ID=" & .CellItemData(Cont, 1) & "")
            
            'Tabla DetallesTraspasos
            dbDatos.Execute "insert into detallestraspasos(IDTraspaso,Codigo,Descripcion,kilates,peso,precio,costo,cantidad,tipo,serie,IDEmpeno,SucursalOrigen,SucursalDestino,IDUsuario) values " _
            & "(" & IDTraspaso & ",'" & rcArticulos!Codigo & "','" & rcArticulos!Descripcion & "'," & rcArticulos!Kilates & "," & rcArticulos!Peso & "," & rcArticulos!Precio & "," & rcArticulos!Costo & "," _
            & rcArticulos!Cantidad & "," & rcArticulos!Tipo & ",'" & rcArticulos!Serie & "'," & rcArticulos!IDEmpeno & "," & frmMDI.IDSucursal & "," & Val(txtDestino.Tag) & "," & frmMDI.IDUsuario & ")"
            
            'DetallesEntradaInventario
            dbDatos.Execute "update detallesentradainventario set cantidad=cantidad-1,SucursalDestino=" & Val(txtDestino.Tag) & ",TipoSalida=" & SALIDATRASPASO & " where ID=" & grdTraspasos.CellItemData(Cont, 1) & ""
        
        Next Cont
    End With
End If

    Set rcSucursal = dbDatos.Execute("select cuenta from sucursales where Clave=" & Val(txtDestino.Tag) & "")
    Cuenta = Left(rcSucursal!Cuenta, 4)
    Cuenta = Cuenta & "01"
    
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                  "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'TP01','" & Trim(Cuenta) & "'," & CDbl(lblTotalTraspasos.Caption) & "," & TIPO_CARGO & ",0,'Traspaso','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabamos el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                  "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'TP50','620350'," & CDbl(lblTotalTraspasos.Caption) & "," & TIPO_ABONO & ",0,'Traspaso','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

    
    'Imprimimos el reporte de los articulos traspasados
    Imprimir Folio
    
error:
    Maneja_Error Err
    Set rcArticulos = Nothing
    Set rcSucursal = Nothing

End Sub

Public Function Buscar(ID As Long)
Dim rcSucursales As ADODB.Recordset

On Error GoTo error

Set rcSucursales = dbDatos.Execute("select NombreComercial,ID,Clave from sucursales where ID=" & ID & "")
If Not rcSucursales.BOF And Not rcSucursales.EOF Then
    txtDestino.text = rcSucursales!NombreComercial
    txtDestino.Tag = rcSucursales!Clave
End If

error:
    Maneja_Error Err
    Set rcSucursales = Nothing
    
End Function

Public Function MuestraDatos(ID As Long)
Dim rcArticulos As ADODB.Recordset

On Error GoTo error

Set rcArticulos = dbDatos.Execute("select * from detallesentradainventario where ID=" & ID & "")
If Not rcArticulos.BOF And Not rcArticulos.EOF Then
    With grdTraspasos
        
        .Redraw = False
        .AddRow
        .CellText(.Rows, 1) = rcArticulos!Codigo
        .CellItemData(.Rows, 1) = rcArticulos!ID
        .CellText(.Rows, 2) = rcArticulos!Descripcion
        .CellText(.Rows, 3) = rcArticulos!Peso
        .CellTextAlign(.Rows, 3) = DT_RIGHT
        .CellText(.Rows, 4) = SacaKilates(rcArticulos!Kilates)
        .CellItemData(.Rows, 4) = rcArticulos!Kilates
        .CellTextAlign(.Rows, 4) = DT_RIGHT
        .CellText(.Rows, 5) = rcArticulos!Precio
        .CellTextAlign(.Rows, 5) = DT_RIGHT
        .CellText(.Rows, 6) = rcArticulos!Costo
        .CellTextAlign(.Rows, 6) = DT_RIGHT
        .CellText(.Rows, 7) = rcArticulos!Precio
        .CellTextAlign(.Rows, 7) = DT_RIGHT
        .Redraw = True
    End With
    
    lblTotalTraspasos.Caption = Format(SacaTotal, "###,###,###,###.00")
End If

error:
    Maneja_Error Err
    Set rcArticulos = Nothing

End Function

Private Sub Form_Unload(Cancel As Integer)
Quitar_Flat Fl
End Sub

Private Sub grdTraspasos_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
If grdTraspasos.Rows > 0 And grdTraspasos.SelectedRow > 0 Then
    If KeyCode = vbKeyDelete Then
        
        If MsgBox("Desea eliminar el articulo seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Traspasos") = vbYes Then grdTraspasos.RemoveRow grdTraspasos.SelectedRow: lblTotalTraspasos.Caption = Format(SacaTotal, "###,###,###,###.00")
    
    End If
End If
End Sub

Function SacaTotal()
Dim i As Integer, Total As Double

If grdTraspasos.Rows > 0 Then
    For i = 1 To grdTraspasos.Rows
        Total = Total + IIf(grdTraspasos.CellText(i, 7) = "", 0, grdTraspasos.CellText(i, 7))
    Next i
Else
    Total = 0
End If

SacaTotal = Total
End Function

Private Sub txtCodigo_GotFocus()
Seleccionar_Texto txtCodigo
Cambiar_Color True, txtCodigo
End Sub

Private Sub txtCodigo_LostFocus()
Cambiar_Color False, txtCodigo
End Sub

Private Sub txtDestino_GotFocus()
Seleccionar_Texto txtDestino
Cambiar_Color True, txtDestino
End Sub

Private Sub txtDestino_LostFocus()
Cambiar_Color False, txtDestino
End Sub

Function Imprimir(Folio As Long)
With frmMDI.Cr
   .Reset
   .WindowShowPrintSetupBtn = True
   .WindowShowExportBtn = True
   .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
   .ReportFileName = Path & "\Reportes\Traspasos.rpt"
   .SelectionFormula = "{Traspasos.Folio}=" & Folio & ""
   .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
   .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
   .Formulas(2) = "Encabezado='SUCURSAL DESTINO:" & Trim(txtDestino.text) & "'"
   .WindowTitle = "Reporte de traspasos"
   .DiscardSavedData = True
   .WindowState = crptMaximized
   .Action = 1
End With
End Function
