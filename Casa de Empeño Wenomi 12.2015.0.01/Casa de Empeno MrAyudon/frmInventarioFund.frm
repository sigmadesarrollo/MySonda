VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmInventarioFund 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario Fundición"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInventarioFund.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   13140
   Begin VB.ComboBox cmbKilatajes 
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   7200
      Width           =   1695
   End
   Begin vbAcceleratorGrid6.vbalGrid grdPrendasFund 
      Height          =   6780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   11959
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      GridLineColor   =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   0
      ScrollBarStyle  =   2
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   11895
      TabIndex        =   6
      Top             =   7260
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
      Picture         =   "frmInventarioFund.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   10740
      TabIndex        =   7
      Top             =   7260
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   2
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "       &Aceptar"
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
      Picture         =   "frmInventarioFund.frx":055E
   End
   Begin DevPowerFlatBttn.FlatBttn cmdTodos 
      Height          =   375
      Left            =   9600
      TabIndex        =   8
      Top             =   7260
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmInventarioFund.frx":0AB0
      PictureDisabled =   "frmInventarioFund.frx":0E1A
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCancelar 
      Height          =   375
      Left            =   8400
      TabIndex        =   9
      Top             =   7260
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
      Picture         =   "frmInventarioFund.frx":0F74
      PictureDisabled =   "frmInventarioFund.frx":11C3
   End
   Begin VB.Label lblCostoTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   12120
      TabIndex        =   5
      Top             =   6795
      Width           =   525
   End
   Begin VB.Label lblPeso 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   10440
      TabIndex        =   4
      Top             =   6795
      Width           =   525
   End
   Begin VB.Label lblTotales 
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   6750
      Width           =   13095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Kilataje:"
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
      TabIndex        =   1
      Top             =   7200
      Width           =   945
   End
End
Attribute VB_Name = "frmInventarioFund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbKilatajes_Click()

    If VerificaPrendasMarcadas Then
        
        If MsgBox("Desea cancelar las prendas marcadas ??", vbQuestion + vbYesNo + vbDefaultButton2, "Inventario Fundición") = vbNo Then
                        
            Exit Sub
        Else
            
            MarcarTodas False
        End If
        
    End If
    
    MostrarKilatajes cmbKilatajes.ItemData(cmbKilatajes.ListIndex)

End Sub

Function VerificaPrendasMarcadas() As Boolean
Dim i As Integer
    
    VerificaPrendasMarcadas = False
    With grdPrendasFund
                
        For i = 1 To .Rows
            
            If grdPrendasFund.CellIcon(i, 1) = frmMDI.img.ItemIndex(1) Then
                
                VerificaPrendasMarcadas = True
                Exit For
            
            End If
            
        Next i
    
    End With

End Function

Sub MarcarTodas(Valor As Boolean)
Dim i As Integer, Icono As Integer
    
    If Valor Then
        
        Icono = 1
    Else
        
        Icono = 2
    End If
    
    With grdPrendasFund
                   
        For i = 1 To .Rows
            
            .CellIcon(i, 1) = frmMDI.img.ItemIndex(Icono)
        
        Next i
    
    End With
    
    PonerTotales
End Sub

Private Sub cmbKilatajes_GotFocus()
    Cambiar_Color True, cmbKilatajes
End Sub

Private Sub cmbKilatajes_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbKilatajes_LostFocus()
    Cambiar_Color False, cmbKilatajes
End Sub

Private Sub cmdAceptar_Click()
Dim IDSalida As Long, Folio As Long, i As Integer, Eliminados As Integer, crImporte As Double, Hora As String, Movimiento As Long
    
    'Checo si hay prendas marcadas
    If VerificaPrendasMarcadas Then
        
        If MsgBox("Desea dar de baja las prendas marcadas ??", vbQuestion + vbYesNo + vbDefaultButton1, "Inventario Fundición") = vbYes Then
        
            'Saco el Folio
            Folio = Regresa_Movimiento(False, "FolioSalidaInventario")
            Regresa_Movimiento True, "FolioSalidaInventario"
            
            'Tomo el Movimiento
            Movimiento = Regresa_Movimiento(False)
            Regresa_Movimiento True
            
            'Grabo el Encabezado
            dbDatos.Execute "INSERT INTO salidainventario (Fecha,Folio,IDUsuario,IDSucursal,TipoSalida) VALUES('" & _
                            Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Folio & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & SALIDAVENTAFUNDICION & ")"
            
            'Agarro el ID
            IDSalida = SacaValor("salidainventario", "MAX(ID)")
            
            'Tomo la Hora
            Hora = Time
            
            With grdPrendasFund
                
                crImporte = 0
                For i = 1 To .Rows
                    
                    If .RowVisible(i) And .CellIcon(i, 1) = frmMDI.img.ItemIndex(1) Then
                                            
                        'Grabo el Detalle
                        dbDatos.Execute "INSERT INTO detallessalida (IDSalidaInventario,IDArticulo,Codigo,Descripcion,Kilates,Costo,Peso,Precio,Tipo,IDEmpeno,Observaciones) VALUES (" & _
                                        IDSalida & "," & .CellItemData(i, 3) & ",'" & .CellText(i, 7) & "','" & .CellText(i, 3) & "'," & .CellItemData(i, 4) & "," & CDbl(.CellText(i, 6)) & "," & CDbl(.CellText(i, 5)) & "," & CDbl(.CellText(i, 9)) & "," & Val(.CellText(i, 10)) & "," & Val(.CellText(i, 11)) & ",'" & Trim(.CellText(i, 8)) & "')"
                                            
                        'Pongo la cantidad en cero
                        dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=0,TipoSalida=" & SALIDAVENTAFUNDICION & " WHERE ID=" & Val(.CellItemData(i, 3))
                                                                
                        'Marco la prenda
                        .CellItemData(i, 11) = 1
                    
                    'Saco el Total
                    crImporte = crImporte + CDbl(.CellText(i, 6))
                    End If
                    
                Next i
                
                'Muevo las cuentas contables ***********************************
                'Grabamos el cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,PC,IDUsuario,IDSucursal) VALUES " _
                                & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Venta Fundicion'," & Movimiento & "," & Folio & ",'VF01','200901'," & ConvMoneda(crImporte) & "," & TIPO_CARGO & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                         
                'Grabamos el abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,PC,IDUsuario,IDSucursal) VALUES " _
                                & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Venta Fundicion'," & Movimiento & "," & Folio & ",'VF50','620350'," & ConvMoneda(crImporte) & "," & TIPO_ABONO & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                '****************************************************************
                  
                .Redraw = False
                
                'Quito las prendas marcadas
                For i = .Rows To 1 Step -1
                    
                    If .CellIcon(i, 1) = frmMDI.img.ItemIndex(1) Then .RemoveRow i
                    
                Next i
                
                .ClearSelection
                
                .Redraw = True
                
                'Pongo los totales
                PonerTotales
                
            End With
        
        End If
        
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    MarcarTodas False
End Sub

Private Sub cmdSalir_Click()
    
    If VerificaPrendasMarcadas Then
        
        If MsgBox("Desea cerrar la ventana y cancelar las prendas marcadas ??", vbQuestion + vbYesNo + vbDefaultButton2, "Inventario Fundición") = vbNo Then
                        
            Exit Sub

        End If
        
    End If
        
    Unload Me
End Sub

Private Sub cmdTodos_Click()
    MarcarTodas True
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    CreaEncabezado
    cmbKilatajes.Clear
    cmbKilatajes.AddItem "(TODOS)"
    Cargar_Combos "Descripcion", "Kilatajes", cmbKilatajes, " WHERE IDTipo=1", "Ordenamiento", False, "Clave"
    cmbKilatajes.ListIndex = 0
    cmbKilatajes_Click
    lblTotales.BackColor = RGB(223, 208, 102)
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Sub CreaEncabezado()

    With grdPrendasFund
        .ImageList = frmMDI.img
        .AddColumn "K0", "*", ecgHdrTextALignCentre, , 20, , , , , , , CCLSortString
        .AddColumn "K1", "Contrato", ecgHdrTextALignRight, , 60, , , , , , , CCLSortNumeric
        .AddColumn "K2", "Descripción", ecgHdrTextALignLeft, , 545, , , , , , , CCLSortString
        .AddColumn "K3", "Kilates", ecgHdrTextALignCentre, , 55, , , , , , , CCLSortString
        .AddColumn "K4", "Peso", ecgHdrTextALignRight, , 80, , , , , "###,###0.000", , CCLSortNumeric
        .AddColumn "K5", "Costo", ecgHdrTextALignRight, , 100, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Codigo", ecgHdrTextALignLeft, , 100, False, , , , , , CCLSortString
        .AddColumn "K7", "Observaciones", ecgHdrTextALignLeft, , 100, False, , , , , , CCLSortString
        .AddColumn "K8", "Precio", ecgHdrTextALignRight, , 100, False, , , , , , CCLSortNumeric
        .AddColumn "K9", "Tipo", ecgHdrTextALignRight, , 110, False, , , , , , CCLSortNumeric
        .AddColumn "K10", "IDEmpeno", ecgHdrTextALignRight, , 110, False, , , , , , CCLSortNumeric
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Public Function CargarInventario(FechaIni As Date, FechaFin As Date)
Dim rcInventario As New ADODB.Recordset

On Error GoTo Error
    
    rcInventario.Open "SELECT d.ID,CONCAT(d.Descripcion,' ',d.Observaciones) AS Prenda,d.Peso,k.Clave AS ClaveKilataje,k.Descripcion AS Kilataje,d.Costo,d.Codigo,d.Observaciones,d.Precio,d.Tipo,d.IDEmpeno,em.numcontrato " _
                        & "FROM detallesentradainventario d INNER JOIN entradainventario e ON d.IDEntrada=e.ID LEFT JOIN kilatajes k ON d.Kilates=k.Clave LEFT JOIN empeno em on d.IDEmpeno=em.ID " _
                        & "WHERE d.Cantidad>0 AND d.TipoEntrada=" & D_FUNDICION & " AND DATE_FORMAT(e.Fecha,'%Y%/%m%/%d')>='" & Format(FechaIni, "YYYY/MM/DD") & "' AND DATE_FORMAT(e.Fecha,'%Y%/%m%/%d')<='" & Format(FechaFin, "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockReadOnly
    With grdPrendasFund
        
        .Redraw = False
        .Clear
        While Not rcInventario.EOF
            
            .AddRow
            .CellIcon(.Rows, 1) = frmMDI.img.ItemIndex(2)
            .CellText(.Rows, 2) = rcInventario!NumContrato
            .CellTextAlign(.Rows, 2) = DT_CENTER
            .CellText(.Rows, 3) = rcInventario!Prenda
            .CellItemData(.Rows, 3) = rcInventario!ID
            .CellText(.Rows, 4) = rcInventario!Kilataje
            .CellItemData(.Rows, 4) = rcInventario!ClaveKilataje
            .CellTextAlign(.Rows, 4) = DT_CENTER
            .CellText(.Rows, 5) = rcInventario!Peso
            .CellTextAlign(.Rows, 5) = DT_RIGHT
            .CellText(.Rows, 6) = rcInventario!Costo
            .CellTextAlign(.Rows, 6) = DT_RIGHT
            
            .CellText(.Rows, 7) = rcInventario!Codigo
            .CellText(.Rows, 8) = rcInventario!Observaciones
            .CellText(.Rows, 9) = rcInventario!Precio
            .CellText(.Rows, 10) = rcInventario!Tipo
            .CellText(.Rows, 11) = rcInventario!IDEmpeno
        
        rcInventario.MoveNext
        Wend
        
        .Redraw = True
    End With
    rcInventario.Close
    Set rcInventario = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcInventario = Nothing
End Function

Private Sub grdPrendasFund_Click(ByVal lRow As Long, ByVal lCol As Long)
    If lCol = 1 And lRow > 0 Then grdPrendasFund.CellIcon(lRow, lCol) = IIf(grdPrendasFund.CellIcon(lRow, lCol) = frmMDI.img.ItemIndex(2), frmMDI.img.ItemIndex(1), frmMDI.img.ItemIndex(2))
    PonerTotales
End Sub

Sub MostrarKilatajes(Kilataje As Integer)
Dim i As Integer, Ban As Boolean
    
    With grdPrendasFund
        
        Ban = True
        .Redraw = False
        For i = 1 To .Rows
            
            If .CellItemData(i, 4) <> cmbKilatajes.ItemData(cmbKilatajes.ListIndex) And Kilataje <> 0 Then
                
                .RowVisible(i) = False
            Else
                
                .RowVisible(i) = True
                Colorea grdPrendasFund, CLng(i), IIf(Ban, RGB(236, 252, 222), RGB(255, 255, 255))
                If Ban Then Ban = False Else Ban = True
            End If
                        
        Next i
        .Redraw = True
    End With
    
    PonerTotales
End Sub

Sub PonerTotales()
Dim i As Integer, crPeso As Double, crCosto As Double
    
    crPeso = 0: crCosto = 0
    With grdPrendasFund

        For i = 1 To .Rows
            
            If grdPrendasFund.CellIcon(i, 1) = frmMDI.img.ItemIndex(1) Then
                
                crPeso = crPeso + CDbl(.CellText(i, 5))
                crCosto = crCosto + CDbl(.CellText(i, 6))
            
            End If
        
        Next i
    
    End With
    
    lblPeso.Caption = Format(crPeso, "###,###0.000")
    lblCostoTotal.Caption = Format(crCosto, FMoneda)
End Sub
