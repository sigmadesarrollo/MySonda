VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{0BFA85A1-F9B8-11CF-8939-444553540000}#1.0#0"; "barcode.ocx"
Begin VB.Form frmRemates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pasar a Almoneda"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRemates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   13950
   Begin vbalIml6.vbalImageList lstIcons 
      Left            =   360
      Top             =   7440
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   5740
      Images          =   "frmRemates.frx":000C
      Version         =   131072
      KeyCount        =   5
      Keys            =   "ÿÿÿÿ"
   End
   Begin vbAcceleratorGrid6.vbalGrid grdAlmoneda 
      Height          =   7380
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   13018
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      GridLineColor   =   14737632
      HighlightBackColor=   16744576
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
      Begin VB.ComboBox cmbDestino 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmRemates.frx":1698
         Left            =   1800
         List            =   "frmRemates.frx":16AB
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   12705
      TabIndex        =   2
      Top             =   7500
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
      Picture         =   "frmRemates.frx":16D2
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   11505
      TabIndex        =   3
      Top             =   7500
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
      TextColor       =   4210752
      Object.ToolTipText     =   ""
      Picture         =   "frmRemates.frx":1C24
   End
   Begin BARCODELib.Barcode bcCodigo 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   6240
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
End
Attribute VB_Name = "frmRemates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 30/04/2002
' Modulo frmRemates - frmRemates.frm
' Ultima Modificacion - 30/04/2002
''Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

Option Explicit

Dim Fl() As cFlatControl
Dim Salir As Integer, Descargar As Integer, TipoPrenda As Integer

Private Sub cmbDestino_Click()
Dim Destino As Integer

    If cmbDestino.ListIndex >= 0 Then
        
        Select Case cmbDestino.text

            Case "VENTA"
                
                Destino = 5
        
            Case "FUNDICION"
                
                Destino = 6
        
            Case "OTRO"
                
                Destino = 7
                
            Case "CENTRAL"
                
                Destino = 8
            
            Case Else
                
                Destino = 0
        End Select
        
        grdAlmoneda.CellText(grdAlmoneda.SelectedRow, 2) = cmbDestino.text
        grdAlmoneda.CellItemData(grdAlmoneda.SelectedRow, 2) = Destino 'cmbDestino.ItemData(cmbDestino.ListIndex)
        grdAlmoneda.CancelEdit
    End If

End Sub

Private Sub cmbDestino_GotFocus()
    Cambiar_Color True, cmbDestino
End Sub

Private Sub cmbDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmbDestino.Visible = False
End Sub

Private Sub cmbDestino_LostFocus()
    Cambiar_Color False, cmbDestino
    grdAlmoneda.CancelEdit
End Sub

Private Sub cmdAceptar_Click()

    If grdAlmoneda.Rows > 0 Then
        
        Pasar_Articulos_Almoneda
    
    End If

End Sub

'Imprimimos la piezas  a remate
Private Sub Imprimir_Almoneda()

On Error GoTo Error

    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\ArticulosAlmoneda.rpt"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        
        .SubreportToChange = "Resumen"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{articulos.Destino}='VENTA' AND {articulos.Kilates}>0"
        
        .SubreportToChange = "ResumenFundicion"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{articulos.Destino}='FUNDICION' AND {articulos.Kilates}>0"
        
        .WindowState = crptMaximized
        .WindowTitle = "Contratos a Almoneda"
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub cmdSalir_Click()
    Salir = 1
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

'Inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Salir = 0
    Descargar = 0
    Crear_Encabezados
    Mostrar_Articulos_Almoneda TipoPrenda
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

'Creamos los encabezados del grid
Private Sub Crear_Encabezados()

    With grdAlmoneda
        .ImageList = lstIcons
        .AddColumn "K1", "Contrato", ecgHdrTextALignRight, , 60, , , , , , , CCLSortNumeric
        .AddColumn "K2", "Destino", ecgHdrTextALignLeft, , 80, , , , , , , CCLSortString
        .AddColumn "K3", "Cliente", ecgHdrTextALignLeft, , 342, , , , , , , CCLSortString
        .AddColumn "K4", "Iniciales", ecgHdrTextALignLeft, , 52, False, , , , , , CCLSortString
        .AddColumn "K5", "Fecha", ecgHdrTextALignCentre, , 60, False, , , , "DD/MM/YY", , CCLSortDate
        .AddColumn "K6", "Préstamo", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Avalúo", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K8", "Dias Venc.", ecgHdrTextALignRight, , 70, , , , , , , CCLSortNumeric
        .AddColumn "K9", "", ecgHdrTextALignRight, , 0, , , , , , , CCLSortNumeric
        .AddColumn "K10", "", ecgHdrTextALignRight, , 0, , , , , , , CCLSortNumeric
        .AddColumn "K11", "", ecgHdrTextALignRight, , 0, , , , , , , CCLSortNumeric
        .AddColumn "K12", "", ecgHdrTextALignRight, , 0, , , , , , , CCLSortNumeric
        .AddColumn "K13", "NumPrendas", ecgHdrTextALignRight, , 30, False, , , , , , CCLSortNumeric
        .AddColumn "K14", "Etiqueta", ecgHdrTextALignCentre, , 60, , , , , , , CCLSortString
        
        .AddColumn "K15", "Código", ecgHdrTextALignLeft, , 60, False, , , , , , CCLSortString
        .AddColumn "K16", "Peso", ecgHdrTextALignLeft, , 60, False, , , , , , CCLSortString
        .AddColumn "K17", "Precio", ecgHdrTextALignLeft, , 60, False, , , , , , CCLSortString
        .AddColumn "K18", "Kilates", ecgHdrTextALignLeft, , 60, False, , , , , , CCLSortString
        .AddColumn "K19", "Cantidad", ecgHdrTextALignLeft, , 60, False, , , , , , CCLSortString
        .AddColumn "K20", "Prenda", ecgHdrTextALignLeft, , 60, False, , , , , , CCLSortString
        .AddColumn "K21", "Tipo Contrato", ecgHdrTextALignLeft, , 110, , , , , , , CCLSortString
        .AddColumn "K22", "AvaluoDiamante", ecgHdrTextALignLeft, , 60, False, , , , , , CCLSortString
        .AddColumn "K23", "Prenda", ecgHdrTextALignLeft, , 60, False, , , , , , CCLSortString
        .AddColumn "K24", "Fecha Com.", ecgHdrTextALignLeft, , 110, , , , , , , CCLSortString
    End With

End Sub

'Cargamos los articulos a remate
Private Sub Mostrar_Articulos_Almoneda(IDTipoPrenda As Integer)
Dim rcRemate As New ADODB.Recordset
Dim diasEnajenacion As Integer, strConsulta As String

On Error GoTo Error
    
    If IDTipoPrenda = 0 Then
        
        strConsulta = ""
    ElseIf IDTipoPrenda = -1 Then
        
        strConsulta = "LEFT JOIN detallesempenoautos de ON e.ID=de.IDEmpeno"
    Else
        
        strConsulta = "LEFT JOIN detallesempeno de ON e.ID=de.IDEmpeno"
    End If
    
    diasEnajenacion = Regresa_Valor_BD("DiasEnajenacion")
    rcRemate.Open "SELECT COUNT(e.ID) AS Total FROM empeno e " & strConsulta & " WHERE " & IIf(IDTipoPrenda = 0, "", IIf(IDTipoPrenda = -1, "e.Serie=" & SERIE_B & " AND ", " de.Tipo=" & IDTipoPrenda & " AND ")) & "e.Destino=0 AND e.Pagado=0 AND e.Cancelado=0 AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0," & diasEnajenacion & ") DAY),'%Y%/%m%/%d') <'" & Format(Date, "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockReadOnly
    If rcRemate!Total > 0 Then frmMDI.Bar.Value = 0: frmMDI.Bar.Min = 0: frmMDI.Bar.Max = rcRemate!Total Else Exit Sub
    rcRemate.Close
     
    rcRemate.Open "SELECT DISTINCT e.ID,e.NumContrato,e.Folio,e.Fecha,e.Prestamo,e.Avaluo,e.Origen,e.Vencimiento,e.TipoInteres,e.TipoTasa,e.Serie,c.Iniciales,CONCAT(c.Apellido,' ',c.Nombre) AS Cliente " _
                & "FROM empeno e " & strConsulta & " LEFT JOIN clientes c ON e.IDCliente=c.ID WHERE " & IIf(IDTipoPrenda = 0, "", IIf(IDTipoPrenda = -1, "e.Serie=" & SERIE_B & " AND ", " de.Tipo=" & IDTipoPrenda & " AND ")) & "e.Destino=0 AND e.Pagado=0 AND e.Cancelado=0 AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0," & diasEnajenacion & ") DAY),'%Y%/%m%/%d') <'" & Format(Date, "YYYY/MM/DD") & "' ORDER BY e.NumContrato", dbDatos, adOpenForwardOnly, adLockReadOnly
       
    grdAlmoneda.Redraw = False
    frmMDI.Bar.Visible = True
    
    With rcRemate
    
        While Not .EOF

            DoEvents
                        
            grdAlmoneda.AddRow
            grdAlmoneda.CellText(grdAlmoneda.Rows, 1) = !NumContrato
            grdAlmoneda.CellIcon(grdAlmoneda.Rows, 1) = lstIcons.ItemIndex(2)
            grdAlmoneda.CellItemData(grdAlmoneda.Rows, 1) = !ID
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdAlmoneda.CellText(grdAlmoneda.Rows, 3) = !Cliente
            grdAlmoneda.CellItemData(grdAlmoneda.Rows, 3) = !Origen
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdAlmoneda.CellText(grdAlmoneda.Rows, 4) = !Iniciales
            grdAlmoneda.CellItemData(grdAlmoneda.Rows, 4) = !Folio
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 4) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdAlmoneda.CellText(grdAlmoneda.Rows, 5) = !Fecha
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 5) = DT_CENTER
            grdAlmoneda.CellText(grdAlmoneda.Rows, 6) = !Prestamo
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdAlmoneda.CellText(grdAlmoneda.Rows, 7) = !Avaluo
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 7) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdAlmoneda.CellText(grdAlmoneda.Rows, 8) = DateDiff("D", !Vencimiento, Date)
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 8) = DT_RIGHT
            grdAlmoneda.CellText(grdAlmoneda.Rows, 9) = 0
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 9) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdAlmoneda.CellText(grdAlmoneda.Rows, 10) = 0
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 10) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdAlmoneda.CellText(grdAlmoneda.Rows, 11) = 0
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 11) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdAlmoneda.CellText(grdAlmoneda.Rows, 12) = 0
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 12) = DT_RIGHT Or DT_WORD_ELLIPSIS
                  
            grdAlmoneda.CellText(grdAlmoneda.Rows, 21) = IIf(!TipoInteres = "TRADICIONAL", "TRAD.", !TipoInteres) & "-" & !TipoTasa
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 21) = DT_LEFT
      
      
            grdAlmoneda.CellText(grdAlmoneda.Rows, 24) = DateAdd("D", diasEnajenacion, !Vencimiento)
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 24) = DT_RIGHT
      
            Poner_Colores grdAlmoneda, grdAlmoneda.Rows, CLng(frmMDI.Bar.Value + 1)
            
            'Pongo el Detalle de los Empeños y El Numero de Prendas del Contrato
            grdAlmoneda.CellText(grdAlmoneda.Rows, 13) = DetalleEmpeno(!ID, !Serie)
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 13) = DT_RIGHT Or DT_WORD_ELLIPSIS
        
        .MoveNext
        frmMDI.Bar.Value = CLng(frmMDI.Bar.Value) + 1
        Wend
        
        'Pongo lo totales
        Totales 0, CLng(frmMDI.Bar.Value)
    End With
    rcRemate.Close
    Set rcRemate = Nothing
    grdAlmoneda.Redraw = True
    frmMDI.Bar.Visible = False
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcRemate = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Or Salir = 1 Then
        
        If MarcadosRemate(True) > 0 Then
        
            If MsgBox("Desea salir sin mandar las prendas marcadas a Almoneda ??", vbQuestion + vbYesNo + vbDefaultButton2, "Pasar a Almoneda") = vbYes Then
            
                Descargar = 1
            Else
            
                Descargar = 0
            End If
        
        Else
            
            Descargar = 1
        End If
        
    Else
        
        Descargar = 1
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Descargar = 1 Then
        
        Quitar_Flat Fl
    Else
        
        Cancel = True
    End If
    
End Sub

Private Sub grdAlmoneda_CancelEdit()
    cmbDestino.Visible = False
End Sub

Private Sub grdAlmoneda_Click(ByVal lRow As Long, ByVal lCol As Long)
      
    If lCol = 0 Or lRow = 0 Then
        
        GoTo Mostrar
    ElseIf grdAlmoneda.CellItemData(lRow, 6) = 0 Then
        
        GoTo Mostrar
    ElseIf lCol = 14 And lRow > 0 And grdAlmoneda.CellIcon(lRow, lCol) = 4 And grdAlmoneda.CellItemData(lRow, 6) > 0 Then
        
        grdAlmoneda.CellIcon(lRow, lCol) = 5
        Exit Sub
    ElseIf lCol = 14 And lRow > 0 Then
                
        grdAlmoneda.CellIcon(lRow, lCol) = 4
        DoEvents
        bcCodigo.text = ""
        bcCodigo.text = Mid(grdAlmoneda.CellText(lRow, 15), 1, 12)
        Sleep 1000
        ImprimirEtiqueta grdAlmoneda.CellText(lRow, 15), grdAlmoneda.CellText(lRow, 16), CDbl(grdAlmoneda.CellText(lRow, 17)), SacaKilates(grdAlmoneda.CellText(lRow, 18)), Val(grdAlmoneda.CellText(lRow, 19)), Trim(grdAlmoneda.CellText(lRow, 23))
        Exit Sub
    End If
                      
Mostrar:
    If lCol = 0 Or lRow = 0 Then
        
        Exit Sub
    ElseIf grdAlmoneda.CellItemData(lRow, 6) > 0 Then
    
        Exit Sub
    ElseIf lRow = grdAlmoneda.Rows Then
        
        Exit Sub
    ElseIf lCol = 1 And lRow > 0 And grdAlmoneda.CellIcon(lRow, lCol) = 1 Then
        
        grdAlmoneda.CellIcon(lRow, lCol) = 2
        MuestraOculta grdAlmoneda.CellItemData(lRow, 1), True
    ElseIf lCol = 1 And lRow > 0 Then
        
        grdAlmoneda.CellIcon(lRow, lCol) = 1
        MuestraOculta grdAlmoneda.CellItemData(lRow, 1), False
    End If

End Sub

'Pasamos los articulos que sean a remate y ponemos los movimientos
Private Sub Pasar_Articulos_Almoneda()

    Dim Indice As Long, Movimiento As Long, Folio As Long, Prestamo As Double, Cantidad As Integer, Serie As Integer, IDEntrada As Long, Hora As String
    Dim rcTmp As New ADODB.Recordset

On Error GoTo Error
        
    IDEntrada = MarcadosRemate()
    
    If IDEntrada > 0 Then
        
        If MsgBox("Esta seguro de pasar las prendas seleccionadas a Almoneda ??", vbQuestion + vbYesNo + vbDefaultButton2, "Pasar a Almoneda") = vbYes Then
            
            Screen.MousePointer = vbHourglass
            
            dbReportes.Execute "DELETE FROM articulos"
            
            'Tomo la Hora
            Hora = Time
            
            'Saco el movimiento
            Movimiento = Regresa_Movimiento(False)
            Regresa_Movimiento True
                    
            For Indice = 1 To grdAlmoneda.Rows - 1
                
                If Trim(grdAlmoneda.CellText(Indice, 2)) <> "" Then
                
                    Select Case grdAlmoneda.CellText(Indice, 2)
            
                        Case "VENTA"
                            
                            Serie = SERIE_A
            
                        Case "FUNDICION"
                            
                            Serie = SERIE_B
            
                        Case "OTRO"
                        
                            Serie = SERIE_C
                            
                        Case "CENTRAL"
                        
                            Serie = SERIE_D
                    
                    End Select
                        
                    'Marco la prenda como pasada a Remate
                    rcTmp.Open "SELECT de.IDEmpeno,de.Codigo,de.Cantidad,de.Articulo,de.Peso,de.Kilates,de.Avaluo,de.Prestamo,de.Tipo,de.Observaciones FROM detallesempeno de WHERE de.ID=" & grdAlmoneda.CellItemData(Indice, 6), dbDatos, adOpenForwardOnly, adLockReadOnly
                    
                    With rcTmp
                        
                        If Not .BOF And Not .EOF Then
                            
                            dbReportes.Execute "INSERT INTO articulos (IDEmpeno,Articulo,Peso,Kilates,Avaluo,Prestamo,Cantidad,Destino,Observaciones) VALUES (" & _
                                                grdAlmoneda.CellItemData(Indice, 3) & ",'" & grdAlmoneda.CellText(Indice, 15) & " " & IIf(Val(grdAlmoneda.CellText(Indice, 19)) > 0, Val(grdAlmoneda.CellText(Indice, 19)) & " ", "") & grdAlmoneda.CellText(Indice, 3) & "'," & ConvMoneda(grdAlmoneda.CellText(Indice, 16)) & "," & Val(grdAlmoneda.CellText(Indice, 18)) & "," & ConvMoneda(grdAlmoneda.CellText(Indice, 7)) & "," & ConvMoneda(grdAlmoneda.CellText(Indice, 6)) & "," & Val(grdAlmoneda.CellText(Indice, 19)) & ",'" & Trim(grdAlmoneda.CellText(Indice, 2)) & "','" & Trim(grdAlmoneda.CellText(Indice, 20)) & "')"
                            
                            dbDatos.Execute "UPDATE detallesempeno SET Almoneda=1,Destino=" & Val(grdAlmoneda.CellItemData(Indice, 2)) & " WHERE ID=" & grdAlmoneda.CellItemData(Indice, 6)
            
                        End If
                    End With
                    
                    rcTmp.Close
                    
                    'Verifico si ya se pasaron todas las prendas para marcar la boleta que paso a Remate
                    'Y quito el Contrato
                    If Val(grdAlmoneda.CellText(Indice, 21)) = SERIE_B Then
                        
                        dbDatos.Execute "UPDATE empeno SET Pagado=1,Destino=" & D_ALMONEDA & ", FechaAlmoneda='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Almoneda=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & grdAlmoneda.CellItemData(Indice, 3)
                    Else
                        
                        rcTmp.Open "SELECT SUM(de.Almoneda) AS Almoneda,COUNT(de.ID) AS NumPrendas FROM detallesempeno de WHERE de.IDEmpeno=" & grdAlmoneda.CellItemData(Indice, 3), dbDatos, adOpenForwardOnly, adLockReadOnly
                        If rcTmp!Almoneda = rcTmp!numprendas Then
                            
                            dbDatos.Execute "UPDATE empeno SET Pagado=1,Destino=" & D_ALMONEDA & ", FechaAlmoneda='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Almoneda=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & grdAlmoneda.CellItemData(Indice, 3)
                        End If
                        rcTmp.Close
                        Set rcTmp = Nothing
                        
                    End If
                    
                    'Saco el Folio, el Préstamo y la Cantidad
                    Folio = CLng(grdAlmoneda.CellItemData(Indice, 5))
                    Cantidad = Val(grdAlmoneda.CellItemData(Indice, 4))
                    Prestamo = CDbl(grdAlmoneda.CellText(Indice, 6)) * Cantidad
                    
                    'Muevo las Cuentas Contables
                    If grdAlmoneda.CellText(Indice, 2) = "VENTA" Then
                        
                        Pasar_Inventario Indice, IDEntrada, ENTRADAALMONEDA
                        
                        'Grabamos el cargo
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & Folio & ",'RE01','620301'," & ConvMoneda(Prestamo) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                 
                        'Grabamos el abono
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & Folio & ",'RE50','201750'," & ConvMoneda(Prestamo) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                        
                        
                        
                    ElseIf grdAlmoneda.CellText(Indice, 2) = "CENTRAL" Then
                    
                    Pasar_Inventario Indice, IDEntrada, D_CENTRAL
                        
                        'Grabamos el cargo
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & Folio & ",'RE01','310101'," & ConvMoneda(Prestamo) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                 
                        'Grabamos el abono
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & Folio & ",'RE50','201750'," & ConvMoneda(Prestamo) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                        
                        
                    Else
                        
                        Pasar_Inventario Indice, IDEntrada, IIf(Serie = SERIE_B, D_FUNDICION, D_OTRO)
                        
                        'Grabamos el cargo
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & Folio & ",'RE01','200901'," & ConvMoneda(Prestamo) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                 
                        'Grabamos el abono
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & Folio & ",'RE50','201750'," & ConvMoneda(Prestamo) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                    End If
                    
                End If
            
            Next Indice
            
            'Borro las prendas marcadas
            Borrar_Articulos
            
            Sleep 1000
            
            'Imprimo el reporte de las prendas pasadas
            Imprimir_Almoneda
            
            'Pongo los nuevos totales
            Totales 1
            
            Screen.MousePointer = vbDefault
            
        End If
        
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcTmp = Nothing
    Screen.MousePointer = vbDefault
End Sub

'Borramos los articulos seleccionados
Private Sub Borrar_Articulos()
Dim Indice As Long, Ban As Boolean, Eliminados As Long

    grdAlmoneda.Redraw = False
    grdAlmoneda.RemoveRow (grdAlmoneda.Rows)
    Eliminados = 0
    
    'Quito los Detalles de las prendas Marcados
    For Indice = 1 To grdAlmoneda.Rows
        
        If Trim(grdAlmoneda.CellText(Indice - Eliminados, 2)) <> "" Then
            
            If DescuentaPrendas(grdAlmoneda.CellItemData(Indice - Eliminados, 3)) Then Eliminados = Eliminados + 1
            grdAlmoneda.RemoveRow Indice - Eliminados
            Eliminados = Eliminados + 1
        End If
        
    Next Indice
    
    'Oculto los Detalles de los Empeños
    For Indice = 1 To grdAlmoneda.Rows
            
        If Val(grdAlmoneda.CellItemData(Indice, 6)) > 0 Then grdAlmoneda.RowVisible(Indice) = False
    Next Indice
        
    'Coloreo los Contratos
    Ban = False

    For Indice = 1 To grdAlmoneda.Rows
        
        If grdAlmoneda.RowVisible(Indice) Then
        
            If Ban Then Ban = False Else Ban = True
            Poner_Colores grdAlmoneda, Indice, IIf(Ban, 1, 2)
        
        End If
        
    Next Indice
    
    'Coloreo el Detalle de los Contratos
    For Indice = 1 To grdAlmoneda.Rows

        If grdAlmoneda.CellItemData(Indice, 6) > 0 Then
        
            SombreaGrid grdAlmoneda, 239, 239, 239, 255, 255, 255, CInt(Indice)
        
        End If
        
    Next Indice
    
    grdAlmoneda.Redraw = True
    
End Sub

Private Sub grdAlmoneda_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, sText As String
    
    If lCol <> 2 Then Exit Sub
    If lRow = grdAlmoneda.Rows Then Exit Sub
    If grdAlmoneda.CellIcon(lRow, 1) <> -1 Then Exit Sub
    
    grdAlmoneda.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight

    If Not IsMissing(grdAlmoneda.CellText(lRow, lCol)) Then
        
        sText = grdAlmoneda.CellText(lRow, lCol)
    Else
        
        sText = ""
    End If
   
    cmbDestino.Move lLeft + 40, lTop + 25, lWidth - 60
    cmbDestino.Visible = True
    cmbDestino.ZOrder
    cmbDestino.SetFocus
End Sub

Function Pasar_Inventario(Renglon As Long, IDEntrada As Long, TipoEntrada As Integer)

    Dim rcTmp As New ADODB.Recordset
    
    With rcTmp
                
        If Val(grdAlmoneda.CellText(Renglon, 21)) = SERIE_B Then
            
            'Tabla de DetalleEntradaInventario
            dbDatos.Execute "INSERT INTO detallesentradainventario(IDEntrada,Codigo,Tipo,Cantidad,Descripcion,Peso,Kilates,Precio,Costo,Estado,Marca,Modelo,Serie,Color,Tamano,TipoPrenda,Observaciones,IDEmpeno,SucursalOrigen,TipoEntrada,PrecioVitrina,IDDetallesEmpeno) VALUES (" & _
                            IDEntrada & ",'" & grdAlmoneda.CellText(Renglon, 15) & "',0,1,'" & grdAlmoneda.CellText(Renglon, 3) & "',0,0," & ConvMoneda(grdAlmoneda.CellText(Renglon, 17)) & "," & ConvMoneda(grdAlmoneda.CellText(Renglon, 6)) & ",'','','','','','',0,'" & Trim(grdAlmoneda.CellText(Renglon, 20)) & "'," & CLng(grdAlmoneda.CellItemData(Renglon, 3)) & "," & frmMDI.IDSucursal & "," & ENTRADAALMONEDA & "," & ConvMoneda(grdAlmoneda.CellText(Renglon, 17)) & "," & grdAlmoneda.CellItemData(Renglon, 6) & ")"
            
        Else
            
            .Open "SELECT e.ID,e.Fecha,e.Folio,e.Vencimiento,e.TipoInteres,d.ID AS IDPrenda,d.Articulo,d.Kilates,d.Peso,d.Avaluo,d.Cantidad,d.Prestamo,d.Tipo,d.IDEmpeno,d.Marca,d.Modelo,d.Serie,d.Color,d.Tamano,d.Codigo,d.TipoPrenda,d.Observaciones,d.Estado,d.CantidadPiedras,d.PesoPiedras,d.CantidadDiamantes,d.Puntos,d.PrestamoDiamante FROM detallesempeno d INNER JOIN empeno e ON d.IDEmpeno=e.ID WHERE d.ID=" & Val(grdAlmoneda.CellItemData(Renglon, 6)) & " ORDER BY d.Codigo", dbDatos, adOpenForwardOnly, adLockReadOnly
            
            'Tabla de DetalleEntradaInventario
            dbDatos.Execute "INSERT INTO detallesentradainventario(IDEntrada,Codigo,Tipo,Cantidad,Descripcion,Peso,Kilates,Precio,Costo,Estado,Marca,Modelo,Serie,Color,Tamano,TipoPrenda,Observaciones,IDEmpeno,SucursalOrigen,TipoEntrada,PrecioVitrina,CantidadPiedras,PesoPiedras,CantidadDiamantes,Puntos,PrestamoDiamante,IDDetallesEmpeno) VALUES (" & _
                            IDEntrada & ",'" & !Codigo & "'," & !Tipo & "," & !Cantidad & ",'" & !Articulo & "'," & ConvMoneda(!Peso) & "," & !Kilates & "," & ConvMoneda(grdAlmoneda.CellText(Renglon, 17)) & "," & ConvMoneda(!Prestamo) & ",'" & !Estado & "','" & !Marca & "','" & !Modelo & "','" & !Serie & "','" & !Color & "','" & !Tamano & "'," & !TipoPrenda & ",'" & !Observaciones & "'," & !IDEmpeno & "," & frmMDI.IDSucursal & "," & TipoEntrada & "," & ConvMoneda(grdAlmoneda.CellText(Renglon, 17)) & "," & !CantidadPiedras & "," & ConvMoneda(!PesoPiedras) & "," & !CantidadDiamantes & "," & ConvMoneda(!Puntos) & "," & ConvMoneda(!PrestamoDiamante) & "," & !IDPrenda & ")"
            
            .Close
            Set rcTmp = Nothing
        End If
    
    End With
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function

Function Totales(Optional x As Integer = 0, Optional NumContratos As Long)
Dim i  As Long, TotPrestamo As Double, TotInteres As Double
Dim z As Integer

    If grdAlmoneda.Rows = 0 Then Exit Function
    If NumContratos > 0 Then z = 0 Else z = 1
    For i = 1 To grdAlmoneda.Rows
        
        If grdAlmoneda.CellItemData(i, 1) > 0 Then
            
            TotPrestamo = TotPrestamo + CDbl(grdAlmoneda.CellText(i, 6))
            TotInteres = TotInteres + CDbl(grdAlmoneda.CellText(i, 7))
            NumContratos = NumContratos + z
        End If

    Next i

    'Pongo los Totales ************
    grdAlmoneda.AddRow
    grdAlmoneda.CellText(grdAlmoneda.Rows, 6) = TotPrestamo
    grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 6) = DT_RIGHT
    grdAlmoneda.CellText(grdAlmoneda.Rows, 7) = TotInteres
    grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 7) = DT_RIGHT
    
    For i = 1 To grdAlmoneda.Columns
        
        grdAlmoneda.CellBackColor(grdAlmoneda.Rows, i) = RGB(223, 208, 102)
        grdAlmoneda.CellForeColor(grdAlmoneda.Rows, i) = RGB(29, 64, 226)
    
    Next i

    grdAlmoneda.CellText(grdAlmoneda.Rows, 1) = "Num. " & NumContratos
    grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 1) = DT_RIGHT
    '***************************
    
End Function

Function MuestraOculta(ID As Long, Opcion As Boolean, Optional OcultaTodas As Boolean = False)
    Dim i As Long

    For i = 1 To grdAlmoneda.Rows

        If grdAlmoneda.CellItemData(i, 3) = ID And grdAlmoneda.CellItemData(i, 5) > 0 Then
            
            grdAlmoneda.RowVisible(i) = Opcion
        
        End If
    
    Next i

End Function

Function DetalleEmpeno(ID As Long, Serie As Integer) As Integer
Dim Cantidad As Integer, Codigo As String, Peso As Double, Kilates As Integer, NumPrenda As Integer, crPrecio As Double, AvaluoDiam As Double, strSql As String, strDescripcion As String, strPrenda As String
Dim rcConsulta As New ADODB.Recordset

On Error GoTo Error
    
    NumPrenda = 0
    With rcConsulta
        
        If Serie <> SERIE_B Then 'Serie = SERIE_A Or Serie = SERIE_C
            strSql = "SELECT empeno.NumContrato,empeno.Serie,d.ID,d.IDEmpeno,d.Codigo,d.Cantidad,d.Tipo,d.Articulo,d.Peso,d.Prestamo,d.Avaluo,d.Observaciones,d.Estado,d.Kilates,d.PesoPiedras,d.PrestamoDiamante,kilatajes.Descripcion AS Kilataje " _
                    & "FROM detallesempeno d LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave INNER JOIN empeno ON d.IDEmpeno=empeno.ID WHERE d.Almoneda=0 AND d.IDEmpeno=" & ID
        Else
            
            strSql = "SELECT empeno.NumContrato,empeno.Serie,d.ID,d.IDEmpeno,d.MarcayModelo,d.Placas,d.Año,d.Color,d.SerieChasis,d.NumMotor,d.NumTarjetaCircu,empeno.Prestamo,empeno.Avaluo,d.Observaciones " _
                    & "FROM detallesempenoautos d INNER JOIN empeno ON d.IDEmpeno=empeno.ID WHERE d.IDEmpeno=" & ID
        End If
        
        .Open strSql, dbDatos, adOpenForwardOnly, adLockReadOnly
        
        While Not .EOF
            
            If Serie <> SERIE_B Then 'Serie = SERIE_A Or Serie = SERIE_C
                
                Codigo = !Codigo
                Cantidad = !Cantidad
                Peso = !Peso
                Kilates = IIf(IsNull(!Kilates), 0, !Kilates)
                strDescripcion = !Codigo & " - " & !Cantidad & " " & !Articulo & " " & !Observaciones & " " & !Kilataje & IIf(rcConsulta!Tipo = 1, " " & Format(rcConsulta!Peso, "###.000") & " Grms.", "") & IIf(IsNull(rcConsulta!Estado) Or Trim(rcConsulta!Estado) = "", "", " ESTADO: " & rcConsulta!Estado)
                strPrenda = !Articulo
                
                'Saco el Precio
                crPrecio = Redondeo(!Prestamo * (1 + (Regresa_Valor_BD("GtosVenta") / 100)))
                
                If !PrestamoDiamante > 0 Then
                
                    AvaluoDiam = CDbl(Regresa_Valor_BD("PrestamoAvaluoDiamante"))
                    crPrecio = Redondeo(crPrecio + ((!PrestamoDiamante * 100) / AvaluoDiam))
                Else
                    
                    AvaluoDiam = 0
                End If
            
            Else
            
                Codigo = CreaCodigoBarras(Trim(Format(frmMDI.IDSucursal, "000")), ENTRADAEMPENO, !NumContrato, 1)
                Cantidad = 1
                Peso = 0
                Kilates = 0
                strPrenda = ""
                crPrecio = !Prestamo * (1 + (Regresa_Valor_BD("PrecioAutos") / 100))
                strDescripcion = "MARCA Y MODELO: " & !MarcayModelo & ", PLACAS: " & !Placas & ", AÑO: " & !Año & ", COLOR: " & !Color & ", SERIE CHASIS: " & !SerieChasis & ", NUM. MOTOR: " & !NumMotor & ", TARJETA CIRC.: " & !NumTarjetaCircu
            End If
            
            NumPrenda = NumPrenda + 1
            grdAlmoneda.AddRow
            grdAlmoneda.CellText(grdAlmoneda.Rows, 3) = strDescripcion
            grdAlmoneda.CellItemData(grdAlmoneda.Rows, 3) = !IDEmpeno
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 3) = DT_LEFT Or DT_END_ELLIPSIS
            grdAlmoneda.CellItemData(grdAlmoneda.Rows, 4) = Cantidad
            grdAlmoneda.CellItemData(grdAlmoneda.Rows, 5) = !NumContrato
            grdAlmoneda.CellText(grdAlmoneda.Rows, 6) = !Prestamo
            grdAlmoneda.CellItemData(grdAlmoneda.Rows, 6) = !ID
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 6) = DT_RIGHT
            grdAlmoneda.CellText(grdAlmoneda.Rows, 7) = !Avaluo
            grdAlmoneda.CellTextAlign(grdAlmoneda.Rows, 7) = DT_RIGHT
            
            grdAlmoneda.CellText(grdAlmoneda.Rows, 15) = Codigo
            grdAlmoneda.CellText(grdAlmoneda.Rows, 16) = Peso
            grdAlmoneda.CellText(grdAlmoneda.Rows, 17) = crPrecio
            grdAlmoneda.CellText(grdAlmoneda.Rows, 18) = Kilates
            grdAlmoneda.CellText(grdAlmoneda.Rows, 19) = Cantidad
            grdAlmoneda.CellText(grdAlmoneda.Rows, 20) = !Observaciones
            grdAlmoneda.CellText(grdAlmoneda.Rows, 21) = !Serie
            
            grdAlmoneda.CellText(grdAlmoneda.Rows, 23) = strPrenda
            
            grdAlmoneda.RowVisible(grdAlmoneda.Rows) = False
            SombreaGrid grdAlmoneda, 239, 239, 239, 255, 255, 255, grdAlmoneda.Rows
        
        .MoveNext
        Wend
        
        .Close
        Set rcConsulta = Nothing
    End With
    
    DetalleEmpeno = NumPrenda
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Function

Function MarcadosRemate(Optional Salir As Boolean = False) As Long
Dim Indice As Long, Folio As Long

On Error GoTo Error

    MarcadosRemate = 0
    
    For Indice = 1 To grdAlmoneda.Rows - 1

        If Trim(grdAlmoneda.CellText(Indice, 2)) <> "" Then
        
            If Salir Then
                
                MarcadosRemate = 1
            Else
                
                'Saco el Folio
                Folio = Regresa_Movimiento(False, "FolioInventario")
                Regresa_Movimiento True, "FolioInventario"
            
                'Tabla Entrada Inventario
                dbDatos.Execute "INSERT INTO entradainventario(Fecha,Folio,TipoEntrada,IDUsuario,IDSucursal) VALUES " _
                                & "('" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Folio & "," & ENTRADAALMONEDA & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Saco el ID de la Entrada
                MarcadosRemate = SacaValor("entradainventario", "MAX(ID)")
            End If
            
            Exit For
        End If

    Next Indice
    Exit Function
    
Error:
    Maneja_Error Err
End Function

Function DescuentaPrendas(IDEmpeno As Long) As Boolean
Dim Indice As Long
    
    DescuentaPrendas = False
    
    For Indice = 1 To grdAlmoneda.Rows
        
        If grdAlmoneda.CellItemData(Indice, 1) = IDEmpeno Then
        
            grdAlmoneda.CellText(Indice, 13) = Val(grdAlmoneda.CellText(Indice, 13)) - 1
            If Val(grdAlmoneda.CellText(Indice, 13)) = 0 Then
            
                grdAlmoneda.RemoveRow Indice: DescuentaPrendas = True
            Else
                
                grdAlmoneda.CellIcon(Indice, 1) = 1
            End If
                
            Exit For
        End If
        
    Next Indice

End Function

Sub ImprimirEtiqueta(Codigo As String, Peso As Double, Precio As Double, Kilates As String, Cantidad As Integer, strPrenda As String)
Dim Impresora As Printer
Dim i As Integer

On Error GoTo Error

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
            Impresora.Print Kilates
            
            'Imprimo la prenda
            .CurrentX = Regresa_Valor("ETIQUETAS", "PrendaX", 0)
            .CurrentY = Regresa_Valor("ETIQUETAS", "PrendaY", 0)
            Impresora.Print strPrenda
                
            'Imprimo el precio
            .CurrentX = Regresa_Valor("ETIQUETAS", "PrecioX", 0)
            .CurrentY = Regresa_Valor("ETIQUETAS", "PrecioY", 0)
            Impresora.Print Format(Precio, FMoneda)
        
            Sleep 500
            
            'Imprimo el Código de Barras
            .PaintPicture bcCodigo.Picture, Regresa_Valor("ETIQUETAS", "CodigoX", 0), Regresa_Valor("ETIQUETAS", "CodigoY", 0), Regresa_Valor("ETIQUETAS", "Anchocodigo", 0), Regresa_Valor("ETIQUETAS", "Altocodigo", 0)
        
        .EndDoc
        Next i

    End With
    Set Impresora = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set Impresora = Nothing
End Sub

Public Sub MuestraPrendas(IDTipo As Integer)
    TipoPrenda = 0
    TipoPrenda = IDTipo
    Me.Show
End Sub
