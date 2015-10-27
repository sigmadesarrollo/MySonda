VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCorteDivisas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre Divisas"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCorteDivisas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   8970
   Begin vbAcceleratorGrid6.vbalGrid grdCorteDivisas 
      Height          =   4500
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   7938
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
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7785
      TabIndex        =   1
      Top             =   4605
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
      Picture         =   "frmCorteDivisas.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   6570
      TabIndex        =   2
      Top             =   4605
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   2
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "        &Aceptar"
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
      Picture         =   "frmCorteDivisas.frx":055E
   End
End
Attribute VB_Name = "frmCorteDivisas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
Dim Movimiento As Long, Folio As Long, crCantidad As Long, i As Integer, Corte As Long, Hora As String

On Error GoTo error
    

    For i = 1 To grdCorteDivisas.Rows
        
        Corte = 0
        Corte = Val(SacaValor("auxiliar", "ID", " WHERE IDDivisa=" & grdCorteDivisas.CellItemData(i, 1) & " AND Concepto='Corte de Divisas' AND Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND PC='" & NombrePc & "'"))
        If Corte > 0 Then GoTo CierreHecho
        
        'Tomo la Hora
        Hora = Time
        
        'Saca el Movimiento
        Movimiento = Regresa_Movimiento(False)
        Regresa_Movimiento True
        
        'Saca el Folio
        Folio = Regresa_Movimiento(False, "FolioBoveda")
        Regresa_Movimiento True, "FolioBoveda"
        
        'Tomo la Cantidad
        crCantidad = CDbl(grdCorteDivisas.CellText(i, 6))
        
        'Grabamos el cargo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal,IDDivisa) VALUES " _
                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'CV01','910901'," & ConvMoneda(crCantidad) & "," & TIPO_CARGO & ",0,'Corte de Divisas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & grdCorteDivisas.CellItemData(i, 1) & ")"
                      
        'Grabamos el abono
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal,IDDivisa) VALUES " _
                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'CV50','999450'," & ConvMoneda(crCantidad) & "," & TIPO_ABONO & ",0,'Corte de Divisas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & grdCorteDivisas.CellItemData(i, 1) & ")"
             
        'Marco los movimientos de divisas contemplados en el corte
        dbDatos.Execute "UPDATE auxiliar SET Corte=1 WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND PC='" & NombrePc & "' AND IDDivisa=" & grdCorteDivisas.CellItemData(i, 1)
     
CierreHecho:
    Next i
        
    'Imprimo el Corte
    Imprimir
    Exit Sub
    
error:
    Maneja_Error Err
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    
    Crea_Encabezado
    
    Carga_Divisas
    
    CentrarForm Me, frmMDI
End Sub

Sub Crea_Encabezado()

    With grdCorteDivisas
        .AddColumn "C1", "Divisa", , , 150, , , , , , , CCLSortNumeric
        .AddColumn "C2", "Dotaciones", ecgHdrTextALignRight, , 88, , , , , "###,###,###,###0", , CCLSortNumeric
        .AddColumn "C3", "Retiros", ecgHdrTextALignRight, , 88, , , , , "###,###,###,###0", , CCLSortNumeric
        .AddColumn "C4", "Entradas", ecgHdrTextALignRight, , 85, , , , , "###,###,###,###0", , CCLSortNumeric
        .AddColumn "C5", "Salidas", ecgHdrTextALignRight, , 85, , , , , "###,###,###,###0", , CCLSortNumeric
        .AddColumn "C6", "Saldo", ecgHdrTextALignRight, , 90, , , , , "###,###,###,###0", , CCLSortNumeric
    End With

End Sub

Sub Carga_Divisas()
Dim rcConsulta As New ADODB.Recordset, rcAux As New ADODB.Recordset
Dim Compras As Long, Ventas As Long, Dotaciones As Long, Retiros As Long, IDDivisa As Integer, strDivisa As String

On Error GoTo error
        
    rcAux.Open "SELECT DISTINCT IDDivisa FROM auxiliar a WHERE a.IDDivisa>0 AND a.Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND a.PC='" & NombrePc & "' AND (a.Cuenta='999401' OR a.Cuenta='999450' OR a.Cuenta='710301' OR a.Cuenta='710350')", dbDatos, adOpenForwardOnly, adLockReadOnly
    While Not rcAux.EOF
        
        strDivisa = "": Dotaciones = 0: Retiros = 0: Compras = 0: Ventas = 0
        rcConsulta.Open "SELECT dv.Descripcion,a.Cuenta,a.Importe,a.Iniciales,a.Serie,a.IDDivisa FROM auxiliar a INNER JOIN monedas dv ON a.IDDivisa=dv.Clave WHERE a.IDDivisa=" & rcAux!IDDivisa & " AND a.Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND a.PC='" & NombrePc & "' AND (a.Cuenta='999401' OR a.Cuenta='999450' OR a.Cuenta='710301' OR a.Cuenta='710350')", dbDatos, adOpenForwardOnly, adLockReadOnly
        With grdCorteDivisas
            
            IDDivisa = rcConsulta!IDDivisa
            strDivisa = rcConsulta!Descripcion
            While Not rcConsulta.EOF
            
                If rcConsulta!Cuenta = "999401" And rcConsulta!Iniciales = "DODV01" Then
                    
                    Dotaciones = Dotaciones + rcConsulta!Importe
                
                ElseIf rcConsulta!Cuenta = "999450" And rcConsulta!Iniciales = "REDV50" Then
                    
                    Retiros = Retiros + rcConsulta!Importe
                    
                ElseIf rcConsulta!Cuenta = "710301" And rcConsulta!Serie = 2 Then
                    
                    Compras = Compras + rcConsulta!Importe
                    
                ElseIf rcConsulta!Cuenta = "710350" And rcConsulta!Serie = 2 Then
                    
                    Ventas = Ventas + rcConsulta!Importe
                End If
                
            rcConsulta.MoveNext
            Wend
            
            .AddRow
            .CellText(.Rows, 1) = strDivisa
            .CellItemData(.Rows, 1) = IDDivisa
            .CellText(.Rows, 2) = Dotaciones
            .CellTextAlign(.Rows, 2) = DT_RIGHT
            .CellText(.Rows, 3) = Retiros
            .CellTextAlign(.Rows, 3) = DT_RIGHT
            .CellText(.Rows, 4) = Compras
            .CellTextAlign(.Rows, 4) = DT_RIGHT
            .CellText(.Rows, 5) = Ventas
            .CellTextAlign(.Rows, 5) = DT_RIGHT
            .CellText(.Rows, 6) = (Dotaciones + Compras) - (Retiros + Ventas)
            .CellTextAlign(.Rows, 6) = DT_RIGHT
        End With
        rcConsulta.Close
    
    rcAux.MoveNext
    Wend
    rcAux.Close
    Set rcAux = Nothing
    Set rcConsulta = Nothing
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcConsulta = Nothing
    Set rcAux = Nothing
End Sub

Sub Imprimir()
Dim i As Integer
    
    dbReportes.Execute "DELETE FROM cortedivisas WHERE PC='" & NombrePc & "'"
    With grdCorteDivisas
    
        For i = 1 To .Rows
            
            dbReportes.Execute "INSERT INTO cortedivisas (IDDivisa,Dotacion,Retiro,Compras,Ventas,PC) VALUES (" & _
                                .CellItemData(i, 1) & "," & ConvMoneda(.CellText(i, 2)) & "," & ConvMoneda(.CellText(i, 3)) & "," & ConvMoneda(.CellText(i, 4)) & "," & ConvMoneda(.CellText(i, 5)) & ",'" & NombrePc & "')"
            
        Next i
    
    End With
    
    Screen.MousePointer = vbHourglass
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{cortedivisas.PC}='" & NombrePc & "'"
        .ReportFileName = Path & "\Reportes\CierreDivisas.rpt"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .WindowTitle = "Cierre Divisas"
        .Action = 1
    End With
    Screen.MousePointer = vbDefault

End Sub
