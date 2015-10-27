VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmRepRegistro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Mensual del Registro"
   ClientHeight    =   1440
   ClientLeft      =   3180
   ClientTop       =   6045
   ClientWidth     =   5745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5745
   StartUpPosition =   1  'CenterOwner
   Begin vbAcceleratorGrid6.vbalGrid grdReportes 
      Height          =   4815
      Left            =   480
      TabIndex        =   8
      Top             =   2160
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   8493
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
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
   End
   Begin VB.TextBox txtFechaIni 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtFechaFin 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFechaFin 
      Height          =   300
      Left            =   3840
      TabIndex        =   1
      Top             =   600
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
      Picture         =   "frmRepRegistro.frx":0000
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFechaIni 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   600
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
      Picture         =   "frmRepRegistro.frx":0115
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   840
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
      Picture         =   "frmRepRegistro.frx":022A
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   240
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
      Picture         =   "frmRepRegistro.frx":077C
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha inicial:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha final:"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   1410
   End
End
Attribute VB_Name = "frmRepRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim FechaIni As String, FechaFin As String, Mes As String, Año As String


Private Sub cmdImprimir_Click()
Dim rcConsulta As New ADODB.Recordset

       FechaIni = txtFechaIni.text
       FechaFin = txtFechaFin.text
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then
        Exit Sub
    End If
        If CDate(FechaIni) <= CDate(FechaFin) Then
        
            rcConsulta.Open "SELECT * FROM sucursales WHERE Activa=1", dbDatos, adOpenForwardOnly, adLockReadOnly
            
            If rcConsulta.RecordCount = 0 Then
                    MsgBox "No se encontraron registros !!", vbInformation, "Error"
            Else
                    Mes = Mid(txtFechaIni.text, 4, 3)
                    Año = Mid(txtFechaIni.text, 8)
                    
                    llenarGrid CDate(FechaIni), CDate(FechaFin)
                    If grdReportes.Rows > 0 Then
                        Exportar_Excel
                        'RepMesReg
                        Unload Me
                    End If
            End If
        Else
                    MsgBox "Las Fechas no son correctas !!", vbInformation, "Error"
            
        End If
    
  
End Sub

Private Sub cmdMosFechaFin_Click()
    txtFechaFin.text = frmCalendario.Fecha(txtFechaFin.text, 1)
End Sub

Private Sub cmdMosFechaIni_Click()
    txtFechaIni.text = frmCalendario.Fecha(txtFechaIni.text, 1)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CrearEncabezados()
    
    With grdReportes
                 
        ' -- Agregar las columnas
        .AddColumn "K1", "N° Contrato", , , , , , , , "0000", , CCLSortNumeric
        .AddColumn "K2", "Fecha", , , , , , , , "dd/mm/yy", , CCLSortDate
        .AddColumn "K3", "Nombre", , , , , , , , , , CCLSortString
        .AddColumn "K4", "Domicilio", , , , , , , , , , CCLSortString
        .AddColumn "K5", "Prestamo", , , , , , , , "$ #0.#0", , CCLSortNumeric
        .AddColumn "K6", "Descripción", , , , , , , , , , CCLSortString
        .AddColumn "K7", "Factura", , , , , , , , , , CCLSortString
        .AddColumn "K8", "CAT", , , , , , , , "$ #0.#0", , CCLSortNumeric
        .Redraw = False
        .Clear
                    
    End With
    
End Sub


Private Sub llenarGrid(Fecha1 As Date, Fecha2 As Date)
Dim columna As Long, Fila As Long, i As Long
Dim rcConsulta As New ADODB.Recordset
Dim rcTmp As New ADODB.Recordset
Dim CAT As Double

Dim SubQuery1 As String
Dim SubQuery2 As String
Dim SubQuery3 As String
Dim SQLQuery As String

On Error GoTo Error

    Screen.MousePointer = vbHourglass
    
    Set rcConsulta = New ADODB.Recordset
    rcConsulta.CursorLocation = adUseClient
    
    'ORO-ELECTRONICOS
    SubQuery1 = "SELECT DISTINCT empeno.NumContrato As Folio, empeno.Fecha AS Fecha, concat(clientes.Nombre,' ',clientes.Apellido) AS Nombre, CONCAT(clientes.Direccion,' ',clientes.NoExterior,' ',clientes.NoInterior,' Col.',clientes.Colonia) as Domicilio,empeno.Prestamo AS Prestamo, " & _
                "CONCAT(detallesempeno.Articulo,' ',detallesempeno.Marca,' ',detallesempeno.Modelo) AS Descripcion,0 as Campo,empeno.VenPeriodo AS Periodo,empeno.TipoTasa AS TipoTasa,empeno.TipoInteres AS TipoInteres,empeno.Serie AS Serie " & _
                "FROM empeno INNER JOIN clientes ON clientes.ID=empeno.IDCliente INNER JOIN detallesempeno ON detallesempeno.IDEmpeno=empeno.ID " & _
                "WHERE empeno.Origen=1 AND empeno.Cancelado=0 AND DATE_FORMAT(empeno.Fecha,'%Y%/%m%/%d')>='" & Format(Fecha1, "YYYY/MM/DD") & "' AND DATE_FORMAT(empeno.Fecha,'%Y%/%m%/%d')<='" & Format(Fecha2, "YYYY/MM/DD") & "' "
    
    'INMUEBLES
    'SubQuery2 = "SELECT DISTINCT empeno.NumContrato As Folio, empeno.Fecha AS Fecha, concat(clientes.Nombre,' ',clientes.Apellido) AS Nombre, CONCAT(clientes.Direccion,' ',clientes.NoExterior,' ',clientes.NoInterior,' Col.',clientes.Colonia) as Domicilio,empeno.Prestamo AS Prestamo, " & _
    '            "CONCAT(mld_tipo_inmuebles.Descripcion,' ',detallesempenoinmuebles.DescInmuebleOtro) AS Descripcion,detallesempenoinmuebles.Superficie As Campo,empeno.VenPeriodo AS Periodo,empeno.TipoTasa AS TipoTasa,empeno.TipoInteres AS TipoInteres,empeno.Serie AS Serie " & _
    '            "FROM empeno INNER JOIN clientes ON clientes.ID=empeno.IDCliente INNER JOIN detallesempenoinmuebles ON detallesempenoinmuebles.IDEmpeno=empeno.ID LEFT JOIN mld_tipo_inmuebles ON detallesempenoinmuebles.IdTipoInmueble = mld_tipo_inmuebles.Id " & _
    '            "WHERE empeno.Origen=1 AND empeno.Cancelado=0 AND SERIE = " & SERIE_E & " AND DATE_FORMAT(empeno.Fecha,'%Y%/%m%/%d')>='" & Format(Fecha1, "YYYY/MM/DD") & "' AND DATE_FORMAT(empeno.Fecha,'%Y%/%m%/%d')<='" & Format(Fecha2, "YYYY/MM/DD") & "' "
    
    'AUTOS
    SubQuery3 = "SELECT DISTINCT empeno.NumContrato As Folio, empeno.Fecha AS Fecha, concat(clientes.Nombre,' ',clientes.Apellido) AS Nombre, CONCAT(clientes.Direccion,' ',clientes.NoExterior,' ',clientes.NoInterior,' Col.',clientes.Colonia) as Domicilio,empeno.Prestamo AS Prestamo, " & _
                "CONCAT('AUTO ',detallesempenoautos.MarcayModelo,' ',detallesempenoautos.Color) AS Descripcion,detallesempenoautos.Año As Campo,empeno.VenPeriodo AS Periodo,empeno.TipoTasa AS TipoTasa,empeno.TipoInteres AS TipoInteres,empeno.Serie AS Serie " & _
                "FROM empeno INNER JOIN clientes ON clientes.ID=empeno.IDCliente INNER JOIN detallesempenoautos ON detallesempenoautos.IDEmpeno=empeno.ID " & _
                "WHERE empeno.Origen=1 AND empeno.Cancelado=0 AND Serie= " & SERIE_B & " AND DATE_FORMAT(empeno.Fecha,'%Y%/%m%/%d')>='" & Format(Fecha1, "YYYY/MM/DD") & "' AND DATE_FORMAT(empeno.Fecha,'%Y%/%m%/%d')<='" & Format(Fecha2, "YYYY/MM/DD") & "' "
    
    SQLQuery = "SELECT * FROM ( " & SubQuery1 & " UNION " & SubQuery3 & " ) AS A ORDER BY Fecha"
    
    grdReportes.Clear False
    'rcConsulta.Open "SELECT DISTINCT empeno.NumContrato As Folio, empeno.Fecha AS Fecha, concat(clientes.Nombre,' ',clientes.Apellido) AS Nombre, concat(clientes.Direccion,' Col. ',clientes.Colonia) AS Domicilio, empeno.Prestamo AS Prestamo, CONCAT(detallesempeno.Articulo,' ',detallesempeno.Marca,' ',detallesempeno.Modelo) AS Descripcion,empeno.VenPeriodo AS Periodo,empeno.TipoTasa AS TipoTasa,empeno.TipoInteres AS TipoInteres,empeno.Serie AS Serie From empeno INNER JOIN clientes ON clientes.ID=empeno.IDCliente LEFT JOIN detallesempeno ON detallesempeno.IDEmpeno=empeno.ID WHERE DATE_FORMAT(empeno.Fecha,'%Y%/%m%/%d')>='" & Format(Fecha1, "YYYY/MM/DD") & "' AND DATE_FORMAT(empeno.Fecha,'%Y%/%m%/%d')<='" & Format(Fecha2, "YYYY/MM/DD") & "' ORDER BY empeno.Fecha", dbDatos, adOpenForwardOnly, adLockReadOnly
    
    rcConsulta.Open SQLQuery, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not rcConsulta.EOF Then
        While Not rcConsulta.EOF
                         
            DoEvents
                     
            CAT = 0
            CAT = Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "Cat", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie='" & rcConsulta!Serie & "' AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion='" & rcConsulta!Periodo & "'"))
            
            With grdReportes
            
                .AddRow
                .CellDetails .Rows, 1, rcConsulta!Folio, DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellIcon(.Rows, 1) = 3
                .CellDetails .Rows, 2, rcConsulta!Fecha, DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 3, rcConsulta!Nombre, DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 4, rcConsulta!Domicilio, DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 5, rcConsulta!Prestamo, DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 6, rcConsulta!Descripcion & " " & IIf(rcConsulta!Campo > 0, CStr(rcConsulta!Campo), ""), DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 8, CAT, DT_LEFT Or DT_WORD_ELLIPSIS
                
                'Pongo el Fondo
                i = i + 1
                Poner_Colores grdReportes, .Rows, i
                
            End With
            
            rcConsulta.MoveNext
            
        Wend
    Else
        Screen.MousePointer = vbDefault
        MsgBox "No se encontraron registros.", vbInformation, Me.Caption
    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
    grdReportes.Redraw = True
    
Exit Sub
        
Error:
    Maneja_Error Err
    Resume
    Set rcConsulta = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub Exportar_Excel()
Dim Excel As Object, i As Integer, Col As Integer, Y As Integer, str As String, detalles As Boolean, Pos As Long, Ban As Boolean
Dim rcConsulta As New ADODB.Recordset
On Error GoTo Error
Err.Clear

   ' If MsgBox(" Desea imprimir los detalles ??", vbQuestion + vbYesNo + vbDefaultButton1, "Reporte Detallado") = vbYes Then
        
   '     detalles = True
   ' Else
        
   '     detalles = False
   ' End If
   
   rcConsulta.Open "SELECT * FROM sucursales WHERE Activa=1", dbDatos, adOpenForwardOnly, adLockReadOnly
   
   

    Screen.MousePointer = vbHourglass
    DoEvents
    
    'Creo la Referencia al Excel
    Set Excel = CreateObject("Excel.application")
    
    With Excel
        
        'Barra de Progreso
   '     PBar.Min = 0
   '     PBar.Max = grdReportes.Rows
        
        'Agrego un Nuevo Libro
        .Workbooks.Add
        
        'Creo los Encabezados
        
        With .ActiveSheet

            .Range("A2").Value = "REPORTE MENSUAL DEL REGISTRO PORMENORIZADO DE LOS CONTRATOS"
            .Range("A2:H2").Merge
            
            .Range("A3").Value = "DE MUTUO CON INTERES Y GARANTIA PRENDARIA PRESENTADO AL PADRON DE CASAS DE EMPEÑO"
            .Range("A3:H3").Merge
           ' .Range("A3").HorizontalAlignment = xlVAlignCenter
           ' .Range("A3").VerticalAlignment = xlVAlignCenter
                        
        End With
        
       ' .Cells(2, 2) = "REPORTE MENSUAL DEL REGISTRO PORMENORIZADO DE LOS CONTRATOS"
       ' .Cells(3, 2) = "DE MUTUO CON INTERES Y GARANTIA PRENDARIA PRESENTADO AL PADRON DE CASAS DE EMPEÑO"
        .Cells(7, 2) = "Nombre: "
        .Cells(7, 3) = UCase(rcConsulta!RazonSocial)
        
        .Cells(8, 2) = "Domicilio: "
        .Cells(8, 3) = UCase(rcConsulta!Direccion)
        
        .Cells(9, 2) = "No. Constancia: "
        .Cells(9, 3) = UCase(Regresa_Valor_BD("NumConstancia"))
        
        .Cells(7, 6) = "RFC: " & UCase(rcConsulta!RFC)
        
        .Cells(8, 6) = "CIUDAD: " & UCase(rcConsulta!Ciudad)
        
        .Cells(9, 6) = "TELEFONO: " & UCase(rcConsulta!Telefono)
        
        .Cells(10, 6) = "E-MAIL: " & IIf(IsNull(rcConsulta!Email), "", rcConsulta!Email)
        
        .Cells(11, 2) = "MES: " & UCase(Mes)
        .Cells(11, 4) = "AÑO: " & Año
        
        
        For i = 1 To grdReportes.Columns
            DoEvents
            .Cells(13, i).formula = grdReportes.ColumnHeader("K" & i)
        Next i
        
        Pos = 13
        
        For i = 1 To grdReportes.Rows
            DoEvents
            
            For Y = 1 To grdReportes.Columns
                
                'Omitir los Detalles
               ' If detalles = False And y = 1 Then
               '     If grdReportes.CellItemData(I, IIf(cmbTipoReporte.text = "EMPEÑOS", 4, 3)) > 0 Then
               '         Ban = True
               '         Exit For
               '     End If
               ' End If
                
                .Cells(Pos + 1, Y).formula = grdReportes.CellText(i, Y)
                
            Next Y
        
            If Ban Then
                Ban = False
            Else
                Pos = Pos + 1
            End If
            
       '     PBar.Value = I
        Next i

        ' autoajustar las columnas
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
        .Columns("J:J").EntireColumn.AutoFit
        .Columns("K:K").EntireColumn.AutoFit
        .Columns("L:L").EntireColumn.AutoFit
        .Columns("M:M").EntireColumn.AutoFit
    
        'Aplicando Formato de Fuente
        .Range("B7:B11").Select
        With .Selection.Font
            .Name = "Arial"
            .Bold = True
            .Size = 10
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
        End With
        
        .Range("A13:H13").Select
        With .Selection.Font
            .Name = "Arial"
            .Bold = True
            .Size = 10
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
        End With
        
        .Range("A2:H3").Select
        With .Selection.Font
            .Name = "Arial"
            .Bold = True
            .Size = 14
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
        End With
        
        'Aplicando Bordes Superior e Inferior
        On Error Resume Next
        .Range("A13:H13").Select
        With .Selection
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End With
        Err.Clear
        
        str = "M" & grdReportes.Rows + 1
        '.ActiveSheet.Range("A1", str).HorizontalAlignment = xlHAlignLeft
        .ActiveSheet.Range("A1", str).HorizontalAlignment = -4131
        .Selection.Interior.ColorIndex = 35
        
        'Hago Visible la Referencia
        .Visible = True

    End With
Exit Sub

Error:
    
    Set Excel = Nothing
    Maneja_Error Err
'    PBar.Value = 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    CrearEncabezados
End Sub
