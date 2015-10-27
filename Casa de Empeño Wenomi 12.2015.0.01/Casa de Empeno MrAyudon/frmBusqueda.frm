VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmBusqueda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda Contratos"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBusqueda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   14595
   Begin VB.TextBox txtIniciales 
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosCliente 
      Height          =   270
      Left            =   4125
      TabIndex        =   2
      Top             =   345
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
      AlignCaption    =   4
      AutoSize        =   0   'False
      Caption         =   ". . ."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
   Begin VB.Frame Frame1 
      Caption         =   "Búsqueda por fechas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   5280
      Begin VB.TextBox txtHasta 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   1440
      End
      Begin VB.TextBox txtDesde 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
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
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   630
         Width           =   1440
      End
      Begin DevPowerFlatBttn.FlatBttn cmdDesde 
         Height          =   300
         Left            =   1560
         TabIndex        =   5
         Top             =   630
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         AlignCaption    =   4
         AlignPicture    =   4
         AutoSize        =   0   'False
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         MousePointer    =   1
         TextColor       =   -2147483630
         Object.ToolTipText     =   ""
         Picture         =   "frmBusqueda.frx":000C
      End
      Begin DevPowerFlatBttn.FlatBttn cmdHasta 
         Height          =   300
         Left            =   3510
         TabIndex        =   6
         Top             =   630
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         AlignCaption    =   4
         AlignPicture    =   4
         AutoSize        =   0   'False
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         MousePointer    =   1
         TextColor       =   -2147483630
         Object.ToolTipText     =   ""
         Picture         =   "frmBusqueda.frx":0121
      End
      Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
         Height          =   375
         Left            =   4005
         TabIndex        =   7
         Top             =   570
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
         Picture         =   "frmBusqueda.frx":0236
      End
      Begin VB.Label Label3 
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
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label2 
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
         TabIndex        =   4
         Top             =   360
         Width           =   765
      End
   End
   Begin vbAcceleratorGrid6.vbalGrid grdBusqueda 
      Height          =   5850
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   10319
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
   Begin VB.TextBox txtNombre 
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
      Top             =   360
      Width           =   3975
   End
   Begin vbalIml6.vbalImageList lstIcons 
      Left            =   3960
      Top             =   840
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   5740
      Images          =   "frmBusqueda.frx":05BB
      Version         =   131072
      KeyCount        =   5
      Keys            =   "ÿÿÿÿ"
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   13440
      TabIndex        =   14
      Top             =   7380
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
      Picture         =   "frmBusqueda.frx":1C47
   End
   Begin DevPowerFlatBttn.FlatBttn cmdPerdida 
      Height          =   375
      Left            =   12240
      TabIndex        =   15
      Top             =   7380
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
      Picture         =   "frmBusqueda.frx":2199
      PictureDisabled =   "frmBusqueda.frx":26EB
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Iniciales:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1800
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "frmBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 04/04/02
' Modulo frmBusqueda - frmBusqueda.frm
' Ultima Modificacion - 05/04/02
' Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdBuscar_Click()

    If Validar_Fechas Then Buscar 2, , Format(txtDesde.text, "MM/DD/YY"), Format(txtHasta.text, "MM/DD/YY")
End Sub

'Validamos las fechas para la busqueda
Private Function Validar_Fechas() As Boolean
    
    Validar_Fechas = True
  
    If CDate(txtDesde.text) > CDate(txtHasta.text) Then
        
        MsgBox "La fecha inicial no puede ser mayor a la fecha final", vbExclamation + vbOKOnly, "Búsqueda contratos"
        txtDesde.SetFocus
        Exit Function
    End If
  
End Function

Private Sub cmdPerdida_Click()
    
    If grdBusqueda.Rows = 0 Then Exit Sub
    If grdBusqueda.SelectedRow > 0 Then
                                                        
        If grdBusqueda.CellItemData(grdBusqueda.SelectedRow, 10) = 0 And grdBusqueda.CellItemData(grdBusqueda.SelectedRow, 6) = 0 And grdBusqueda.CellItemData(grdBusqueda.SelectedRow, 2) = 0 Then
            
            If MsgBox("Desea marcar el contrato seleccionado como perdido ??", vbQuestion + vbYesNo + vbDefaultButton2, "Búsqueda Contratos") = vbYes Then
                
                dbDatos.Execute "UPDATE empeno SET Perdida=1 WHERE ID=" & Val(grdBusqueda.CellItemData(grdBusqueda.SelectedRow, 1))
                grdBusqueda.CellItemData(grdBusqueda.SelectedRow, 10) = 1
                Colorea grdBusqueda, grdBusqueda.SelectedRow, RGB(244, 119, 66)
                grdBusqueda.ClearSelection
            
            End If
            
        ElseIf grdBusqueda.CellItemData(grdBusqueda.SelectedRow, 10) = 1 And grdBusqueda.CellItemData(grdBusqueda.SelectedRow, 6) = 0 And grdBusqueda.CellItemData(grdBusqueda.SelectedRow, 2) = 0 Then
            
            If MsgBox("Desea eliminar la marca de perdido al contrato seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Búsqueda Contratos") = vbYes Then
                
                dbDatos.Execute "UPDATE empeno SET Perdida=0 WHERE ID=" & Val(grdBusqueda.CellItemData(grdBusqueda.SelectedRow, 1))
                grdBusqueda.CellItemData(grdBusqueda.SelectedRow, 10) = 0
                Poner_Colores grdBusqueda, grdBusqueda.SelectedRow, grdBusqueda.CellItemData(grdBusqueda.SelectedRow, 9)
                grdBusqueda.ClearSelection
            
            End If
            
        Else
            
            grdBusqueda.ClearSelection
        End If
    
    Else
        
        MsgBox "Seleccione el contrato que desea marcar como perdido !!", vbInformation, "Búsqueda Contratos"
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdDesde_Click()
    txtDesde.text = frmCalendario.Fecha(txtDesde.text)
End Sub

Private Sub cmdHasta_Click()
    txtHasta.text = frmCalendario.Fecha(txtHasta.text)
End Sub

Private Sub cmdMosCliente_Click()
    frmMostrarCliente.ver Me, txtNombre
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

'Inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    CentrarForm Me, frmMDI
    txtDesde.text = Format(Date - 1, "DD/MM/YYYY")
    txtHasta.text = Format(Date, "DD/MM/YYYY")
    Poner_Flat Fl, Me.Controls, Me
    Crear_Encabezados
    Screen.MousePointer = vbDefault
End Sub

'Creamos los encabezados del grid
Private Sub Crear_Encabezados()

    With grdBusqueda
        .ImageList = lstIcons
        .AddColumn "K1", "Fecha", ecgHdrTextALignCentre, , 75, , , , , "DD/MM/YY", , CCLSortDate
        .AddColumn "K2", "Contrato", ecgHdrTextALignRight, , 80, , , , , , , CCLSortNumeric
        .AddColumn "K3", "Folio", ecgHdrTextALignRight, , 80, False, , , , , , CCLSortNumeric
        .AddColumn "K4", "Cliente", ecgHdrTextALignLeft, , 260, , , , , , , CCLSortString
        .AddColumn "K5", "Préstamo", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Vencimiento", ecgHdrTextALignCentre, , 70, , , , , "DD/MM/YY", , CCLSortDate
        .AddColumn "K7", "Origen/Folio", ecgHdrTextALignRight, , 86, , , , , , , CCLSortString
        .AddColumn "K8", "Dest./Folio", ecgHdrTextALignRight, , 86, , , , , , , CCLSortString
        .AddColumn "K9", "Fecha Movto.", ecgHdrTextALignCentre, , 78, , , , , , , CCLSortDate
        .AddColumn "K10", "Tasa", ecgHdrTextALignLeft, , 120, , , , , , , CCLSortString
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdBusqueda_ColumnClick(ByVal lCol As Long)
    Ordenar_Grid lCol, grdBusqueda, 5, 6
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

Private Sub txtIniciales_GotFocus()
    Seleccionar_Texto txtIniciales
    Cambiar_Color True, txtIniciales
End Sub

Private Sub txtIniciales_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Buscar 3, txtIniciales.text
    Pasar_Foco KeyAscii
End Sub

Private Sub txtIniciales_LostFocus()
    Cambiar_Color False, txtIniciales
End Sub

Private Sub txtNombre_GotFocus()
    Seleccionar_Texto txtNombre
    Cambiar_Color True, txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Buscar 1, txtNombre.text
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNombre_LostFocus()
    Cambiar_Color False, txtNombre
End Sub

'buscamos al cliente y le mandamos el parametro del tipo de busqueda
Public Sub Buscar(Opcion As Integer, Optional Cliente As String = "", Optional Desde As String = "", Optional Hasta As String = "")
Dim rcBusqueda As New ADODB.Recordset
Dim i As Long
  
On Error GoTo error

    Screen.MousePointer = vbHourglass
  
    'Depende del metodo por el cual se hara el filtro (nombre, iniciales, fecha)
    If Opcion <> 2 And Opcion <> 3 Then
    
        rcBusqueda.Open "SELECT clientes.Iniciales,clientes.Nombre,clientes.Apellido,empeno.ID,empeno.Fecha,empeno.NumContrato,empeno.Folio,empeno.Prestamo,empeno.Origen,empeno.FolioOrigen,empeno.Destino,empeno.FolioDestino,empeno.FechaMovimiento,empeno.Vencimiento,empeno.TipoInteres,empeno.TipoTasa,empeno.Perdida,empeno.Pagado,empeno.Cancelado FROM empeno INNER JOIN clientes ON empeno.IDCliente=clientes.ID WHERE CONCAT(clientes.Apellido,' ',clientes.Nombre)='" & Cliente & "' ORDER BY empeno.Fecha,empeno.Folio", dbDatos, adOpenForwardOnly, adLockOptimistic
    
    ElseIf Opcion = 2 Then
        
        txtIniciales.text = ""
        txtNombre.text = ""
        rcBusqueda.Open "SELECT clientes.iniciales,clientes.Nombre,clientes.Apellido,empeno.ID,empeno.Fecha,empeno.NumContrato,empeno.Folio,empeno.Prestamo,empeno.Origen,empeno.FolioOrigen,empeno.Destino,empeno.FolioDestino,empeno.FechaMovimiento,empeno.Vencimiento,empeno.TipoInteres,empeno.TipoTasa,empeno.Perdida,empeno.Pagado,empeno.Cancelado FROM empeno INNER JOIN clientes ON empeno.IDCliente=clientes.ID WHERE empeno.Fecha BETWEEN '" & Format(Desde, "YYYY/MM/DD") & "' AND '" & Format(Hasta, "YYYY/MM/DD") & "' ORDER BY empeno.Fecha,empeno.Folio", dbDatos, adOpenForwardOnly, adLockOptimistic
    
    ElseIf Opcion = 3 Then
        
        rcBusqueda.Open "SELECT clientes.iniciales,clientes.Nombre,clientes.Apellido,empeno.ID,empeno.Fecha,empeno.NumContrato,empeno.Folio,empeno.Prestamo,empeno.Origen,empeno.FolioOrigen,empeno.Destino,empeno.FolioDestino,empeno.FechaMovimiento,empeno.Vencimiento,empeno.TipoInteres,empeno.Perdida,empeno.Pagado,empeno.Cancelado FROM empeno INNER JOIN clientes ON empeno.IDCliente=clientes.ID WHERE clientes.Iniciales='" & Cliente & "' ORDER BY empeno.Fecha,empeno.Folio", dbDatos, adOpenForwardOnly, adLockOptimistic
    End If
  
    If rcBusqueda.BOF Or rcBusqueda.EOF Then
        
        MsgBox "No se encontró información relacionada con el cliente en la fecha seleccionada !!", vbInformation, "Búsqueda contratos"
    
    Else
    
        txtIniciales.text = rcBusqueda!Iniciales
        grdBusqueda.Redraw = False
        grdBusqueda.Clear
    
        With rcBusqueda
            While Not .EOF
                i = i + 1
                grdBusqueda.AddRow
                grdBusqueda.CellText(grdBusqueda.Rows, 1) = !Fecha
                grdBusqueda.CellItemData(grdBusqueda.Rows, 1) = !ID
                grdBusqueda.CellIcon(grdBusqueda.Rows, 1) = 3
                grdBusqueda.CellTextAlign(grdBusqueda.Rows, 1) = DT_CENTER Or DT_WORD_ELLIPSIS
                grdBusqueda.CellText(grdBusqueda.Rows, 2) = !NumContrato
                grdBusqueda.CellItemData(grdBusqueda.Rows, 2) = !cancelado
                grdBusqueda.CellTextAlign(grdBusqueda.Rows, 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
                grdBusqueda.CellText(grdBusqueda.Rows, 3) = !Folio
                grdBusqueda.CellTextAlign(grdBusqueda.Rows, 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
                grdBusqueda.CellText(grdBusqueda.Rows, 4) = !apellido & " " & !Nombre
                grdBusqueda.CellTextAlign(grdBusqueda.Rows, 4) = DT_LEFT Or DT_WORD_ELLIPSIS
                grdBusqueda.CellText(grdBusqueda.Rows, 5) = Format(!Prestamo, "Currency")
                grdBusqueda.CellTextAlign(grdBusqueda.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                grdBusqueda.CellText(grdBusqueda.Rows, 6) = !Vencimiento
                grdBusqueda.CellItemData(grdBusqueda.Rows, 6) = !Pagado
                grdBusqueda.CellTextAlign(grdBusqueda.Rows, 6) = DT_CENTER Or DT_WORD_ELLIPSIS
                grdBusqueda.CellText(grdBusqueda.Rows, 7) = OD_Origen(!origen) & IIf(!cancelado = 1 And !Destino = 0, "/Cancelado", "/" & !FolioOrigen)
                grdBusqueda.CellTextAlign(grdBusqueda.Rows, 7) = DT_RIGHT Or DT_WORD_ELLIPSIS
                grdBusqueda.CellText(grdBusqueda.Rows, 8) = OD_Origen(!Destino) & IIf(!cancelado = 1 And !Destino > 0, "/Cancelado", IIf(Val(!Destino) = D_VENTA Or Val(!Destino) = OD_REFRENDO, "/" & !foliodestino, ""))
                grdBusqueda.CellTextAlign(grdBusqueda.Rows, 8) = DT_RIGHT Or DT_WORD_ELLIPSIS
                grdBusqueda.CellText(grdBusqueda.Rows, 9) = !FechaMovimiento
                grdBusqueda.CellItemData(grdBusqueda.Rows, 9) = i
                grdBusqueda.CellTextAlign(grdBusqueda.Rows, 9) = DT_CENTER Or DT_WORD_ELLIPSIS
                grdBusqueda.CellText(grdBusqueda.Rows, 10) = !TipoInteres & "-" & !TipoTasa
                grdBusqueda.CellTextAlign(grdBusqueda.Rows, 10) = DT_LEFT Or DT_WORD_ELLIPSIS
                grdBusqueda.CellItemData(grdBusqueda.Rows, 10) = !Perdida
                
                Poner_Colores grdBusqueda, grdBusqueda.Rows, i
          
                DetalleEmpeños !ID
            .MoveNext
            Wend
            
        End With
    
        grdBusqueda.Redraw = True
    End If
    rcBusqueda.Close
    Set rcBusqueda = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcBusqueda = Nothing
    Screen.MousePointer = vbDefault
End Sub

Function DetalleEmpeños(ID As Long)
Dim rcConsulta As ADODB.Recordset

On Error GoTo error

    Set rcConsulta = dbDatos.Execute("select detallesempeno.IDEmpeno,detallesempeno.Cantidad,detallesempeno.Articulo,detallesempeno.Peso,kilatajes.Descripcion,detallesempeno.Prestamo,detallesempeno.Avaluo from detallesempeno Left Join kilatajes on detallesempeno.Kilates=kilatajes.Clave where detallesempeno.IDEmpeno=" & ID & "")

    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        rcConsulta.MoveFirst
        While Not rcConsulta.EOF
            grdBusqueda.AddRow
            grdBusqueda.CellText(grdBusqueda.Rows, 3) = rcConsulta!Cantidad
            grdBusqueda.CellTextAlign(grdBusqueda.Rows, 3) = DT_RIGHT
        
            grdBusqueda.CellText(grdBusqueda.Rows, 4) = rcConsulta!Articulo & " " & Format(rcConsulta!Peso, "###.000") & " Grms. " & rcConsulta!Descripcion
            grdBusqueda.CellItemData(grdBusqueda.Rows, 4) = rcConsulta!IDEmpeno
            grdBusqueda.CellTextAlign(grdBusqueda.Rows, 4) = DT_LEFT
        
            grdBusqueda.CellText(grdBusqueda.Rows, 5) = rcConsulta!Prestamo
            grdBusqueda.CellTextAlign(grdBusqueda.Rows, 5) = DT_RIGHT
        
            grdBusqueda.RowVisible(grdBusqueda.Rows) = False
            SombreaGrid grdBusqueda, 239, 239, 239, 255, 255, 255, grdBusqueda.Rows
        rcConsulta.MoveNext
        Wend
    End If

error:

    If Err > 0 Then Maneja_Error Err
    Set rcConsulta = Nothing
End Function

Function MuestraOculta(ID As Long, Opcion As Boolean)
Dim i As Long

    For i = 1 To grdBusqueda.Rows

        If grdBusqueda.CellItemData(i, 4) = ID Then
            grdBusqueda.RowVisible(i) = Opcion
        End If

    Next i

End Function

Private Sub grdBusqueda_Click(ByVal lRow As Long, ByVal lCol As Long)

    If lCol = 0 Or lRow = 0 Then
        
        Exit Sub
    ElseIf lCol = 1 And lRow > 0 And grdBusqueda.CellIcon(lRow, lCol) = 3 And grdBusqueda.RowVisible(grdBusqueda.SelectedRow) = True Then
        
        grdBusqueda.CellIcon(lRow, lCol) = 4
        MuestraOculta grdBusqueda.CellItemData(lRow, 1), True
    ElseIf lCol = 1 And lRow > 0 Then
        
        grdBusqueda.CellIcon(lRow, lCol) = 3
        MuestraOculta grdBusqueda.CellItemData(lRow, 1), False
    End If

End Sub
