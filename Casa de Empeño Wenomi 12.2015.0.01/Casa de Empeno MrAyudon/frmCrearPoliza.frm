VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.2#0"; "vbalSGrid6.ocx"
Begin VB.Form frmCrearPoliza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Poliza"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCrearPoliza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   10620
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFecha 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   360
      Width           =   1440
   End
   Begin VB.TextBox txtConcepto 
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
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   8055
   End
   Begin VB.TextBox txtNoPoliza 
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
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sumas Iguales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   6
      Top             =   6840
      Width           =   4575
      Begin VB.Label lblCargo 
         Alignment       =   1  'Right Justify
         Caption         =   "<Cargo>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2010
      End
      Begin VB.Label lblAbono 
         Alignment       =   1  'Right Justify
         Caption         =   "<Abono>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   2010
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Poliza"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8280
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton opDiario 
         Appearance      =   0  'Flat
         Caption         =   "&Diario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton opEgreso 
         Appearance      =   0  'Flat
         Caption         =   "&Egreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton opIngreso 
         Appearance      =   0  'Flat
         Caption         =   "&Ingreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdPoliza 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8916
      GridLines       =   -1  'True
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
      HeaderHeight    =   18
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
      HighlightSelectedIcons=   0   'False
      DefaultRowHeight=   18
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Left            =   4200
      TabIndex        =   13
      Top             =   360
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
      Picture         =   "frmCrearPoliza.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   9330
      TabIndex        =   16
      Top             =   7680
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
      Picture         =   "frmCrearPoliza.frx":0121
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCrearPoliza 
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   7680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "&Crear Poliza"
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
      TextColor       =   8537065
      Object.ToolTipText     =   ""
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   8160
      TabIndex        =   18
      Top             =   7680
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      TextColor       =   8537065
      Object.ToolTipText     =   ""
      Picture         =   "frmCrearPoliza.frx":0673
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
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
      Left            =   2760
      TabIndex        =   14
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Concepto:"
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
      TabIndex        =   10
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Numero de Poliza:"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1725
   End
End
Attribute VB_Name = "frmCrearPoliza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cIngresos = 1
Private Const cEgresos = 2
Private Const cDiario = 3
Private Const cDeOrden = 4
Private Const cEstadisticas = 5

Private Const cNormal = 1
Private Const cSinAfectar = 2

Private Const cCargo = 1
Private Const cAbono = 2


Private fl() As New cFlatControl

Private m_Cargo As Currency
Private m_Abono As Currency

Private Property Let Cargo(Valor As Currency)
   m_Cargo = Valor
End Property

Private Property Get Cargo() As Currency
   Cargo = m_Cargo
End Property

Private Property Let Abono(Valor As Currency)
   m_Abono = Valor
End Property

Private Property Get Abono() As Currency
   Abono = m_Abono
End Property

Private Sub cmdCrearPoliza_Click()
   ExportarPoliza
End Sub

Private Sub cmdMosFecha_Click()
   txtFecha.text = frmCalendario.Fecha(Now)
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdAceptar_Click()
   GrabarPoliza
   Limpiar
End Sub

Private Sub GrabarPoliza()
   Dim Indice As Long
   
   For Indice = 1 To grdPoliza.Rows - 1
      
      With grdPoliza
      
         dbDatos.Execute "INSERT INTO Polizas (Fecha,Entrada,Cuenta,Concepto,Cargo,Abono,PC) VALUES ('" & _
                         Format(txtFecha.text, "YYYY/MM/DD") & "'," & IIf(.CellIcon(Indice, 1) = 4, 1, 0) & ",'" & _
                         .CellText(Indice, 2) & "','" & .CellText(Indice, 3) & "'," & Val(.CellText(Indice, 4)) & "," & Val(.CellText(Indice, 5)) & ",'" & NombrePc & "')"
      
      End With
      
      
   Next Indice
   
   MsgBox "La Poliza se ha guardado correctamente", vbInformation Or vbOKOnly
      
End Sub

Private Sub Form_Load()
   Inicializar
End Sub

Private Sub Inicializar()
   CentrarForm Me, frmMDI
   CrearEncabezados
   Poner_Flat fl, Me.Controls, Me
   Limpiar
End Sub

Private Sub CrearEncabezados()
   With grdPoliza
      .ImageList = frmMDI.img
      .AddColumn "K1", "Entrada", ecgHdrTextALignCentre, , 60, , , , False
      .AddColumn "K2", "Cuenta", ecgHdrTextALignLeft, , 119, , , , , , , CCLSortString
      .AddColumn "K3", "Concepto", ecgHdrTextALignLeft, , 311, , , , , , , CCLSortString
      .AddColumn "K4", "Cargo", ecgHdrTextALignRight, , 87, , , , , , , CCLSortString
      .AddColumn "K5", "Abono", ecgHdrTextALignRight, , 87, , , , , , , CCLSortString
      .AddRow
      .CellIndent(.Rows, 1) = 20
      .CellIcon(.Rows, 1) = 3
   End With
End Sub

Private Sub Limpiar()
   lblCargo.Caption = ""
   lblAbono.Caption = ""
   txtConcepto.text = ""
   txtFecha.text = Format(Now, "dd/mm/yyyy")
   grdPoliza.Clear
   grdPoliza.AddRow
   grdPoliza.CellIndent(grdPoliza.Rows, 1) = 20
   grdPoliza.CellIcon(grdPoliza.Rows, 1) = 3
End Sub

Private Sub grdPoliza_CancelEdit()
   txtEdit.Visible = False
End Sub

Private Sub grdPoliza_Click(ByVal lRow As Long, ByVal lCol As Long)
   If lRow > 0 And lCol = 1 Then
      If grdPoliza.CellIcon(lRow, lCol) = 3 Then
         grdPoliza.CellIcon(lRow, lCol) = 4
      Else
         grdPoliza.CellIcon(lRow, lCol) = 3
      End If
   End If
End Sub

Private Sub grdPoliza_ColumnWidthChanged(ByVal lCol As Long, lWidth As Long, bCancel As Boolean)
   grdPoliza.CellText(1, lCol) = lWidth
End Sub

Private Sub grdPoliza_PreCancelEdit(ByVal lRow As Long, ByVal lCol As Long, newValue As Variant, bStayInEditMode As Boolean)
   If (txtEdit.text = "") Then
      'Status = "Enter some text."
      ' This would be a good place for a popup message bubble
      ' either use the OS or use a VB window that's
      ' transparent to the mouse by subclassing WM_NCHITTEST = HT_NOWHERE
      'MsgBox "Please enter some text into the cell.", vbExclamation
      bStayInEditMode = True
   Else
      'Status = "Ready"
      If lCol < 4 Then
         'grdPoliza.CellText(grdPoliza.EditRow, grdPoliza.EditCol) = txtEdit.text
         grdPoliza.CellDetails grdPoliza.EditRow, grdPoliza.EditCol, txtEdit.text, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
      Else
         grdPoliza.CellDetails grdPoliza.EditRow, grdPoliza.EditCol, txtEdit.text, DT_RIGHT, , , , , 4
         SumasIguales
      End If
      Agregar_Renglon (grdPoliza.EditRow)
   End If
End Sub


Private Sub Agregar_Renglon(Row As Long)
   
   With grdPoliza
   
      If .CellText(Row, 2) <> "" And .CellText(Row, 3) <> "" And (.CellText(Row, 4) <> "" Or .CellText(Row, 5) <> "") Then
         If .CellText(.Rows, 2) <> "" And .CellText(.Rows, 3) <> "" And (.CellText(.Rows, 4) <> "" Or .CellText(.Rows, 5) <> "") Then
            .AddRow
            .SelectedCol = 2
            .SelectedRow = .Rows
            .CellIndent(.Rows, 1) = 20
            .CellIcon(.Rows, 1) = 3
         End If
      End If
      
   End With
   
   
End Sub

Private Sub grdPoliza_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
   Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
   Dim sText As String
   
   'si el el icono salimos del modo de edicion
   If lCol = 1 Then
      bCancel = True
      Exit Sub
   End If
   
   
   'si es cargo o abono checamos que no haya una cantidad en la otra celda
   If lCol = 4 Or lCol = 5 Then
      
      'si hay una cantidad en cargo y se kiere editar abono se sale
      If grdPoliza.CellText(lRow, 4) <> "" And lCol = 5 Then
         bCancel = True
         Exit Sub
      End If
      
      'si hay una cantidad en abono y se kiere editar cargo se sale
      If grdPoliza.CellText(lRow, 5) <> "" And lCol = 4 Then
         bCancel = True
         Exit Sub
      End If
   End If
   
   
   
   ' Don't allow editing the icon-only columns:
'   If (grdPoliza.ColumnKey(lCol) = "file") Or (grdPoliza.ColumnKey(lCol) = "col8") Then
'      bCancel = True
'      Exit Sub
'   End If
   
   ' Get boundary of the cell:
   grdPoliza.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
   
   ' Get the text:
   If Not IsMissing(grdPoliza.CellText(lRow, lCol)) Then
      sText = grdPoliza.CellFormattedText(lRow, lCol)
   Else
      sText = ""
   End If
   
   ' If the user has initiated edit mode by a key, we want
   ' to add this to the text.  This is really a common
   ' thing and should probably be supported automatically
   ' in the grid:
   If Not (iKeyAscii = 0) Then
      sText = Chr$(iKeyAscii) '& sText
      txtEdit.text = sText
      txtEdit.SelStart = 1
      txtEdit.SelLength = Len(sText)
   Else
      txtEdit.text = sText
      txtEdit.SelStart = 0
      txtEdit.SelLength = Len(sText)
   End If
   
   ' Set the text properties to match the grid cell being edited:
   Set txtEdit.Font = grdPoliza.CellFont(lRow, lCol)
   If grdPoliza.CellBackColor(lRow, lCol) = -1 Then
      txtEdit.BackColor = grdPoliza.BackColor
   Else
      txtEdit.BackColor = grdPoliza.CellBackColor(lRow, lCol)
   End If
   
   ' Move the text box to the edit position, make it visible and give it the focus:
   If lCol = 2 Then
      txtEdit.Move lLeft + grdPoliza.Left + 25, lTop + grdPoliza.Top + Screen.TwipsPerPixelY, lWidth, lHeight
   Else
      txtEdit.Move lLeft + grdPoliza.Left, lTop + grdPoliza.Top + Screen.TwipsPerPixelY, lWidth, lHeight
   End If
   
   If lCol = 2 Then
      frmMostrarCuentas.Ver Me, txtEdit
   Else
      txtEdit.Visible = True
      txtEdit.ZOrder
      txtEdit.SetFocus
   End If
   
   
   
End Sub

Private Sub txtEdit_GotFocus()
   Cambiar_Color True, txtEdit
   'Seleccionar_Texto txtEdit
End Sub

Public Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
   If (KeyCode = vbKeyReturn) Then
      ' Request Commit edit.  This will fire the
      ' grid's PreCancelEdit event, which gives you
      ' an opportunity to validate the data and put
      ' it in the cell if good.  The CancelEdit
      ' event will then fire afterwards.
      grdPoliza.EndEdit
   ElseIf (KeyCode = vbKeyEscape) Then
      ' Cancel edit.  This skips PreCancelEdit and
      ' fires the CancelEdit event
      grdPoliza.CancelEdit
   ElseIf (grdPoliza.SingleClickEdit) Then
      Select Case KeyCode
      
      End Select
   End If
End Sub

Private Sub txtEdit_LostFocus()
   Cambiar_Color False, txtEdit
End Sub

Private Sub SumasIguales()
   Dim Indice As Long
   
   Cargo = 0
   Abono = 0
   
   For Indice = 1 To grdPoliza.Rows
      Cargo = Cargo + Val(grdPoliza.CellText(Indice, 4))
      Abono = Abono + Val(grdPoliza.CellText(Indice, 5))
   Next Indice
   
   
   lblCargo.Caption = Format(Cargo, FMoneda)
   lblAbono.Caption = Format(Abono, FMoneda)
   
   If Cargo = Abono Then
      lblCargo.BackColor = vbGreen
      lblAbono.BackColor = vbGreen
   ElseIf Cargo > Abono Then
      lblCargo.BackColor = vbGreen
      lblAbono.BackColor = vbRed
   ElseIf Abono > Cargo Then
      lblCargo.BackColor = vbRed
      lblAbono.BackColor = vbGreen
   End If
   

End Sub

'Procedimientos para grabar la poliza

Private Sub ExportarPoliza()
   Dim strArchivo As String
   Dim lArchivo As Long
   Dim Indice As Long
   
   If opDiario.Value Then
      strArchivo = "PD-" & txtNoPoliza.text
   ElseIf opIngreso.Value Then
      strArchivo = "PI-" & txtNoPoliza.text
   ElseIf opEgreso.Value Then
      strArchivo = "PE-" & txtNoPoliza.text
   End If
   
   strArchivo = strArchivo & ".txt"
   
   lArchivo = Grabar_Archivo(strArchivo)
   Crear_Encabezado lArchivo, txtNoPoliza.text, 1, 1, 1
   
   With grdPoliza
   
      For Indice = 1 To .Rows - 1
         Escribe_Archivo Replace(.CellText(Indice, 2), "-", ""), txtNoPoliza.text, Val(.CellText(Indice, 4)), Val(.CellText(Indice, 5)), .CellText(Indice, 3), CInt(lArchivo)
      Next Indice
   
   End With
   
   Close #lArchivo
   
   MsgBox "La Poliza se ha creado correctamente", vbOKOnly + vbInformation
End Sub

'Creamos el archivo de la poliza
Private Function Grabar_Archivo(Archivo As String) As Long
   Dim strArchivo As String
   Dim iArchivo As Long
   
   'strArchivo = "C:\Polizas\PI" & Poliza & ".txt"
   strArchivo = App.Path & "\Polizas\" & Archivo
   
   If Dir(App.Path & "\Polizas", vbDirectory) = "" Then MkDir App.Path & "\Polizas"
   
   iArchivo = FreeFile
   Open strArchivo For Output Access Write As #iArchivo
   Grabar_Archivo = iArchivo
      
   
End Function

'Creamos el encabezado del archivo
Private Sub Crear_Encabezado(iArchivo As Long, Poliza As String, RenIni As Long, RenFin As Long, Colini As Long)
   Dim Encabezado As String
   Dim Fecha As String
   Dim DiarioAgrupador As String
   Dim strPoliza As String
   Dim Concepto As String
   Dim Diario As String
   Dim Sistema As String
   
   strPoliza = Val(Poliza)
   strPoliza = Mid("00000000", 1, 8 - Len(strPoliza)) & strPoliza 'asignamos todos los espacios a la poliza
   
   Fecha = Format(txtFecha.text, "YYYYMMDD") 'Format(Regresa_Fecha(RenIni, RenFin, Colini), "YYYYMMDD")
   Fecha = Fecha & Space(8 - Len(Fecha))
   
   Concepto = txtConcepto.text 'Regresa_Concepto(RenIni, RenFin, Colini)
   Concepto = Concepto & Space(100 - Len(Concepto))
   Diario = "000"
   Sistema = "10"
   
   Encabezado = "P " & Fecha & " " & Format(IIf(opIngreso.Value, cIngresos, IIf(opEgreso.Value, cEgresos, cDiario)), "000") & " " & strPoliza & " " & cNormal & " " & Diario & " " & Concepto & " " & Sistema & " " & "2" & " "
   
   Print #iArchivo, Encabezado
End Sub

'Escribimos el archivo
Public Sub Escribe_Archivo(Cta As String, Poliza As String, Cargo As Currency, Abono As Currency, Descripcion As String, Archivo As Integer)
   Dim strCta As String
   Dim strPoliza As String
   Dim strImporte As String
   Dim strCadena As String
   Dim TipoCargoAbono As Integer
   Const Movimiento = "M"
      
   strCta = Cta
   strCta = strCta & Space(20 - Len(strCta))
   strPoliza = IIf(opIngreso.Value, "PI-", IIf(opEgreso.Value, "PE-", "PD-")) & Poliza
   strPoliza = strPoliza & Space(10 - Len(strPoliza))
   If Cargo > 0 Then
      strImporte = Quitar_Simbolos(Format(Cargo, "Currency"))
      TipoCargoAbono = cCargo
   ElseIf Abono > 0 Then
      strImporte = Quitar_Simbolos(Format(Abono, "Currency"))
      TipoCargoAbono = cAbono
   End If
   
   strImporte = Space(16 - Len(strImporte)) & strImporte
   Descripcion = Mid(Descripcion, 1, 30)
   Descripcion = Descripcion & Space(30 - Len(Descripcion))
   strCadena = Movimiento & " " & strCta & " " & strPoliza & " " & TipoCargoAbono & " " & strImporte & Space(22) & Descripcion & " " '& Chr(10) & Chr(13)
   Print #Archivo, strCadena
End Sub

'quita todo los simbolos  monetarios
Public Function Quitar_Simbolos(Cadena As String) As String
   Quitar_Simbolos = Format(Cadena, "#########0.00")
End Function
