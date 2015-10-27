VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCattipos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de Tipos"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCattipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11175
   Begin VB.Frame Frame2 
      Caption         =   "Caracteristicas Compra / Venta"
      Height          =   975
      Left            =   4680
      TabIndex        =   12
      Top             =   600
      Width           =   6375
      Begin VB.ComboBox cmbUnidad 
         Height          =   315
         ItemData        =   "frmCattipos.frx":000C
         Left            =   4320
         List            =   "frmCattipos.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "IDTipoAlertaMPC"
         Top             =   480
         Width           =   1965
      End
      Begin VB.ComboBox cmbTipoBienes 
         Height          =   315
         ItemData        =   "frmCattipos.frx":0010
         Left            =   120
         List            =   "frmCattipos.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "IDTipoAlertaMPC"
         Top             =   480
         Width           =   4125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
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
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Bien:"
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
         TabIndex        =   13
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Caracteristicas Prestamos"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   4455
      Begin VB.ComboBox cmbTipoGarantia 
         Height          =   315
         ItemData        =   "frmCattipos.frx":0014
         Left            =   120
         List            =   "frmCattipos.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "IDTipoAlertaMPC"
         Top             =   480
         Width           =   4125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Garantía:"
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
         TabIndex        =   15
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.CheckBox chkPeso 
      Appearance      =   0  'Flat
      Caption         =   "Peso"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   195
      Width           =   645
   End
   Begin VB.CheckBox chkKilataje 
      Appearance      =   0  'Flat
      Caption         =   "Kilataje"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   195
      Width           =   885
   End
   Begin VB.TextBox txtTipo 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   600
      TabIndex        =   0
      Top             =   210
      Width           =   2505
   End
   Begin vbAcceleratorGrid6.vbalGrid grdTipos 
      Height          =   5610
      Left            =   0
      TabIndex        =   7
      Top             =   2280
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   9895
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
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   7980
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   3
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
      MaskColor       =   255
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmCattipos.frx":0018
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5940
      TabIndex        =   10
      Top             =   7965
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
      Picture         =   "frmCattipos.frx":011C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
      Height          =   375
      Left            =   9420
      TabIndex        =   6
      Top             =   1740
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
      Object.ToolTipText     =   ""
      Picture         =   "frmCattipos.frx":066E
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
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
      Left            =   90
      TabIndex        =   8
      Top             =   210
      Width           =   405
   End
End
Attribute VB_Name = "frmCattipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAgregar_Click()
Dim IdTipoBien As Integer
Dim IdUnidad As Integer
    
    
    If cmbTipoBienes.ListIndex = -1 Then
        IdTipoBien = 0
    Else
        IdTipoBien = cmbTipoBienes.ItemData(cmbTipoBienes.ListIndex)
    End If
    
    If cmbUnidad.ListIndex = -1 Then
        IdUnidad = 0
    Else
        IdUnidad = cmbUnidad.ItemData(cmbUnidad.ListIndex)
    End If
    
    
    If Trim(txtTipo.text) <> "" Then
        
        If Val(txtTipo.Tag) = 0 Then
            
            If cmbTipoGarantia.ListIndex = -1 Then MsgBox "Seleccione el Tipo de Garantia.", vbCritical, "Catalogo de Tipos": cmbTipoGarantia.SetFocus: Exit Sub
            
            
            dbDatos.Execute "INSERT INTO tipo (Descripcion,Kilataje,Peso,IdTipoGarantia,IdTipoUnidad,IdTipoBienes) VALUES ('" & _
                            Trim(txtTipo.text) & "'," & chkKilataje.Value & "," & chkPeso.Value & "," & Val(cmbTipoGarantia.ItemData(cmbTipoGarantia.ListIndex)) & "," & Val(IdUnidad) & "," & Val(IdTipoBien) & ")"
            Cargar_Tipos
            txtTipo.text = ""
            chkKilataje.Value = 0
            chkPeso.Value = 0
            cmbTipoGarantia.ListIndex = -1
            cmbTipoBienes.ListIndex = -1
            cmbUnidad.ListIndex = -1
            txtTipo.SetFocus
            
        Else

            If MsgBox("Desea guardar los cambios realizados ??", vbQuestion + vbYesNo + vbDefaultButton1, "Catálogo de Tipos") = vbYes Then
                
                If cmbTipoGarantia.ListIndex = -1 Then MsgBox "Seleccione el Tipo de Garantia.", vbCritical, "Catalogo de Tipos": cmbTipoGarantia.SetFocus: Exit Sub
                
                dbDatos.Execute "UPDATE tipo SET Descripcion='" & Trim(txtTipo.text) & "',Kilataje=" & chkKilataje.Value & ",Peso=" & chkPeso.Value & ",IdTipoGarantia=" & Val(cmbTipoGarantia.ItemData(cmbTipoGarantia.ListIndex)) & ",IdTipoUnidad=" & Val(IdUnidad) & ",IdTipoBienes=" & Val(IdTipoBien) & " WHERE ID=" & Val(txtTipo.Tag)
                Cargar_Tipos
                txtTipo.text = ""
                txtTipo.Tag = ""
                chkKilataje.Value = 0
                chkPeso.Value = 0
                cmbTipoGarantia.ListIndex = -1
                cmbTipoBienes.ListIndex = -1
                cmbUnidad.ListIndex = -1
            End If
            
        End If
        
    End If


End Sub

Private Sub cmdLimpiar_Click()
    txtTipo.text = ""
    txtTipo.Tag = ""
    chkKilataje.Value = 0
    chkPeso.Value = 0
    grdTipos.ClearSelection
    txtTipo.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Crear_Encabezado
    
    'MLD-MODIF.
    Cargar_Combos "Descripcion", "mld_prestamos_tipo_garantias", cmbTipoGarantia, " WHERE Estatus=1", , False
    Cargar_Combos "Descripcion", "mld_metales_tipo_unidades", cmbUnidad, " WHERE Estatus=1", , False
    Cargar_Combos "Descripcion", "mld_metales_tipo_bienes", cmbTipoBienes, " WHERE Estatus=1", , False
    
    Cargar_Tipos
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Sub Crear_Encabezado()

    With grdTipos
        
        .AddColumn "C1", "Tipo", ecgHdrTextALignLeft, , 150, , , , , , , CCLSortString
        .AddColumn "C2", "Kilataje", ecgHdrTextALignLeft, , 60, , , , , , , CCLSortString
        .AddColumn "C3", "Peso", ecgHdrTextALignLeft, , 60, , , , , , , CCLSortString
        .AddColumn "C4", "Tipo Garantia", ecgHdrTextALignLeft, , 180, , , , , , , CCLSortString
        .AddColumn "C5", "Tipo Bien", ecgHdrTextALignLeft, , 180, , , , , , , CCLSortString
        .AddColumn "C6", "Unidad", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        
    
    End With

End Sub

Sub Cargar_Tipos()
Dim rcConsulta As New ADODB.Recordset

    rcConsulta.Open "SELECT * FROM tipo ORDER BY ID", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        
        rcConsulta.MoveFirst
        With grdTipos
            .Clear
            While Not rcConsulta.EOF
                .AddRow
                .CellText(.Rows, 1) = rcConsulta!Descripcion
                .CellItemData(.Rows, 1) = rcConsulta!ID
                .CellText(.Rows, 2) = IIf(rcConsulta!Kilataje, "Si", "No")
                .CellText(.Rows, 3) = IIf(rcConsulta!Peso, "Si", "No")
                .CellText(.Rows, 4) = IIf(IsNull(rcConsulta!IdTipoGarantia), "", UCase(SacaValor("mld_prestamos_tipo_garantias", "Descripcion", " WHERE Id=" & rcConsulta!IdTipoGarantia)))
                .CellText(.Rows, 5) = IIf(IsNull(rcConsulta!IdTipoBienes), "", UCase(SacaValor("mld_metales_tipo_bienes", "Descripcion", " WHERE Id=" & rcConsulta!IdTipoBienes)))
                .CellText(.Rows, 6) = IIf(IsNull(rcConsulta!IdTipoUnidad), "", UCase(SacaValor("mld_metales_tipo_unidades", "Descripcion", " WHERE Id=" & rcConsulta!IdTipoUnidad)))
                
            rcConsulta.MoveNext
            Wend
        End With

    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
    
End Sub

Private Sub grdTipos_DblClick(ByVal lRow As Long, ByVal lCol As Long)
Dim rcAux As New ADODB.Recordset

On Error GoTo Error

    If grdTipos.Rows > 0 And grdTipos.SelectedRow > 0 Then
        
        If MsgBox("Desea editar el tipo seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton1, "Catálogo de Tipos") = vbYes Then
            
            rcAux.Open "SELECT * FROM tipo WHERE ID=" & Val(grdTipos.CellItemData(grdTipos.SelectedRow, 1)), dbDatos, adOpenForwardOnly, adLockOptimistic
            If Not rcAux.BOF And Not rcAux.EOF Then
                txtTipo.text = rcAux!Descripcion
                txtTipo.Tag = rcAux!ID
                chkKilataje.Value = IIf(rcAux!Kilataje, 1, 0)
                chkPeso.Value = IIf(rcAux!Peso, 1, 0)
                
                cmbTipoGarantia.ListIndex = ComboInformacion(cmbTipoGarantia, rcAux!IdTipoGarantia)
                cmbUnidad.ListIndex = ComboInformacion(cmbUnidad, rcAux!IdTipoUnidad)
                cmbTipoBienes.ListIndex = ComboInformacion(cmbTipoBienes, rcAux!IdTipoBienes)
                
            End If
            rcAux.Close
            
        Else
            
            grdTipos.ClearSelection
            txtTipo.SetFocus
        End If
    
    End If

Error:
    Maneja_Error Err
    Set rcAux = Nothing
End Sub

Private Sub grdTipos_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyDelete Then
        
        If grdTipos.Rows > 0 And grdTipos.SelectedRow > 0 Then
            
            If MsgBox("Desea eliminar el tipo: " & Trim(grdTipos.CellText(grdTipos.SelectedRow, 1)) & "", vbQuestion + vbYesNo + vbDefaultButton2, "Catálogo de Tipos") = vbYes Then
                
                dbDatos.Execute "UPDATE tipo SET Estatus=0 WHERE ID=" & Val(grdTipos.CellItemData(grdTipos.SelectedRow, 1))
                Cargar_Tipos
                txtTipo.SetFocus
            
            End If
        
        End If
    
        grdTipos.ClearSelection
        txtTipo.SetFocus
    End If

End Sub

Private Sub txttipo_GotFocus()
    Seleccionar_Texto txtTipo
    Cambiar_Color True, txtTipo
End Sub

Private Sub txttipo_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txttipo_LostFocus()
    Cambiar_Color False, txtTipo
End Sub


'.MLD_MODIF
Private Sub cmbTipoGarantia_GotFocus()
    Cambiar_Color True, cmbTipoGarantia
End Sub

Private Sub cmbTipoGarantia_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoGarantia_LostFocus()
    Cambiar_Color False, cmbTipoGarantia
End Sub

