VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCatTipoPrenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo Prendas Joyería"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatTipoPrenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   5430
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmCatTipoPrenda.frx":000C
      Left            =   1230
      List            =   "frmCatTipoPrenda.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2850
   End
   Begin VB.TextBox txtDescripcion 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1260
      TabIndex        =   1
      Top             =   600
      Width           =   2835
   End
   Begin vbAcceleratorGrid6.vbalGrid grdPrendas 
      Height          =   6480
      Left            =   60
      TabIndex        =   3
      Top             =   915
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   11430
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
   Begin DevPowerFlatBttn.FlatBttn cmdCancelar 
      Height          =   375
      Left            =   1905
      TabIndex        =   4
      Top             =   7455
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
      Picture         =   "frmCatTipoPrenda.frx":003E
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4275
      TabIndex        =   7
      Top             =   7455
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
      Picture         =   "frmCatTipoPrenda.frx":0142
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   465
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
      Picture         =   "frmCatTipoPrenda.frx":0694
   End
   Begin DevPowerFlatBttn.FlatBttn cmdEliminar 
      Height          =   375
      Left            =   3090
      TabIndex        =   8
      Top             =   7455
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Eliminar"
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
      Picture         =   "frmCatTipoPrenda.frx":0BE6
      PictureDisabled =   "frmCatTipoPrenda.frx":1138
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   30
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmCatTipoPrenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbTipo_GotFocus()
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipo_LostFocus()
    Cambiar_Color False, cmbTipo
End Sub

Private Sub cmdAceptar_Click()

    If cmbTipo.ListIndex > -1 And Trim(txtDescripcion.text) <> "" And Val(txtDescripcion.Tag) = 0 Then
        
        dbDatos.Execute "INSERT INTO tipoprenda (Descripcion,IDTipo) VALUES ('" & Trim(txtDescripcion.text) & "'," & cmbTipo.ItemData(cmbTipo.ListIndex) & ")"
        CargarPrendas
        txtDescripcion.text = ""
        txtDescripcion.Tag = ""
        cmbTipo.ListIndex = 0
        
    ElseIf Val(txtDescripcion.Tag) > 0 Then
        
        dbDatos.Execute "UPDATE tipoprenda SET Descripcion='" & Trim(txtDescripcion.text) & "',IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " WHERE ID=" & Val(txtDescripcion.Tag)
        CargarPrendas
        txtDescripcion.text = ""
        txtDescripcion.Tag = ""
        cmbTipo.ListIndex = 0
    End If
End Sub

Private Sub cmdEliminar_Click()

    If grdPrendas.Rows > 0 Then
        
        If grdPrendas.SelectedRow > 0 Then
            
            If MsgBox("Desea eliminar la prenda: " & Trim(grdPrendas.CellText(grdPrendas.SelectedRow, 2)), vbQuestion + vbYesNo + vbDefaultButton2, "Prendas Joyería") = vbYes Then
                
                dbDatos.Execute "DELETE FROM tipoprenda WHERE ID=" & grdPrendas.CellItemData(grdPrendas.SelectedRow, 2)
                CargarPrendas
                txtDescripcion.text = ""
                txtDescripcion.Tag = ""
                txtDescripcion.SetFocus
            End If

        Else
            
            txtDescripcion.SetFocus
        End If

    Else
        
        txtDescripcion.SetFocus
    End If

End Sub

Private Sub cmdCancelar_Click()
    txtDescripcion.text = ""
    txtDescripcion.Tag = ""
    cmbTipo.ListIndex = 0
    grdPrendas.ClearSelection
    txtDescripcion.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Cargar_Combos "Descripcion", "tipo", cmbTipo, " WHERE Kilataje=1 OR Peso=1", "Ordenamiento"
    cmbTipo.ListIndex = 0
    Crear_Encabezado
    CargarPrendas
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Sub Crear_Encabezado()

    With grdPrendas
        
        .AddColumn "C1", "Tipo", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
        .AddColumn "C2", "Descripción", ecgHdrTextALignLeft, , 225, , , , , , , CCLSortString
    End With

End Sub

Sub CargarPrendas()
Dim rcTmp As New ADODB.Recordset

On Error GoTo error

    rcTmp.Open "SELECT tipoprenda.ID,tipoprenda.Descripcion,tipoprenda.IDTipo,tipo.Descripcion AS Tipo " _
                & "FROM tipoprenda INNER JOIN tipo ON tipoprenda.IDTipo=tipo.ID WHERE tipo.Kilataje=1 OR tipo.Peso=1 ORDER BY tipoprenda.IDTipo,tipoprenda.Descripcion", dbDatos, adOpenForwardOnly, adLockOptimistic
    
    If Not rcTmp.BOF And Not rcTmp.EOF Then
        rcTmp.MoveFirst
        With grdPrendas
            .Clear
            While Not rcTmp.EOF
                .AddRow
                .CellText(.Rows, 1) = rcTmp!Tipo
                .CellItemData(.Rows, 1) = rcTmp!IDTipo
                .CellText(.Rows, 2) = rcTmp!Descripcion
                .CellItemData(.Rows, 2) = rcTmp!ID
            rcTmp.MoveNext
            Wend
        End With

    End If
    rcTmp.Close
    Set rcTmp = Nothing
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdPrendas_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    
    With grdPrendas
        
        If .Rows > 0 Then
                
            If .SelectedRow > 0 Then
                
                cmbTipo.ListIndex = ComboInformacion(cmbTipo, .CellItemData(lRow, 1))
                txtDescripcion.text = .CellText(.SelectedRow, 2)
                txtDescripcion.Tag = .CellItemData(.SelectedRow, 2)
                grdPrendas.ClearSelection
                txtDescripcion.SetFocus
            
            End If
            
        End If
        
    End With

End Sub

Private Sub txtDescripcion_GotFocus()
    Seleccionar_Texto txtDescripcion
    Cambiar_Color True, txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDescripcion_LostFocus()
    Cambiar_Color False, txtDescripcion
End Sub
