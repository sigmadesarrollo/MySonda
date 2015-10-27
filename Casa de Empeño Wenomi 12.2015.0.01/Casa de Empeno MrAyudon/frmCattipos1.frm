VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCattipos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de Tipos"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
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
   ScaleHeight     =   6165
   ScaleWidth      =   4350
   Begin VB.CheckBox chkPeso 
      Appearance      =   0  'Flat
      Caption         =   "Peso"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   645
   End
   Begin VB.CheckBox chkKilataje 
      Appearance      =   0  'Flat
      Caption         =   "Kilataje"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
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
      Height          =   4650
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   8202
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
      Left            =   1920
      TabIndex        =   6
      Top             =   5700
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
      Picture         =   "frmCattipos.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3180
      TabIndex        =   7
      Top             =   5685
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
      Picture         =   "frmCattipos.frx":0110
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
      Height          =   375
      Left            =   3180
      TabIndex        =   1
      Top             =   150
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
      Picture         =   "frmCattipos.frx":0662
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
      TabIndex        =   3
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

Dim Fl() As cFlatControl

Private Sub cmdAgregar_Click()

    If Trim(txtTipo.text) <> "" Then
        
        If Val(txtTipo.Tag) = 0 Then
            
            dbDatos.Execute "INSERT INTO tipo (Descripcion,Kilataje,Peso) VALUES ('" & _
                            Trim(txtTipo.text) & "'," & chkKilataje.Value & "," & chkPeso.Value & ")"
            Cargar_Tipos
            txtTipo.text = ""
            chkKilataje.Value = 0
            chkPeso.Value = 0
            txtTipo.SetFocus
            
        Else

            If MsgBox("Desea guardar los cambios realizados ??", vbQuestion + vbYesNo + vbDefaultButton1, "Catálogo de Tipos") = vbYes Then
                
                dbDatos.Execute "UPDATE tipo SET Descripcion='" & Trim(txtTipo.text) & "',Kilataje=" & chkKilataje.Value & ",Peso=" & chkPeso.Value & " WHERE ID=" & Val(txtTipo.Tag)
                Cargar_Tipos
                txtTipo.text = ""
                txtTipo.Tag = ""
                chkKilataje.Value = 0
                chkPeso.Value = 0
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
    Cargar_Tipos
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Sub Crear_Encabezado()

    With grdTipos
        
        .AddColumn "C1", "Tipo", ecgHdrTextALignLeft, , 150, , , , , , , CCLSortString
        .AddColumn "C2", "Kilataje", ecgHdrTextALignLeft, , 60, , , , , , , CCLSortString
        .AddColumn "C3", "Peso", ecgHdrTextALignLeft, , 60, , , , , , , CCLSortString
    
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
            rcConsulta.MoveNext
            Wend
        End With

    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdTipos_DblClick(ByVal lRow As Long, ByVal lCol As Long)
Dim rcAux As New ADODB.Recordset

On Error GoTo error

    If grdTipos.Rows > 0 And grdTipos.SelectedRow > 0 Then
        
        If MsgBox("Desea editar el tipo seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton1, "Catálogo de Tipos") = vbYes Then
            
            rcAux.Open "SELECT * FROM tipo WHERE ID=" & Val(grdTipos.CellItemData(grdTipos.SelectedRow, 1)), dbDatos, adOpenForwardOnly, adLockOptimistic
            If Not rcAux.BOF And Not rcAux.EOF Then
                txtTipo.text = rcAux!Descripcion
                txtTipo.Tag = rcAux!ID
                chkKilataje.Value = IIf(rcAux!Kilataje, 1, 0)
                chkPeso.Value = IIf(rcAux!Peso, 1, 0)
            End If
            rcAux.Close
            
        Else
            
            grdTipos.ClearSelection
            txtTipo.SetFocus
        End If
    
    End If

error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Sub

Private Sub grdTipos_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyDelete Then
        
        If grdTipos.Rows > 0 And grdTipos.SelectedRow > 0 Then
            
            If MsgBox("Desea eliminar el tipo: " & Trim(grdTipos.CellText(grdTipos.SelectedRow, 1)) & "", vbQuestion + vbYesNo + vbDefaultButton2, "Catálogo de Tipos") = vbYes Then
                
                dbDatos.Execute "DELETE FROM tipo WHERE ID=" & Val(grdTipos.CellItemData(grdTipos.SelectedRow, 1))
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
    KeyAscii = mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txttipo_LostFocus()
    Cambiar_Color False, txtTipo
End Sub
