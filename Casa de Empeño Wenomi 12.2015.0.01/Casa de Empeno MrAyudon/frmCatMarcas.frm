VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCatMarcas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de Marcas"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatMarcas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   4515
   Begin VB.TextBox txtDescripcion 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   3270
   End
   Begin vbAcceleratorGrid6.vbalGrid grdMarcas 
      Height          =   5565
      Left            =   15
      TabIndex        =   2
      Top             =   675
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9816
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
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
      Left            =   2145
      TabIndex        =   3
      Top             =   6315
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
      Picture         =   "frmCatMarcas.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   6315
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
      Picture         =   "frmCatMarcas.frx":0110
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
      Default         =   -1  'True
      Height          =   375
      Left            =   3375
      TabIndex        =   1
      Top             =   285
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
      Picture         =   "frmCatMarcas.frx":0662
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmCatMarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAgregar_Click()

    If Trim(txtDescripcion.text) <> "" Then
        
        If Val(txtDescripcion.Tag) = 0 Then
            
            dbDatos.Execute "INSERT INTO marcas (Descripcion) VALUES ('" & _
                            Trim(txtDescripcion.text) & "')"
            Cargar_Marcas
            txtDescripcion.text = ""
            txtDescripcion.SetFocus
            
        Else

            If MsgBox("Desea guardar los cambios realizados ??", vbQuestion + vbYesNo + vbDefaultButton1, "Catálogo de Marcas") = vbYes Then
                
                dbDatos.Execute "UPDATE marcas SET Descripcion='" & Trim(txtDescripcion.text) & "' WHERE ID=" & Val(txtDescripcion.Tag)
                Cargar_Marcas
                txtDescripcion.text = ""
                txtDescripcion.Tag = ""
                
            End If
            
        End If
        
    End If

End Sub

Private Sub cmdLimpiar_Click()
    txtDescripcion.text = ""
    txtDescripcion.Tag = ""
    grdMarcas.ClearSelection
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
    Crear_Encabezado
    Cargar_Marcas
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Sub Crear_Encabezado()

    With grdMarcas
        
        .AddColumn "C1", "Descripción", ecgHdrTextALignLeft, , 272, , , , , , , CCLSortString
    
    End With

End Sub

Sub Cargar_Marcas()
Dim rcConsulta As New ADODB.Recordset

    rcConsulta.Open "SELECT * FROM marcas ORDER BY ID", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        
        rcConsulta.MoveFirst
        With grdMarcas
            .Clear
            While Not rcConsulta.EOF
                .AddRow
                .CellText(.Rows, 1) = rcConsulta!Descripcion
                .CellItemData(.Rows, 1) = rcConsulta!ID
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

Private Sub grdMarcas_DblClick(ByVal lRow As Long, ByVal lCol As Long)
Dim rcAux As New ADODB.Recordset

On Error GoTo error

    If grdMarcas.Rows > 0 And grdMarcas.SelectedRow > 0 Then
        
        If MsgBox("Desea editar la marca seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton1, "Catálogo de Marcas") = vbYes Then
            
            rcAux.Open "SELECT * FROM marcas WHERE ID=" & Val(grdMarcas.CellItemData(grdMarcas.SelectedRow, 1)), dbDatos, adOpenForwardOnly, adLockOptimistic
            If Not rcAux.BOF And Not rcAux.EOF Then
                txtDescripcion.text = rcAux!Descripcion
                txtDescripcion.Tag = rcAux!ID
            End If
            rcAux.Close
            
        Else
            
            grdMarcas.ClearSelection
            txtDescripcion.SetFocus
        End If
    
    End If

error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Sub

Private Sub grdMarcas_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyDelete Then
        
        If grdMarcas.Rows > 0 And grdMarcas.SelectedRow > 0 Then
            
            If MsgBox("Desea eliminar la marca: " & Trim(grdMarcas.CellText(grdMarcas.SelectedRow, 1)) & "", vbQuestion + vbYesNo + vbDefaultButton2, "Catálogo de Marcas") = vbYes Then
                
                dbDatos.Execute "DELETE FROM marcas WHERE ID=" & Val(grdMarcas.CellItemData(grdMarcas.SelectedRow, 1))
                Cargar_Marcas
                txtDescripcion.SetFocus
            
            End If
        
        End If
    
        grdMarcas.ClearSelection
        txtDescripcion.SetFocus
    End If

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
