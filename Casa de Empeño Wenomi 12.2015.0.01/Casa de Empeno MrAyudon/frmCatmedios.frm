VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCatmedios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medios"
   ClientHeight    =   6750
   ClientLeft      =   3465
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
   Icon            =   "frmCatmedios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   4515
   Begin VB.TextBox txtMedio 
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
      Width           =   3315
   End
   Begin vbAcceleratorGrid6.vbalGrid grdMedios 
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
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
      DefaultRowHeight=   18
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   3420
      TabIndex        =   3
      Top             =   285
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmCatmedios.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3450
      TabIndex        =   4
      Top             =   6315
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmCatmedios.frx":055E
   End
   Begin DevPowerFlatBttn.FlatBttn cmdEliminar 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   6315
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmCatmedios.frx":0AB0
      PictureDisabled =   "frmCatmedios.frx":1002
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
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmCatmedios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()
    
    If Trim(txtMedio.text) <> "" Then
        
        dbDatos.Execute "INSERT INTO medios (Descripcion) VALUES ('" & _
                        Trim(txtMedio.text) & "')"
        Cargar_Medios
        txtMedio.text = ""
    End If

End Sub

Private Sub cmdEliminar_Click()

    If grdMedios.SelectedRow > 0 Then
    
        If MsgBox("Desea eliminar el medio: " & Trim(grdMedios.CellText(grdMedios.SelectedRow, 1)) & "", vbQuestion + vbYesNo + vbDefaultButton2, "Medios") = vbYes Then
            
            dbDatos.Execute "DELETE FROM medios WHERE ID=" & grdMedios.CellItemData(grdMedios.SelectedRow, 1)
            Cargar_Medios
        End If
    
    End If
    
    grdMedios.ClearSelection
    txtMedio.SetFocus
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
    Cargar_Medios
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Sub Crear_Encabezado()
    
    With grdMedios
        
        .AddColumn "C1", "Medio", ecgHdrTextALignLeft, , 272, , , , , , , CCLSortString
    End With

End Sub

Sub Cargar_Medios()
Dim rcMedios As New ADODB.Recordset

On Error GoTo error

    rcMedios.Open "SELECT * FROM medios ORDER BY Descripcion", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcMedios.BOF And Not rcMedios.EOF Then
        
        rcMedios.MoveFirst
        With grdMedios
        
            .Redraw = False
            .Clear
            While Not rcMedios.EOF
                .AddRow
                .CellText(.Rows, 1) = rcMedios!Descripcion
                .CellItemData(.Rows, 1) = rcMedios!ID
            rcMedios.MoveNext
            Wend
            .Redraw = True
            
        End With
    
    End If
    rcMedios.Close
    Set rcMedios = Nothing
    
error:
    Maneja_Error Err
    Set rcMedios = Nothing
End Sub

Private Sub txtMedio_GotFocus()
    Seleccionar_Texto txtMedio
    Cambiar_Color True, txtMedio
End Sub

Private Sub txtMedio_KeyPress(KeyAscii As Integer)
    KeyAscii = mayusculas(KeyAscii)
End Sub

Private Sub txtMedio_LostFocus()
    Cambiar_Color False, txtMedio
End Sub
