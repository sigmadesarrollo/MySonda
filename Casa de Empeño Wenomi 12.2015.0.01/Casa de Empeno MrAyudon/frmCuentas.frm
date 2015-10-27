VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas de gastos"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   Icon            =   "frmCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   4515
   Begin VB.TextBox txtConcepto 
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
   Begin vbAcceleratorGrid6.vbalGrid grdConceptos 
      Height          =   5565
      Left            =   15
      TabIndex        =   1
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
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3450
      TabIndex        =   3
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
      Picture         =   "frmCuentas.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdEliminar 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
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
      Picture         =   "frmCuentas.frx":055E
      PictureDisabled =   "frmCuentas.frx":0AB0
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   3420
      TabIndex        =   5
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
      Picture         =   "frmCuentas.frx":1682
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
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()

    If Trim(txtConcepto.text) <> "" Then
    
        Grabar_Cuentas
    Else
        
        txtConcepto.SetFocus
    End If

End Sub

Private Sub cmdEliminar_Click()

    If grdConceptos.SelectedRow > 0 Then
        
        If MsgBox("Desea eliminar la cuenta seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton2, "Cuentas") = vbYes Then
            
            dbDatos.Execute "DELETE FROM cuentasgastos WHERE ID=" & grdConceptos.CellItemData(grdConceptos.SelectedRow, 1)
            grdConceptos.RemoveRow grdConceptos.SelectedRow
            grdConceptos.ClearSelection
            
        End If
    
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtConcepto_GotFocus()
    Seleccionar_Texto txtConcepto
    Cambiar_Color True, txtConcepto
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    KeyAscii = mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtConcepto_LostFocus()
    Cambiar_Color False, txtConcepto
End Sub

Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Crear_Encabezados
    Cargar_Cuentas
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

'creamos los encabezados del grid
Private Sub Crear_Encabezados()

    With grdConceptos
    
        .AddColumn "K1", "Descripción", ecgHdrTextALignLeft, , 272, , , , , , , CCLSortString
    
    End With

End Sub

'grabamos los conceptos
Private Sub Grabar_Cuentas()

On Error GoTo error
   
    grdConceptos.AddRow
    grdConceptos.CellDetails grdConceptos.Rows, 1, txtConcepto.text, DT_LEFT Or DT_WORD_ELLIPSIS
   
    dbDatos.Execute "INSERT INTO cuentasgastos (Fecha,Cuenta,Descripcion) VALUES ('" & _
                    Format(Date, "YYYY/MM/DD") & "','511101','" & Trim(txtConcepto.text) & "')"
   
    grdConceptos.Clear
    Cargar_Cuentas
    txtConcepto.text = ""
    txtConcepto.SetFocus
    
error:
    Maneja_Error Err

End Sub

'Cargamos las cuentas
Private Sub Cargar_Cuentas()
Dim rcConceptos As New ADODB.Recordset
   
On Error GoTo error
    
    rcConceptos.Open "SELECT * FROM cuentasgastos order by descripcion", dbDatos, adOpenForwardOnly, adLockOptimistic
    grdConceptos.Clear
    grdConceptos.Redraw = False

    With rcConceptos
    
        While Not .EOF
            
            grdConceptos.AddRow
            grdConceptos.CellDetails grdConceptos.Rows, 1, !Descripcion, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , , , rcConceptos!ID
        .MoveNext
        Wend
    
    End With

    grdConceptos.Redraw = True
    rcConceptos.Close
    Set rcConceptos = Nothing
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcConceptos = Nothing
End Sub
