VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Begin VB.Form frmMuestraMonedas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   5010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMuestraMonedas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin vbAcceleratorGrid6.vbalGrid grdMonedas 
      Height          =   2310
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   4075
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
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
End
Attribute VB_Name = "frmMuestramonedas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim obj As Object
Dim frm As Form

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Public Function Posicion(Form As Form, caja As Object)
    Set frm = Form
    Set obj = caja
    Position Me, obj
    Me.Show , frmMDI
End Function

Private Sub grdMonedas_DblClick(ByVal lRow As Long, ByVal lCol As Long)

    If grdMonedas.Rows > 0 And grdMonedas.SelectedRow > 0 Then
                    
        obj.Text = grdMonedas.CellText(grdMonedas.SelectedRow, 1)
        frm.TipoCambio grdMonedas.CellItemData(grdMonedas.SelectedRow, 1)  'SacaValor("monedas", "Clave", " WHERE Descripcion='" & Trim(grdMonedas.CellText(grdMonedas.SelectedRow, 1)) & "'")
        Unload Me
    End If

End Sub

Sub Inicializar()
    Crea_Encabezado
    Carga_Datos
End Sub

Sub Crea_Encabezado()

    With grdMonedas
        .ImageList = frmMDI.img
        .AddColumn "K1", "Divisa", ecgHdrTextALignLeft, , 310
    End With

End Sub

Sub Carga_Datos()
Dim rcConsulta As New ADODB.Recordset

    rcConsulta.Open "SELECT * FROM monedas WHERE Descripcion<>'Moneda Nacional' ORDER BY Descripcion", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then

        With grdMonedas
            While Not rcConsulta.EOF
                .AddRow
                .CellDetails .Rows, 1, rcConsulta!Descripcion, DT_LEFT, frmMDI.img.ItemIndex(3), , , , , , rcConsulta!clave
            rcConsulta.MoveNext
            Wend
        End With

    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
End Sub
