VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.2#0"; "vbalSGrid6.ocx"
Begin VB.Form frmRelacionarCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relacionar Cuentas de Contabilidad"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelacionarCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   12540
   Begin DevPowerFlatBttn.FlatBttn cmdSeleccionar 
      Height          =   615
      Left            =   10680
      TabIndex        =   8
      Top             =   240
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1085
      AlignCaption    =   4
      AutoSize        =   0   'False
      Caption         =   "Seleccionar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
   Begin vbAcceleratorSGrid6.vbalGrid grdContpaq 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4471
      RowMode         =   -1  'True
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
      BorderStyle     =   2
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
      DefaultRowHeight=   18
   End
   Begin VB.TextBox txtArchivo 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin DevPowerFlatBttn.FlatBttn cmdArchivo 
      Height          =   225
      Left            =   5760
      TabIndex        =   2
      Top             =   360
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   397
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
   Begin vbAcceleratorSGrid6.vbalGrid grdSistema 
      Height          =   2535
      Left            =   6360
      TabIndex        =   5
      Top             =   960
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4471
      RowMode         =   -1  'True
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
      BorderStyle     =   2
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
      DefaultRowHeight=   18
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdCuentas 
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   5953
      RowMode         =   -1  'True
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
      BorderStyle     =   2
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
      DefaultRowHeight=   18
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   11280
      TabIndex        =   9
      Top             =   7080
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
      Picture         =   "frmRelacionarCuentas.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   9960
      TabIndex        =   10
      Top             =   7080
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
      Picture         =   "frmRelacionarCuentas.frx":055E
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cuentas del Sistema"
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
      Left            =   6360
      TabIndex        =   6
      Top             =   720
      Width           =   1965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cuentas de Contabilidad"
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
      TabIndex        =   4
      Top             =   720
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione el archivo de cuentas a importar:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3165
   End
End
Attribute VB_Name = "frmRelacionarCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fl() As New cFlatControl

Private Sub cmdAceptar_Click()
   GrabarCuentas
End Sub

Private Sub cmdArchivo_Click()
   txtArchivo.text = OpenFileDialog(Me)
   If txtArchivo.text <> "" Then
      LeerArchivoExcel (txtArchivo.text)
   End If
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
   If grdContpaq.SelectedRow > 0 And grdSistema.SelectedRow > 0 Then
      With grdCuentas
         .AddRow
         .CellDetails .Rows, 1, grdSistema.CellText(grdSistema.SelectedRow, 1), DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4, , grdSistema.CellItemData(grdSistema.SelectedRow, 1)
         .CellDetails .Rows, 2, grdSistema.CellText(grdSistema.SelectedRow, 2), DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
         .CellDetails .Rows, 3, grdContpaq.CellText(grdContpaq.SelectedRow, 1), DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
         .CellDetails .Rows, 4, grdContpaq.CellText(grdContpaq.SelectedRow, 2), DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
         
         'grdContpaq.RemoveRow grdContpaq.SelectedRow
         grdSistema.RemoveRow grdSistema.SelectedRow
                  
      End With
   End If
End Sub

Private Sub Form_Load()
   Inicializar
End Sub

Private Sub Inicializar()
   Poner_Flat Fl, Me.Controls, Me
   CentrarForm Me, frmMDI
   CrearEncabezados
    CargarCuentasSistema
End Sub

Private Sub CrearEncabezados()
   With grdContpaq
      .AddColumn "K1", "Cuenta", ecgHdrTextALignLeft, , 116, , , , , , , CCLSortString
      .AddColumn "K2", "Descripcion", ecgHdrTextALignLeft, , 261, , , , , , , CCLSortString
   End With
   
   With grdSistema
      .AddColumn "K1", "Cuenta", ecgHdrTextALignLeft, , 116, , , , , , , CCLSortString
      .AddColumn "K2", "Descripcion", ecgHdrTextALignLeft, , 261, , , , , , , CCLSortString
   End With
   
   With grdCuentas
      .AddColumn "K1", "Cuenta Sistema", ecgHdrTextALignLeft, , 116, , , , , , , CCLSortString
      .AddColumn "K2", "Descripcion", ecgHdrTextALignLeft, , 261, , , , , , , CCLSortString
      .AddColumn "K3", "Cuenta Contpaq", ecgHdrTextALignLeft, , 116, , , , , , , CCLSortString
      .AddColumn "K4", "Descripcion", ecgHdrTextALignLeft, , 261, , , , , , , CCLSortString
   End With
   
End Sub

Private Sub grdContpaq_ColumnWidthChanged(ByVal lCol As Long, lWidth As Long, bCancel As Boolean)
   'grdContpaq.CellText(1, lCol) = lWidth
End Sub

Private Sub grdCuentas_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
   If KeyCode = vbKeyDelete Then
      
      grdSistema.AddRow
      grdSistema.CellDetails grdSistema.Rows, 1, grdCuentas.CellText(grdCuentas.SelectedRow, 1), DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4, , grdCuentas.CellItemData(grdCuentas.SelectedRow, 1)
      grdSistema.CellDetails grdSistema.Rows, 2, grdCuentas.CellText(grdCuentas.SelectedRow, 2), DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
      grdCuentas.RemoveRow grdCuentas.SelectedRow
   
   End If
End Sub

Private Sub txtArchivo_GotFocus()
   Cambiar_Color True, txtArchivo
   Seleccionar_Texto txtArchivo
End Sub

Private Sub txtArchivo_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtArchivo_LostFocus()
   Cambiar_Color False, txtArchivo
End Sub

'cargamos las cuentas del archivo de excel
Private Sub LeerArchivoExcel(Archivo As String)
   Dim Excel As Object 'New Excel.Application
   Dim xLibro As Object 'Excel.Workbook
   Dim Salir As Boolean
   Dim Col As Long
   Dim Row As Long
   Dim Agregar As Boolean
   
   Set Excel = CreateObject("Excel.Application")
   Set xLibro = Excel.Workbooks.Open(Archivo)
   
   
   'Set xLibro = Excel.Workbooks.Open(Archivo)
   
   
   Excel.Visible = False
   
   Salir = False
   Agregar = False
   Col = 1
   Row = 1
   grdContpaq.Clear
   grdContpaq.Redraw = False
   With xLibro
   
      With .Sheets(1)
      
         While Not Salir
            DoEvents
            'checamos si podemos activar agregar
            If Not Agregar Then
               If UCase(Trim(.Cells(Row, 1))) = "C U E N T A" Then
                  Agregar = True
                  Row = Row + 2
               End If
            End If
            
            
            If Agregar Then
               'si el renglon siguiente es vacio, sacamos del ciclor
               If .Cells(Row, 1) <> "" Then
               
                  grdContpaq.AddRow
                  grdContpaq.CellDetails grdContpaq.Rows, 1, RTrim(LTrim(.Cells(Row, 1))), DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
                  grdContpaq.CellDetails grdContpaq.Rows, 2, RTrim(LTrim(.Cells(Row, 2))), DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
               Else
                  Salir = True
               End If
               
            End If
            
            Row = Row + 1
         Wend
      
      End With
   End With
   
   grdContpaq.Redraw = True
   Excel.Workbooks.Close
   Set Excel = Nothing
   Set xLibro = Nothing
   
End Sub

'cargamos la cuenta del sistema
Private Sub CargarCuentasSistema()
   Dim rc As New ADODB.Recordset
   
   rc.Open "SELECT * FROM Cuentas", dbDatos, adOpenDynamic, adLockOptimistic
   
   grdSistema.Clear
   grdCuentas.Clear
   grdSistema.Redraw = False
   grdCuentas.Redraw = False
   While Not rc.EOF
      DoEvents
      With grdSistema
         .AddRow
         .CellDetails .Rows, 1, rc!Cuenta, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4, , rc!ID
         .CellDetails .Rows, 2, rc!Descripcion, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
      End With
      
      With grdCuentas
         If (rc!Cuenta & "") <> "" Then
            .AddRow
            .CellDetails .Rows, 1, rc!Cuenta, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4, , rc!ID
            .CellDetails .Rows, 2, rc!Descripcion, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
            .CellDetails .Rows, 3, rc!Cuenta, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
            .CellDetails .Rows, 4, rc!Descripcion, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , 4
            grdSistema.RemoveRow .Rows
         End If
      End With
      
      
      rc.MoveNext
   Wend
   grdSistema.Redraw = True
   grdCuentas.Redraw = True
   
End Sub

Private Sub GrabarCuentas()
   Dim Indice As Long
   
   
   dbDatos.Execute "UPDATE Cuentas SET CuentaContpaq='',DescripcionContpaq=''"
   
   For Indice = 1 To grdCuentas.Rows
      With grdCuentas
         dbDatos.Execute "UPDATE Cuentas SET CuentaContpaq='" & .CellText(Indice, 3) & "',DescripcionContpaq='" & .CellText(Indice, 4) & "' WHERE ID=" & .CellItemData(Indice, 1)
      End With
   Next Indice
   
   MsgBox "Cuentas grabadas correctamente", vbInformation Or vbOKOnly
   
   Unload Me
End Sub
