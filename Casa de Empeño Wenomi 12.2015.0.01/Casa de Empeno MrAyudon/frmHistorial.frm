VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Begin VB.Form frmHistorial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   9360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmHistorial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin vbAcceleratorGrid6.vbalGrid grdHistorial 
      Height          =   3000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   5292
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
   Begin VB.Label lblAlmoneda 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8760
      TabIndex        =   11
      Top             =   3060
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Almoneda:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7680
      TabIndex        =   10
      Top             =   3060
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Vencidos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6000
      TabIndex        =   9
      Top             =   3060
      Width           =   945
   End
   Begin VB.Label lblVencidos 
      BackColor       =   &H8000000A&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   8
      Top             =   3060
      Width           =   435
   End
   Begin VB.Label lblRefrendos 
      BackColor       =   &H8000000A&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   3060
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Refrendos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4200
      TabIndex        =   6
      Top             =   3060
      Width           =   1080
   End
   Begin VB.Label lblDesempeños 
      BackColor       =   &H8000000A&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3540
      TabIndex        =   5
      Top             =   3060
      Width           =   435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Desempeños:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2160
      TabIndex        =   4
      Top             =   3060
      Width           =   1305
   End
   Begin VB.Label lblEmpeños 
      BackColor       =   &H8000000A&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1530
      TabIndex        =   3
      Top             =   3060
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Emp. activos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   3060
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000000&
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   9345
   End
End
Attribute VB_Name = "frmHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim consulta As ADODB.Recordset
Dim obj As Object
Dim frm As Form
Dim Carga As Boolean
Dim Ban As Boolean

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
    Crea_Encabezados
End Sub

Sub Crea_Encabezados()
    With grdHistorial
        .AddColumn "K0", "Contrato", ecgHdrTextALignCentre, , 60, , , , , , , CCLSortNumeric
        .AddColumn "K1", "Origen", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
        .AddColumn "K2", "Destino", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
        .AddColumn "K3", "Vencimiento", ecgHdrTextALignLeft, , 85, , , , , "dd/mmm/yyyy", , CCLSortString
        .AddColumn "K4", "Movimiento", ecgHdrTextALignLeft, , 85, , , , , "dd/mmm/yyyy", , CCLSortString
        .AddColumn "K5", "Status", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "K6", "Dias", ecgHdrTextALignRight, , 50, , , , , , , CCLSortNumeric
    End With
End Sub

Public Sub muestra_historial(ID As Long)
Dim rcConsulta As New ADODB.Recordset

On Error GoTo Error

    rcConsulta.Open "SELECT NumContrato,Origen,FolioOrigen,Cancelado,Destino,FolioDestino,Pagado,Vencimiento,FechaMovimiento,Almoneda FROM empeno WHERE IDCliente=" & ID & " ORDER BY NumContrato,Folio", dbDatos, adOpenForwardOnly, adLockReadOnly
    With grdHistorial
        
        While Not rcConsulta.EOF
            .AddRow
            .CellText(.Rows, 1) = rcConsulta!NumContrato
            .CellTextAlign(.Rows, 1) = DT_RIGHT
            .CellText(.Rows, 2) = OD_Origen(rcConsulta!Origen) & IIf(rcConsulta!Cancelado = 1 And rcConsulta!Destino = 0, "/Cancelado", "/" & rcConsulta!FolioOrigen)
            If rcConsulta!Cancelado = 1 Then
                .CellText(.Rows, 3) = "Cancelado"
                .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 6) = "Cancelado"
             Else
                .CellText(.Rows, 3) = OD_Origen(rcConsulta!Destino) & IIf(rcConsulta!FolioDestino = 0, "", "/" & rcConsulta!FolioDestino)
                .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 6) = IIf(rcConsulta!Almoneda = 1, "Almoneda", IIf(rcConsulta!Pagado = 1, "Liquidado", IIf(rcConsulta!Vencimiento < Date, "Vencido", "Activo")))
             End If
            .CellText(.Rows, 4) = rcConsulta!Vencimiento
            .CellText(.Rows, 5) = rcConsulta!FechaMovimiento
            
        rcConsulta.MoveNext
        Wend
    End With
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Public Sub Ver(frmOBJ As Form, Ob As Object, Optional BD As Boolean = False, Optional x As Integer, Optional ID As Long)
   Set obj = Ob
   Set frm = frmOBJ
   Position Me, Ob
   Carga = BD
   If x = 0 Then Ban = False
   Inicializar ID
   Me.Show , frmMDI
End Sub

Sub Inicializar(ID As Long)
    'crea_encabezados
    muestra_historial ID
    Carga_Informacion ID
End Sub

Sub Carga_Informacion(ID As Long)
Dim rcTmp As ADODB.Recordset
Dim ctrl As Control

On Error GoTo Error
  
  For Each ctrl In Controls
      If TypeOf ctrl Is Label Then ctrl.BackColor = RGB(223, 208, 102)
  Next

'Empeños
Set rcTmp = dbDatos.Execute("select Count(ID) as registros from empeno where IDCliente=" & ID & " and destino=0 and cancelado=0 and pagado=0 and vencimiento>='" & Format(Date, "YYYY/MM/DD") & "'")
If Not rcTmp.BOF And Not rcTmp.EOF Then
    lblEmpeños.Caption = rcTmp!registros
End If

'Desempeños
Set rcTmp = dbDatos.Execute("select count(id) as registros from empeno where IDCliente=" & ID & " and destino=3")
If Not rcTmp.BOF And Not rcTmp.EOF Then lblDesempeños.Caption = rcTmp!registros

'Refrendos
Set rcTmp = dbDatos.Execute("select count(id) as registros from empeno where IDCliente=" & ID & " and destino=2")
If Not rcTmp.BOF And Not rcTmp.EOF Then lblRefrendos.Caption = rcTmp!registros

'Refrendos
Set rcTmp = dbDatos.Execute("select count(id) as registros from empeno where IDCliente=" & ID & " and destino=4")
If Not rcTmp.BOF And Not rcTmp.EOF Then lblAlmoneda.Caption = rcTmp!registros

'Vencidos
Set rcTmp = dbDatos.Execute("select count(id) as registros from empeno where IDCliente=" & ID & " and destino=0 and vencimiento<'" & Format(Date, "YYYY/MM/DD") & "'")
If Not rcTmp.BOF And Not rcTmp.EOF Then
    lblVencidos.Caption = rcTmp!registros
End If

Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Sub
