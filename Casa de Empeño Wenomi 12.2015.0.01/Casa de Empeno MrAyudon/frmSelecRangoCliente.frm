VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL2.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmSelecRangoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Identificaciones de Clientes"
   ClientHeight    =   3555
   ClientLeft      =   5835
   ClientTop       =   4725
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelecRangoCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   5400
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame frameApellido 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   5175
      Begin VB.TextBox txtApellido 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2760
         MaxLength       =   40
         TabIndex        =   11
         Top             =   120
         Width           =   2250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido Paterno empieze con:"
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
         TabIndex        =   7
         Top             =   120
         Width           =   2550
      End
   End
   Begin VB.Frame frameCliente 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   5175
      Begin VB.TextBox txtNombre 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   8
         Top             =   15
         Width           =   3450
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosCliente3 
         Height          =   255
         Left            =   4545
         TabIndex        =   9
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
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
         TabIndex        =   10
         Top             =   15
         Width           =   765
      End
   End
   Begin VB.OptionButton optTodos 
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.OptionButton optApellidoP 
      Caption         =   "Por Apellido Paterno"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.OptionButton optCliente 
      Caption         =   "Por Cliente"
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
      TabIndex        =   2
      Top             =   240
      Value           =   -1  'True
      Width           =   1815
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2955
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   2
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "       &Imprimir"
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
      Object.ToolTipText     =   ""
      Picture         =   "frmSelecRangoCliente.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2955
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
      Picture         =   "frmSelecRangoCliente.frx":055E
   End
End
Attribute VB_Name = "frmSelecRangoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function Validar() As Boolean
    Validar = True
    
    If optCliente.Value = True Then
        If txtNombre.Tag = "" Or txtNombre.Tag = 0 Then
            MsgBox "Seleccione el Cliente."
            Validar = False
        End If
    ElseIf optApellidoP.Value = True Then
        If Trim(txtApellido.text) = "" Then
            MsgBox "Especifique el Criterio de busqueda para el Apellido."
            Validar = False
        End If
    End If
    
    
End Function

Private Sub cmdAceptar_Click()

    If Validar Then GeneraExcel
End Sub

Private Sub cmdMosCliente3_Click()

    frmMostrarCliente.Ver Me, txtNombre, True, 0
    'frmMostrarCliente.Show
    
    'frmMostrarclienteventasSelec.Ver Me, txtNombre, True, 0
    'frmMostrarclienteventasSelec.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbHourglass
    txtNombre.text = ""
    txtNombre.Tag = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub optApellidoP_Click()
    frameCliente.Enabled = False
    txtNombre.text = ""
    txtNombre.Tag = ""
    
    frameApellido.Enabled = True
    txtApellido.text = ""
End Sub

Private Sub optCliente_Click()
    frameCliente.Enabled = True
    txtNombre.text = ""
    txtNombre.Tag = ""
    
    frameApellido.Enabled = False
    txtApellido.text = ""
    
End Sub

Private Sub optTodos_Click()
    frameCliente.Enabled = False
    txtNombre.text = ""
    txtNombre.Tag = ""
    
    frameApellido.Enabled = False
    txtApellido.text = ""
End Sub

Private Sub txtApellido_GotFocus()
    Seleccionar_Texto txtApellido
    Cambiar_Color True, txtApellido
End Sub

Private Sub txtApellido_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtApellido_LostFocus()
    Cambiar_Color False, txtApellido
End Sub
 
Private Sub GeneraExcel()
    
    Dim Excel As Object, i As Integer, Col As Integer, Y As Integer, str As String, Columnas As Integer
    Dim Rs As New ADODB.Recordset, Sql As String, Filtro As String
    
    Const xlCenter As Long = -4108
    Const xlBottom As Long = -4107
    Const xlLeft As Long = -4131
    
    On Error GoTo Error

    'Screen.MousePointer = vbHourglass
    
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0
    
    DoEvents
    
    If optCliente.Value = True Then
        Filtro = " WHERE ID=" & txtNombre.Tag
    ElseIf optApellidoP.Value = True Then
        Filtro = " WHERE Apellido LIKE '%" & Trim(txtApellido.text) & "%'"
    Else
        Filtro = ""
    End If
    
    Rs.Open "SELECT * FROM clientes " & Filtro, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not Rs.EOF Then
        
        ProgressBar1.Max = Val(SacaValor("clientes", "COUNT(ID)", Filtro))
        
        'Creo la Referencia al Excel
        Set Excel = CreateObject("Excel.application")
        
        With Excel
        
            'Agrego un Nuevo Libro
            .Workbooks.Add
            
            '.................................................................
            
            .Range("A:A,B:B").Select
            .Range("B1").Activate
            With .Selection.Font
                .Name = "Calibri"
                .Size = 8
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                
            End With
            .Range("C12").Select
            .Columns("B:B").EntireColumn.AutoFit
                    
            '-----------------------------------------------------------
            Y = 5
            i = 0
            Do While Not Rs.EOF
            
                .Range("A" & CStr(Y)).formula = "NOMBRE:"
                .Range("A" & CStr(Y + 1)).formula = "DIRECCION:"
                .Range("A" & CStr(Y + 2)).formula = "COLONIA:"
                .Range("A" & CStr(Y + 3)).formula = "MUNICIPIO:"
                .Range("A" & CStr(Y + 4)).formula = "ESTADO:"
                .Range("A" & CStr(Y + 5)).formula = "IDENTIFICACION:"
                .Range("A" & CStr(Y + 6)).formula = "NUMERO:"
                .Range("A" & CStr(Y + 7)).formula = "TELEFONO:"
                
                .Range("B" & CStr(Y)).formula = Rs!Nombre & " " & Rs!Apellido
                .Range("B" & CStr(Y + 1)).formula = Rs!Direccion
                .Range("B" & CStr(Y + 2)).formula = Rs!Colonia
                .Range("B" & CStr(Y + 3)).formula = Rs!Municipio
                .Range("B" & CStr(Y + 4)).formula = Rs!Estado
                .Range("B" & CStr(Y + 5)).formula = UCase(Rs!Identificacion)
                .Range("B" & CStr(Y + 6)).formula = Rs!NumeroIdentificacion
                .Range("B" & CStr(Y + 7)).formula = Rs!Tel
                
                If Dir(Path & "\Fotos\" & Rs!Nombre & " " & Rs!Apellido & "-CRED1.jpg") <> "" Then
                   .Range("D" & CStr(Y)).Select
                   .ActiveSheet.Pictures.Insert(Path & "\Fotos\" & Rs!Nombre & " " & Rs!Apellido & "-CRED1.jpg").Select
                   .Selection.ShapeRange.Height = 127.5590551181
                End If
                
                If Dir(Path & "\Fotos\" & Rs!Nombre & " " & Rs!Apellido & "-CRED2.jpg") <> "" Then
                   .Range("G" & CStr(Y)).Select
                   .ActiveSheet.Pictures.Insert(Path & "\Fotos\" & Rs!Nombre & " " & Rs!Apellido & "-CRED2.jpg").Select
                   .Selection.ShapeRange.Height = 127.5590551181
                End If
                    
                Y = Y + 10
                i = i + 1
                
                ProgressBar1.Value = i
                
                Rs.MoveNext
            Loop
            '-----------------------------------------------------------
                
            .Range("A1").formula = "REPORTE DE IDENTIFICACIONES DE CLIENTES"
            .Range("A1:G1").Select
            .Selection.Merge
            With .Selection.Font
                .Name = "Calibri"
                .Size = 14
            End With
            
            .Range("A1").Select
            .Selection.Font.Size = 14
            With .Selection
                .HorizontalAlignment = xlCenter
            End With
                        
            '.................................................................
                
            .Range("A2").formula = "SUCURSAL: " & Sucursal.NombreComercial
            .Range("A2:G2").Select
            .Selection.Merge
            With .Selection.Font
                .Name = "Calibri"
                .Size = 12
            End With
            
            .Range("A2").Select
            .Selection.Font.Size = 12
            With .Selection
                .HorizontalAlignment = xlCenter
            End With
                
            '.................................................................
                
            .Range("A1").Select
            'Hago Visible la Referencia
            .Visible = True

        End With
        
        Set Excel = Nothing
    
    Else
    
        MsgBox "No hay informacion que reportar.", vbInformation, "Identificaciones de Clientes"
        
    End If
    Rs.Close
    Set Rs = Nothing
    
    ProgressBar1.Value = 0
    ProgressBar1.Max = 100
    
Exit Sub
    
Error:
    'Resume
    Set Excel = Nothing
    Screen.MousePointer = vbDefault
    Maneja_Error Err
End Sub

'MLD-MODIF.- Buscamos el id cliente
Public Sub Buscar(ID As Long)
    Dim Rs As New ADODB.Recordset
    
    On Error GoTo Error
    
    Rs.Open "SELECT CONCAT(Nombre,' ',Apellido) AS Nombre FROM Clientes WHERE Id=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not Rs.EOF Then
        
        txtNombre.text = Rs!Nombre
        txtNombre.Tag = ID
        
    End If
    Rs.Close
    Set Rs = Nothing
    
Exit Sub
    
Error:
    Maneja_Error Err
End Sub

