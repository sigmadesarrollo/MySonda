VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmImpresoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Impresoras"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImpresoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   5640
   Begin VB.Frame Frame9 
      Caption         =   "Impresoras"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1845
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5415
      Begin VB.ComboBox cmbImpAlmoneda 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2775
      End
      Begin VB.ComboBox cmbFormatos 
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox cmbImpReportes 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ComboBox cmbImpContratos 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cmbImpTickets 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox cmbPCEti 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2520
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ComboBox cmbImpEti 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2775
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosImpresoras 
         Height          =   225
         Left            =   6720
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Etiquetas Almoneda:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   18
         Top             =   1110
         Width           =   2280
      End
      Begin VB.Label Label95 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formato:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6720
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label94 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora Reportes:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label Label93 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contratos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   390
         Width           =   1170
      End
      Begin VB.Label Label80 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora Tickets:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   1470
         Width           =   2115
      End
      Begin VB.Label Label81 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Etiquetas Empeño:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   750
         Width           =   2055
      End
      Begin VB.Label Label82 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PC Etiquetas:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   11
         Top             =   2520
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4425
      TabIndex        =   5
      Top             =   2940
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
      Picture         =   "frmImpresoras.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3285
      TabIndex        =   4
      Top             =   2940
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Aceptar"
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
      Picture         =   "frmImpresoras.frx":055E
   End
   Begin VB.Label Label83 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar la impresora correspondiente al tipo de impresión que desea obtener (Consulte a un técnico especializado)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   810
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   5130
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmImpresoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fl() As cFlatControl

Private Sub cmbImpAlmoneda_GotFocus()
    Cambiar_Color True, cmbImpAlmoneda
End Sub

Private Sub cmbImpAlmoneda_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbImpAlmoneda_LostFocus()
    Cambiar_Color False, cmbImpAlmoneda
End Sub

Private Sub cmbImpContratos_GotFocus()
    Cambiar_Color True, cmbImpContratos
End Sub

Private Sub cmbImpContratos_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbImpContratos_LostFocus()
    Cambiar_Color False, cmbImpContratos
End Sub

Private Sub cmbImpEti_GotFocus()
    Cambiar_Color True, cmbImpEti
End Sub

Private Sub cmbImpEti_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbImpEti_LostFocus()
    Cambiar_Color False, cmbImpEti
End Sub

Private Sub cmbImpTickets_GotFocus()
    Cambiar_Color True, cmbImpTickets
End Sub

Private Sub cmbImpTickets_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbImpTickets_LostFocus()
    Cambiar_Color False, cmbImpTickets
End Sub

Private Sub cmdAceptar_Click()
    
    If cmbImpContratos.text <> "" And cmbImpEti.text <> "" And cmbImpAlmoneda.text <> "" Then
        
        Graba_Valor "Impresoras", "ImpresoraContratos", cmbImpContratos.text
        Graba_Valor "Impresoras", "ImpresoraEtiquetas", cmbImpEti.text
        Graba_Valor "Impresoras", "ImpresoraEtiquetasAlmoneda", cmbImpAlmoneda.text
        Graba_Valor "Impresoras", "ImpresoraTickets", cmbImpTickets.text
        
        MsgBox "Configuración guardada con éxito !!", vbInformation, "Configuración Impresoras"
    Else
        
        MsgBox "Seleccione la impresora para cada documento !!", vbInformation, "Configuración Impresoras"
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    
    Cargar_Impresoras
    
    cmbImpContratos.ListIndex = ComboInformacion(cmbImpContratos, 0, Regresa_Valor("Impresoras", "ImpresoraContratos", ""))
    cmbImpEti.ListIndex = ComboInformacion(cmbImpEti, 0, Regresa_Valor("Impresoras", "ImpresoraEtiquetas", ""))
    cmbImpAlmoneda.ListIndex = ComboInformacion(cmbImpAlmoneda, 0, Regresa_Valor("Impresoras", "ImpresoraEtiquetasAlmoneda", ""))
    cmbImpTickets.ListIndex = ComboInformacion(cmbImpAlmoneda, 0, Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
    
    CentrarForm Me, frmMDI
    Poner_Flat fl, Me.Controls, Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat fl
End Sub

Private Sub Cargar_Impresoras()
Dim prn As Printer

On Error GoTo error

    For Each prn In Printers
        
        cmbImpContratos.AddItem prn.DeviceName
        cmbImpEti.AddItem prn.DeviceName
        cmbImpAlmoneda.AddItem prn.DeviceName
        cmbImpTickets.AddItem prn.DeviceName
        
    Next
    
    Exit Sub
    
error:
    Maneja_Error Err
End Sub

