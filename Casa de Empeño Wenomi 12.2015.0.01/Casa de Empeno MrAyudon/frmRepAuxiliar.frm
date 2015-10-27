VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmRepAuxiliar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Auxiliar"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepAuxiliar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   5325
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Por el Criterio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   5205
      Begin VB.TextBox txtA 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1470
      End
      Begin VB.TextBox txtDe 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   1470
      End
      Begin VB.OptionButton optFechas 
         Appearance      =   0  'Flat
         Caption         =   "Fechas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1170
      End
      Begin VB.OptionButton optHoy 
         Appearance      =   0  'Flat
         Caption         =   "Del dia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1245
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
         Height          =   300
         Index           =   0
         Left            =   2940
         TabIndex        =   5
         Top             =   720
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         AlignCaption    =   4
         AlignPicture    =   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         MousePointer    =   1
         PlaySounds      =   0   'False
         TextColor       =   -2147483630
         Object.ToolTipText     =   ""
         Picture         =   "frmRepAuxiliar.frx":000C
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
         Height          =   300
         Index           =   1
         Left            =   4845
         TabIndex        =   6
         Top             =   720
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         AlignCaption    =   4
         AlignPicture    =   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         MousePointer    =   1
         PlaySounds      =   0   'False
         TextColor       =   -2147483630
         Object.ToolTipText     =   ""
         Picture         =   "frmRepAuxiliar.frx":0121
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha final:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha inicial:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.ComboBox cmbCuentas 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   3945
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   4095
      TabIndex        =   11
      Top             =   2100
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
      Picture         =   "frmRepAuxiliar.frx":0236
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   2100
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   2
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "        &Imprimir"
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
      Picture         =   "frmRepAuxiliar.frx":0788
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   930
   End
End
Attribute VB_Name = "frmRepAuxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbCuentas_GotFocus()
    cmbCuentas.BackColor = &HC0FFFF
End Sub

Private Sub cmbCuentas_LostFocus()
    cmbCuentas.BackColor = vbWhite
End Sub

Private Sub cmdAceptar_Click()

    If optHoy.Value Then
        
        Reporte_Auxiliar False, Date, Date, IIf(cmbCuentas.ItemData(cmbCuentas.ListIndex) = 0, "Todas", cmbCuentas.ItemData(cmbCuentas.ListIndex))
        Sleep 1000
        Imprimir_Auxiliar False, Date, Date, IIf(cmbCuentas.text = "Todas", "", cmbCuentas.ItemData(cmbCuentas.ListIndex))
    Else
        
        Reporte_Auxiliar True, txtDe, txtA, IIf(cmbCuentas.ItemData(cmbCuentas.ListIndex) = 0, "Todas", cmbCuentas.ItemData(cmbCuentas.ListIndex))
        Sleep 1000
        Imprimir_Auxiliar True, txtDe, txtA, IIf(cmbCuentas.text = "Todas", "", cmbCuentas.ItemData(cmbCuentas.ListIndex))
    End If

End Sub

Private Sub cmdMosFecha_Click(Index As Integer)
    Select Case Index
        Case 0
            txtDe.text = frmCalendario.Fecha(txtDe.text, 1)
        
        Case 1
            txtA.text = frmCalendario.Fecha(txtA.text, 1)
    
    End Select
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmbCuentas.AddItem "Todas"
    Cargar_Combos "concepto", "mayor", "cuentas", cmbCuentas
    cmbCuentas.ListIndex = 0
    Poner_Flat Fl, Me.Controls, Me
    Oculta
    CentrarForm Me, frmMDI
End Sub

Private Sub Cargar_Combos(Campo As String, Campo2 As String, Tabla As String, Combo As ComboBox)
Dim rcTipos As New ADODB.Recordset

On Error GoTo Error

    rcTipos.Open "SELECT DISTINCT " & Campo & "," & Campo2 & " FROM " & Tabla & " order by " & Campo & "", dbDatos, adOpenForwardOnly, adLockReadOnly
    With rcTipos
        While Not .EOF
            Combo.AddItem .Fields(Campo)
            Combo.ItemData(Combo.NewIndex) = !Mayor
        .MoveNext
        Wend
    End With
    rcTipos.Close
    Set rcTipos = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcTipos = Nothing
End Sub

Sub Oculta()
    Label2.Visible = False
    txtDe.Visible = False
    cmdMosFecha(0).Visible = False
    Label3.Visible = False
    txtA.Visible = False
    cmdMosFecha(1).Visible = False
    txtDe.text = Format(Date, "DD/MMM/YYYY")
    txtA.text = Format(Date, "DD/MMM/YYYY")
End Sub

Sub Muestra()
    Label2.Visible = True
    txtDe.Visible = True
    cmdMosFecha(0).Visible = True
    Label3.Visible = True
    txtA.Visible = True
    cmdMosFecha(1).Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub optFechas_Click()
    Muestra
End Sub

Private Sub optHoy_Click()
    Oculta
End Sub

'Sacamos el reporte auxiliar
Private Sub Reporte_Auxiliar(Optional Opcion As Boolean = False, Optional FechaIni As String, Optional FechaFin As String, Optional Cuenta As String = "")
Dim rcAuxiliar As New ADODB.Recordset
Dim rcSaldo As New ADODB.Recordset
Dim rcAux As New ADODB.Recordset
Dim crCargo As Currency, crAbono As Currency, crSaldo As Currency, strCuenta As String, strConcepto As String

On Error GoTo Error

    dbReportes.Execute "DELETE FROM cortecuentas"
   
    If Cuenta = "Todas" Then
        
        rcAux.Open "SELECT DISTINCT cuentas.Mayor,cuentas.Concepto FROM auxiliar INNER JOIN cuentas ON auxiliar.Cuenta=cuentas.Cuenta GROUP BY auxiliar.Cuenta ORDER BY auxiliar.Cuenta", dbDatos, adOpenForwardOnly, adLockReadOnly
    Else
        
        rcAux.Open "SELECT DISTINCT cuentas.Mayor,cuentas.Concepto FROM auxiliar INNER JOIN cuentas ON auxiliar.Cuenta=cuentas.Cuenta WHERE cuentas.Mayor='" & Trim(Cuenta) & "' GROUP BY auxiliar.Cuenta ORDER BY auxiliar.Cuenta", dbDatos, adOpenForwardOnly, adLockReadOnly
    End If
    
    With rcAux
    
        While Not .EOF
            
            crCargo = 0
            crAbono = 0
            strCuenta = !Mayor
            strConcepto = !Concepto
            
            'Saco el saldo Inicial
            rcSaldo.Open "SELECT SUM(auxiliar.Importe) AS Importe,cuentas.Mayor,auxiliar.Cuenta,cuentas.Concepto,Mid(auxiliar.Cuenta,5,2) AS Tipo FROM auxiliar INNER JOIN cuentas ON auxiliar.Cuenta=cuentas.Cuenta WHERE cuentas.Mayor='" & Trim(!Mayor) & "' AND Fecha<'" & Format(CDate(FechaIni), "YYYY/MM/DD") & "' GROUP BY auxiliar.Cuenta,Mid(auxiliar.Cuenta,5,2) ORDER BY auxiliar.Cuenta", dbDatos, adOpenForwardOnly, adLockReadOnly
            If Not rcSaldo.BOF And Not rcSaldo.EOF Then
                crCargo = 0
                crAbono = 0
                strCuenta = rcSaldo!Mayor
                strConcepto = rcSaldo!Concepto
                While Not rcSaldo.EOF
                    If strCuenta = rcSaldo!Mayor And Right(rcSaldo!Cuenta, 2) = "01" Then
                        
                        crCargo = crCargo + rcSaldo!Importe
                    ElseIf strCuenta = rcSaldo!Mayor And Right(rcSaldo!Cuenta, 2) = "50" Then
                        
                        crAbono = crAbono + rcSaldo!Importe
                    End If
                rcSaldo.MoveNext
                Wend
            End If
            rcSaldo.Close
            
            'Saco el Saldo
            crSaldo = crCargo - crAbono
            
            'Grabo el Saldo
            dbReportes.Execute "INSERT INTO cortecuentas (Cuenta,Descripcion,Fecha,Concepto,Folio,Cargo,Abono,Saldo,PC) VALUES " & _
                                "('" & Trim(strCuenta) & "','" & Trim(strConcepto) & "','" & Format(CDate(FechaIni), "YYYY/MM/DD") & "','SALDO INICIAL',0," & crCargo & "," & crAbono & "," & crSaldo & ",'" & Trim(NombrePc) & "')"
            
            
            'Saco el Detalle de la cuenta----------------------------------------------------------
            rcAuxiliar.Open "SELECT auxiliar.*,cuentas.Mayor,cuentas.Concepto AS CuentaConcepto FROM auxiliar,cuentas WHERE cuentas.Cuenta=auxiliar.Cuenta AND Fecha BETWEEN '" & Format(CDate(FechaIni), "YYYY/MM/DD") & "' AND '" & Format(CDate(FechaFin), "YYYY/MM/DD") & "' AND cuentas.Mayor='" & Trim(strCuenta) & "' ORDER BY cuentas.Mayor,auxiliar.ID", dbDatos, adOpenForwardOnly, adLockReadOnly
            While Not rcAuxiliar.EOF
            
                crCargo = 0
                crAbono = 0
                
                If strCuenta = "710300" And rcAuxiliar!Serie <> 1 Then GoTo NoAuxiliar
                
                If strCuenta <> rcAuxiliar!Mayor Then
                   
                    crSaldo = 0
                    strCuenta = rcAuxiliar!Mayor
                End If
               
                If Right(rcAuxiliar!Cuenta, 2) = "01" Then
                   
                    crCargo = rcAuxiliar!Importe
                    crSaldo = crSaldo + crCargo
                Else
                   
                    crAbono = rcAuxiliar!Importe
                    crSaldo = crSaldo - crAbono
                End If
        
                dbReportes.Execute "INSERT INTO cortecuentas (Cuenta,Descripcion,Fecha,Concepto,Folio,Cargo,Abono,Saldo,PC) VALUES " & _
                                                           "('" & rcAuxiliar!Mayor & "','" & rcAuxiliar!cuentaconcepto & "','" & Format(rcAuxiliar!Fecha, "YYYY/MM/DD") & "','" & rcAuxiliar!Concepto & "'," & rcAuxiliar!Folio & "," & crCargo & "," & crAbono & "," & crSaldo & ",'" & rcAuxiliar!PC & "')"
NoAuxiliar:
            rcAuxiliar.MoveNext
            Wend
            rcAuxiliar.Close
            '------------------------------------------------------------------------------------------
        
        .MoveNext
        Wend
    
    End With
    
    rcAux.Close
    Set rcAuxiliar = Nothing
    Set rcAux = Nothing
    Set rcSaldo = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcAuxiliar = Nothing
    Set rcAux = Nothing
    Set rcSaldo = Nothing
End Sub

Private Sub Imprimir_Auxiliar(Optional Opcion As Boolean, _
                              Optional FechaIni As String, _
                              Optional FechaFin As String, _
                              Optional Cuenta As String = "")
    On Error GoTo Error

    With frmMDI.Cr
        .Reset
        .ReportFileName = Path & "\Reportes\Auxiliar.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = IIf(Cuenta = "", "", "{CorteCuentas.Cuenta}='" & Cuenta & "'")

        If FechaIni <> "" And FechaFin <> "" And Cuenta <> "" Then
            .Formulas(0) = "Encabezado='De la Fecha " & Format(FechaIni, "DD/MMM/YYYY") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        ElseIf FechaIni <> "" And FechaFin <> "" Then
            .Formulas(0) = "Encabezado='De la Fecha " & Format(FechaIni, "DD/MMM/YYYY") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        Else
            .Formulas(0) = "Encabezado='De la fecha " & Format(Date, "DD/MMM/YYYY") & "'"
        End If

        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(3) = "Usuario='" & frmMDI.Usuario & "'"
        .WindowShowPrintSetupBtn = True
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .WindowTitle = "Reporte Auxiliar"
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub txtA_GotFocus()
    Seleccionar_Texto txtA
    Cambiar_Color True, txtA
End Sub

Private Sub txtA_LostFocus()
    Cambiar_Color False, txtA
End Sub

Private Sub txtDe_GotFocus()
    Seleccionar_Texto txtDe
    Cambiar_Color True, txtDe
End Sub

Private Sub txtDe_LostFocus()
    Cambiar_Color False, txtDe
End Sub
