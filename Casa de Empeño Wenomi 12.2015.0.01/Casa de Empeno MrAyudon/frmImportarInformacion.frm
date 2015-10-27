VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmImportarInformacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Informacion"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImportarInformacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   7335
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.ComboBox cmbSucursal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox ChkImprimir 
         Caption         =   "Reporte Impreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox ChkImportar 
         Caption         =   "Importar Datos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
         Height          =   300
         Left            =   6360
         TabIndex        =   3
         Top             =   360
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
         Picture         =   "frmImportarInformacion.frx":000C
      End
      Begin VB.Label lblSaldo 
         AutoSize        =   -1  'True
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
         Left            =   5160
         TabIndex        =   6
         Top             =   960
         Width           =   75
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
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
         Left            =   4320
         TabIndex        =   5
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
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
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal:"
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
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1110
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   2760
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
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmImportarInformacion.frx":0121
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AutoSize        =   0   'False
      Caption         =   "&Aceptar"
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
   End
End
Attribute VB_Name = "frmImportarInformacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.I. Jorge Gabriel Colio Ramos
' Mazatlan, Sin. 10/08/2002
' Modulo frmImportarInformacion - frmImportarInformacion.frm
' Ultima Modificacion - 16/08/2002
'
'////////////////////////////////////////////////////////////////



Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()
'On Error GoTo Error
'checa si el combo box y la caja de texto estan con informacion
If cmbSucursal.Text = "" Or txtFecha.Text = "" Then
   MsgBox ("Datos incompletos, favor de rellenar las cajas de texto correspodientes")
   If cmbSucursal.Text = "" Then
      cmbSucursal.SetFocus
   ElseIf txtFecha.Text = "" Then
      txtFecha.SetFocus
   End If
   Exit Sub
End If

'Checa que alguno de los check list este seleccionado
If ChkImportar.Value = 0 And ChkImprimir.Value = 0 Then
   MsgBox ("No ha seleccionado ninguna opcion")
ElseIf ChkImportar.Value = 1 And ChkImprimir.Value = 1 Then
   Importar
   Sleep 1000
   Imprimir_Reportes (txtFecha.Text)
   ChkImportar.Value = 0
   ChkImportar.Enabled = False
ElseIf ChkImportar.Value = 1 And ChkImprimir.Value = 0 Then
   Importar
   ChkImportar.Value = 0
   ChkImportar.Enabled = False
ElseIf ChkImportar.Value = 0 And ChkImprimir.Value = 1 Then
   Imprimir_Reportes (txtFecha.Text)
End If

Error:
   Maneja_Error Err

End Sub

Private Sub cmdMosFecha_Click()
       txtFecha.Text = frmCalendario.Fecha(txtFecha.Text)
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub


Private Sub Form_Load()
    inicializar
End Sub

'Inicializamos la forma
Private Sub inicializar()
  Dim rcSucursal As New ADODB.Recordset
   
   Screen.MousePointer = vbHourglass
   Me.Top = 0
   Me.Left = 0
   Poner_Flat Fl, Me.Controls, Me
   Screen.MousePointer = vbDefault
   
  'Relleno el combo box de sucursales
  rcSucursal.Open "SELECT DISTINCT Sucursal FROM Importacion", dbDatos, adOpenDynamic, adLockOptimistic
  cmbSucursal.Clear
  With rcSucursal
    While Not .EOF
      cmbSucursal.AddItem !Sucursal
      .MoveNext
    Wend
    .Close
  End With
  
  Screen.MousePointer = default

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Quitar_Flat Fl
   Screen.MousePointer = vbDefault
End Sub

Private Sub Imprimir_Reportes(Fecha As String)
On Error GoTo Error

Dim rcImp As New ADODB.Recordset
Dim SA, Debe, Haber, Depositos, Cheques, SN As Currency

rcImp.Open "SELECT SUM(Importe) AS SaldoAnterior FROM Importacion WHERE Cuenta='SA' AND Sucursal='" & cmbSucursal.Text & "' AND Fecha=#" & Format(txtFecha.Text, "MM/DD/YY") & "#", dbDatos, adOpenDynamic, adLockOptimistic
SA = rcImp!SaldoAnterior
rcImp.Close

rcImp.Open "SELECT SUM(Importe) AS Debe FROM Importacion WHERE Cuenta='DEBE' AND Sucursal='" & cmbSucursal.Text & "' AND Fecha=#" & Format(txtFecha.Text, "MM/DD/YY") & "#", dbDatos, adOpenDynamic, adLockOptimistic
Debe = rcImp!Debe
rcImp.Close

rcImp.Open "SELECT SUM(Importe) AS Haber FROM Importacion WHERE Cuenta='HABER' AND Sucursal='" & cmbSucursal.Text & "' AND Fecha=#" & Format(txtFecha.Text, "MM/DD/YY") & "#", dbDatos, adOpenDynamic, adLockOptimistic
Haber = rcImp!Haber
rcImp.Close

rcImp.Open "SELECT SUM(Importe) AS Depositos FROM Importacion WHERE Cuenta='DEPOSITOS' AND Sucursal='" & cmbSucursal.Text & "' AND Fecha=#" & Format(txtFecha.Text, "MM/DD/YY") & "#", dbDatos, adOpenDynamic, adLockOptimistic
Depositos = rcImp!Depositos
rcImp.Close

rcImp.Open "SELECT SUM(Importe) AS Cheques FROM Importacion WHERE Cuenta='CHEQUES' AND Sucursal='" & cmbSucursal.Text & "' AND Fecha=#" & Format(txtFecha.Text, "MM/DD/YY") & "#", dbDatos, adOpenDynamic, adLockOptimistic
Cheques = rcImp!Cheques
rcImp.Close

rcImp.Open "SELECT SUM(Importe) AS SaldoNuevo FROM Importacion WHERE Cuenta='SN' AND Sucursal='" & cmbSucursal.Text & "' AND Fecha=#" & Format(txtFecha.Text, "MM/DD/YY") & "#", dbDatos, adOpenDynamic, adLockOptimistic
SN = rcImp!SaldoNuevo
rcImp.Close
 
  'imprimimos el reporte de corte de caja
  With frmMDI.Cr
    .Reset
    .DiscardSavedData = True
    .DataFiles(0) = path & "\Base De Datos\Datos.mdb"
    .DataFiles(1) = path & "\Base De Datos\Datos.mdb"
    .DataFiles(2) = path & "\Base De Datos\Datos.mdb"
    .password = Chr(10) & "administrativo"
    .ReportFileName = path & "\Reportes\CierreDiarioImp.rpt"
    .SelectionFormula = "{Importacion.Fecha}=date(" & Format(CDate(txtFecha.Text), "YYYY,MM,DD") & ") AND {Importacion.Sucursal}='" & cmbSucursal.Text & "'"
    .Formulas(0) = "Sucursal='" & IIf(crSucursal = "", cmbSucursal.Text, crSucursal) & "'"
    .Formulas(1) = "Fecha='" & IIf(crFecha = "", txtFecha.Text, crFecha) & "'"
    .Formulas(2) = "SaldoAnterior=" & SN & ""
    .Formulas(3) = "Debe=" & Debe & ""
    .Formulas(4) = "Haber=" & Haber & ""
    .Formulas(5) = "Depositos=" & Depositos & ""
    .Formulas(6) = "Cheques=" & Cheques & ""
    .Formulas(7) = "efectivo=" & SN & ""
    .WindowShowPrintSetupBtn = True
    .Destination = crptToWindow
    .Action = 1
  End With
  
Error:
   Maneja_Error Err
      
   Set rcImp = Nothing
End Sub

Private Sub Importar()
   Screen.MousePointer = vbHourglass
   Dim Archivo As String
   Dim rcImportar As New ADODB.Recordset
   Dim rcCuentas As New ADODB.Recordset
   Dim Nombre, NoSucursal, Fecha, Hora As String
   Dim SaldoAnterior, Debe, Haber, Depositos, Cheques, importe, SaldoNuevo As Long
   Dim Boveda, Bancos, empeño, Joyeria, Apartados, Faltante, Prestamos As Long
   Dim Mov As Integer
      
   Archivo = frmMDI.cm.FileName
   
   'checamos que la cadena de busqueda contega una ruta hacia el archivo
   If Archivo = "" Then GoTo Error
   
   Nombre = Regresa_Valor_Exp("SUCURSAL", "Nombre", "", Archivo)
   NoSucursal = Regresa_Valor_Exp("SUCURSAL", "NoSucursal", "", Archivo)
   Fecha = Regresa_Valor_Exp("SUCURSAL", "Fecha", "", Archivo)
   Hora = Regresa_Valor_Exp("SUCURSAL", "Hora", "", Archivo)

   
   SaldoAnterior = Regresa_Valor_Exp("SALDOS", "Saldo Anterior", "0", Archivo)
   Haber = Regresa_Valor_Exp("SALDOS", "Haber", "0", Archivo)
   Debe = Regresa_Valor_Exp("SALDOS", "Debe", "0", Archivo)
   Depositos = Regresa_Valor_Exp("SALDOS", "Depositos", "0", Archivo)
   Cheques = Regresa_Valor_Exp("SALDOS", "Cheques", "0", Archivo)
   SaldoNuevo = Regresa_Valor_Exp("SALDOS", "Saldo Nuevo", "0", Archivo)
   
   If Nombre = "" Then
    MsgBox ("El archivo a importar no contiene los datos del Auxiliar Diario")
    Exit Sub
   End If
   
   Boveda = Regresa_Valor_Exp("FINANCIERO", "Boveda", "0", Archivo)
   Bancos = Regresa_Valor_Exp("FINANCIERO", "Bancos", "0", Archivo)
   empeño = Regresa_Valor_Exp("FINANCIERO", "Empeño", "0", Archivo)
   Joyeria = Regresa_Valor_Exp("FINANCIERO", "Joyeria", "0", Archivo)
   Apartados = Regresa_Valor_Exp("FINANCIERO", "Apartados", "0", Archivo)
   Faltante = Regresa_Valor_Exp("FINANCIERO", "Faltante", "0", Archivo)
   Prestamos = Regresa_Valor_Exp("FINANCIERO", "Prestamos", "0", Archivo)
   
   If Boveda = "" Then
    MsgBox ("El archivo a importar no contiene los datos del Reporte Financiero")
    Exit Sub
   End If
   
   dbDatos.Execute "INSERT INTO Importacion (Sucursal,NoSucursal,Fecha,Hora,Cuenta,Importe) VALUES " & _
      "('" & Nombre & "'," & Val(NoSucursal) & ", #" & Format(Fecha, "MM/DD/YY") & "#,#" & Format(Hora, "HH:MM:SS") & "#,'SA'," & SaldoAnterior & ")"
      
   dbDatos.Execute "INSERT INTO Importacion (Sucursal,NoSucursal,Fecha,Hora,Cuenta,Importe) VALUES " & _
      "('" & Nombre & "'," & Val(NoSucursal) & ", #" & Format(Fecha, "MM/DD/YY") & "#,#" & Format(Hora, "HH:MM:SS") & "#,'DEBE'," & Debe & ")"
      
   dbDatos.Execute "INSERT INTO Importacion (Sucursal,NoSucursal,Fecha,Hora,Cuenta,Importe) VALUES " & _
      "('" & Nombre & "'," & Val(NoSucursal) & ", #" & Format(Fecha, "MM/DD/YY") & "#,#" & Format(Hora, "HH:MM:SS") & "#,'HABER'," & Haber & ")"
      
   dbDatos.Execute "INSERT INTO Importacion (Sucursal,NoSucursal,Fecha,Hora,Cuenta,Importe) VALUES " & _
      "('" & Nombre & "'," & Val(NoSucursal) & ", #" & Format(Fecha, "MM/DD/YY") & "#,#" & Format(Hora, "HH:MM:SS") & "#,'DEPOSITOS'," & Depositos & ")"
      
   dbDatos.Execute "INSERT INTO Importacion (Sucursal,NoSucursal,Fecha,Hora,Cuenta,Importe) VALUES " & _
      "('" & Nombre & "'," & Val(NoSucursal) & ", #" & Format(Fecha, "MM/DD/YY") & "#,#" & Format(Hora, "HH:MM:SS") & "#,'CHEQUES'," & Cheques & ")"
      
   dbDatos.Execute "INSERT INTO Importacion (Sucursal,NoSucursal,Fecha,Hora,Cuenta,Importe) VALUES " & _
      "('" & Nombre & "'," & Val(NoSucursal) & ", #" & Format(Fecha, "MM/DD/YY") & "#,#" & Format(Hora, "HH:MM:SS") & "#,'SN'," & SaldoNuevo & ")"
      
   
   rcCuentas.Open "SELECT * FROM Cuentas", dbDatos, adOpenDynamic, adLockOptimistic
   
   With rcCuentas
      While Not .EOF
         If Len(Regresa_Valor_Exp("CUENTAS", !Cuenta, "", Archivo)) > 0 Then
            importe = CCur(Regresa_Valor_Exp("CUENTAS", !Cuenta, "", Archivo))
            Mov = Val(Regresa_Valor_Exp("MOVIMIENTOS", !Cuenta, "", Archivo))
            If Right(!Cuenta, 2) = "01" Then
               dbDatos.Execute "INSERT INTO Importacion (Sucursal,NoSucursal,Fecha,Hora,Cuenta,Movimientos,Cargo) VALUES " & _
                  "('" & Nombre & "'," & Val(NoSucursal) & ", #" & Format(Fecha, "MM/DD/YY") & "#,#" & Format(Hora, "HH:MM:SS") & "#,'" & !Cuenta & "'," & Mov & "," & importe & ")"
            Else
               dbDatos.Execute "INSERT INTO Importacion (Sucursal,NoSucursal,Fecha,Hora,Cuenta,Movimientos,Abono) VALUES " & _
                  "('" & Nombre & "'," & Val(NoSucursal) & ", #" & Format(Fecha, "MM/DD/YY") & "#,#" & Format(Hora, "HH:MM:SS") & "#,'" & !Cuenta & "'," & Mov & "," & importe & ")"
            End If
         End If
         .MoveNext
      Wend
      .Close
   End With
   
   'Pasamos del txt el Reporte Financiero
   dbDatos.Execute "INSERT INTO Financiero (Sucursal,Fecha,Boveda,Bancos,Empeño,Joyeria,Apartados,Faltante,Prestamos) VALUES " & _
      "(" & Val(NoSucursal) & ",#" & Format(Fecha, "MM/DD/YY") & "#," & CCur(Boveda) & "," & CCur(Bancos) & "," & CCur(empeño) & "," & CCur(Joyeria) & "," & CCur(Apartados) & "," & CCur(Faltante) & "," & CCur(Prestamos) & ")"
   
   MsgBox "El Archivo ha sido importado exitosamente", vbOKOnly + vbInformation
   
   Kill Archivo
        
Error:
   Maneja_Error Err
   
   Set rcImportar = Nothing
   
   Screen.MousePointer = vbDefault
End Sub


