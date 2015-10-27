VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCancelacionCierre 
   Caption         =   "Cancelar Cierre de Caja"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCancelacionCierre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbCaja 
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
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   690
      Width           =   2775
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmmm-yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1410
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
      Picture         =   "frmCancelacionCierre.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCancelar 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   930
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Cancelar"
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
      Picture         =   "frmCancelacionCierre.frx":055E
      PictureDisabled =   "frmCancelacionCierre.frx":0AB0
   End
   Begin VB.Label lblAjuste 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   1200
      TabIndex        =   8
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "AJUSTE:"
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
      TabIndex        =   7
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "IMPORTE:"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CAJA:"
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
      TabIndex        =   1
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FECHA A CANCELAR:"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1950
   End
End
Attribute VB_Name = "frmCancelacionCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim Bandera As Boolean

Private Sub cmbCaja_Click()
Dim rcCorte As New ADODB.Recordset

    If cmbCaja.ListIndex > -1 Then
                
        rcCorte.Open "SELECT Importe FROM auxiliar WHERE PC='" & Trim(cmbCaja.text) & "' AND Concepto='Corte de Caja' AND Fecha='" & Format(CDate(txtFecha.text), "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockReadOnly
        If Not rcCorte.BOF And Not rcCorte.EOF Then
        
            lblImporte.Caption = Format(rcCorte!Importe, FMoneda)
        End If
        rcCorte.Close
        
        rcCorte.Open "SELECT Importe FROM auxiliar WHERE PC='" & Trim(cmbCaja.text) & "' AND Concepto='Ajuste' AND Fecha='" & Format(CDate(txtFecha.text), "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockReadOnly
        If Not rcCorte.BOF And Not rcCorte.EOF Then
        
            lblAjuste.Caption = Format(rcCorte!Importe, FMoneda)
        End If
        rcCorte.Close
        Set rcCorte = Nothing
    Else
        
        lblImporte.Caption = Format(0, FMoneda)
        lblAjuste.Caption = Format(0, FMoneda)
    End If
End Sub

Private Sub cmbCaja_GotFocus()
    Cambiar_Color True, cmbCaja
End Sub

Private Sub cmbCaja_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbCaja_LostFocus()
    Cambiar_Color False, cmbCaja
End Sub

Private Sub cmdCancelar_Click()
    
    If cmbCaja.ListIndex > -1 Then
        
        dbDatos.Execute "UPDATE auxiliar SET corte=0 WHERE PC='" & Trim(cmbCaja.text) & "' AND Fecha='" & Format(CDate(txtFecha.text), "YYYY/MM/DD") & "'"
        dbDatos.Execute "UPDATE auxiliar SET Importe=0,Concepto=CONCAT(Concepto,' ','(CANCELADO)') WHERE PC='" & Trim(cmbCaja.text) & "' AND (Concepto='Corte de Caja' OR Concepto='Ajuste') AND Fecha='" & Format(CDate(txtFecha.text), "YYYY/MM/DD") & "'"
        Bandera = True
        MsgBox "Corte de caja cancelado con éxito !!", vbInformation, "Cancelaciones"
    Else
        
        Bandera = False
        
    End If
'''''    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Bandera = False
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtFecha_GotFocus()
    Seleccionar_Texto txtFecha
    Cambiar_Color True, txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFecha_LostFocus()
    Cambiar_Color False, txtFecha
    If Not IsDate(txtFecha.text) And Trim(txtFecha.text) <> "__/__/____" Then
        
        MsgBox "Introduzca una fecha válida !!", vbInformation, "Cancelaciones"
        txtFecha.SetFocus
        
    ElseIf IsDate(txtFecha.text) Then
        
        CargaCajas "DISTINCT PC", "auxiliar", cmbCaja, " WHERE Concepto='Corte de Caja' AND Fecha='" & Format(CDate(txtFecha.text), "YYYY/MM/DD") & "'"
        cmbCaja_Click
    End If
End Sub

Public Function Cancelar() As Boolean
    Me.Show vbModal
    Cancelar = Bandera
End Function

Sub CargaCajas(Campo As String, Tabla As String, Combo As ComboBox, Optional Condicion As String = "", Optional CampoOrdenamiento As String = "", Optional Limpiar As Boolean = True, Optional CampoClave As String = "ID")
Dim crCajas As New ADODB.Recordset
   
On Error GoTo error
    
    Combo.Clear
    crCajas.Open "SELECT " & Campo & " AS Valor FROM " & Tabla & IIf(Condicion = "", "", Condicion) & IIf(CampoOrdenamiento = "", "", " ORDER BY " & CampoOrdenamiento), dbDatos, adOpenForwardOnly, adLockReadOnly
    
    With crCajas
        
        While Not .EOF
            
            Combo.AddItem .Fields("Valor")
            'Combo.ItemData(Combo.NewIndex) = !idd
        .MoveNext
        Wend
    
    End With

    crCajas.Close
    Set crCajas = Nothing
    Exit Sub
    
error:
    Maneja_Error Err
    Set crCajas = Nothing
End Sub
