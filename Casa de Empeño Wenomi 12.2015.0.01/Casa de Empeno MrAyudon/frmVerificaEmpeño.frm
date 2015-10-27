VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmVerificaEmpeño 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Empeños"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerificaEmpeño.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   11865
   Begin VB.Frame frmSepararlote 
      Caption         =   "Lotes"
      Height          =   5220
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11715
      Begin VB.CheckBox chkAutomovil 
         Appearance      =   0  'Flat
         Caption         =   "Autos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtFolio 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin vbAcceleratorGrid6.vbalGrid grdArticulos 
         Height          =   3135
         Left            =   15
         TabIndex        =   2
         Top             =   2040
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   5530
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
         HeaderFlat      =   -1  'True
         BorderStyle     =   2
         ScrollBarStyle  =   2
         Editable        =   -1  'True
         DisableIcons    =   -1  'True
      End
      Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         AlignCaption    =   4
         AlignPicture    =   2
         AutoSize        =   0   'False
         Caption         =   "   &Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   16777215
         MaskColor       =   16777215
         MousePointer    =   1
         TextColor       =   -2147483630
         Object.ToolTipText     =   ""
         Picture         =   "frmVerificaEmpeño.frx":000C
      End
      Begin VB.Label lblAvaluo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   9840
         TabIndex        =   28
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label lblPrestamo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   9840
         TabIndex        =   27
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Avalúo:"
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
         Left            =   8400
         TabIndex        =   26
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Préstamo:"
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
         Left            =   8400
         TabIndex        =   25
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "Cp:"
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
         Left            =   5880
         TabIndex        =   22
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "Identificación:"
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
         Left            =   4560
         TabIndex        =   21
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
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
         Left            =   2280
         TabIndex        =   20
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
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
         TabIndex        =   19
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "Col:"
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
         TabIndex        =   18
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
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
         Left            =   2640
         TabIndex        =   17
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
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
         TabIndex        =   16
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
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
         TabIndex        =   15
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos:"
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
         Left            =   3840
         TabIndex        =   14
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   960
         TabIndex        =   13
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblApellido 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   4800
         TabIndex        =   12
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblDireccion 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1080
         TabIndex        =   11
         Top             =   960
         Width           =   7215
      End
      Begin VB.Label lblColonia 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   600
         TabIndex        =   10
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblMunicipio 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3540
         TabIndex        =   9
         Top             =   1320
         Width           =   2235
      End
      Begin VB.Label lblCp 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   6240
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   750
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblTelefono 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3165
         TabIndex        =   6
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label lblIdentificacion 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   5880
         TabIndex        =   5
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contrato:"
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
         TabIndex        =   4
         Top             =   240
         Width           =   795
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   10680
      TabIndex        =   23
      Top             =   5400
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
      Picture         =   "frmVerificaEmpeño.frx":0391
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   9480
      TabIndex        =   24
      Top             =   5400
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
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmVerificaEmpeño.frx":0422
   End
End
Attribute VB_Name = "frmVerificaEmpeño"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()
Dim Movimiento As Long, Folio As Long, Iniciales As String, Serie As Integer, Prestamo As Double, Hora As String

    If Val(txtFolio.Tag) > 0 Then

        With grdArticulos
        
            If MsgBox("Son correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Pago de Empeños") = vbNo Then Exit Sub
            Folio = txtFolio.text
            Iniciales = SacaIniciales
            Prestamo = lblPrestamo.Caption

            If chkAutomovil.Value = 0 Then
                
                Serie = SERIE_A
            Else
                
                Serie = SERIE_B
            End If
        
            'Saco el Movimiento
            Movimiento = Regresa_Movimiento(False, "Movimiento")
            Regresa_Movimiento True, "Movimiento"
            
            'Tomo la Hora
            Hora = Time
            
            'Grabamos el cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Empeño'," & Movimiento & "," & Folio & ",'" & Iniciales & "','201701'," & Prestamo & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
  
            'Grabamos el abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Empeño'," & Movimiento & "," & Folio & ",'" & Iniciales & "','110150'," & Prestamo & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
  
'''            'Grabamos abono 199450
'''            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'''                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Empeño'," & Movimiento & "," & Folio & ",'" & Iniciales & "','199450'," & Prestamo & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
            dbDatos.Execute "UPDATE empeno SET Verificado=1 WHERE ID=" & Val(txtFolio.Tag)
        
            'Abro el cajon del dinero
            Abrir_Cajon
        
            Limpiar "Lotes"
            chkAutomovil.Value = 0
            grdArticulos.ClearItems
            txtFolio.SetFocus
        End With

    End If

End Sub

Private Sub cmdBuscar_Click()
    txtFolio_KeyPress vbKeyReturn
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    frmSepararlote.BorderStyle = 0
    Crear_Encabezados
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Crear_Encabezados()

    With grdArticulos
        .AddColumn "K1", "Tipo", ecgHdrTextALignLeft, , 97, True, , , , , , CCLSortString
        .AddColumn "K2", "Cant.", ecgHdrTextALignRight, , 40, , , , , , , CCLSortNumeric
        .AddColumn "K3", "Prenda", ecgHdrTextALignLeft, , 170, , , , , , , CCLSortNumeric
        .AddColumn "K4", "Peso", ecgHdrTextALignRight, , 55, , , , , , , CCLSortNumeric
        .AddColumn "K5", "Kílates", ecgHdrTextALignRight, , 65, , , , , , , CCLSortString
        .AddColumn "K6", "Avalúo", ecgHdrTextALignRight, , 90, , , , , "###,###,###,###0.00", , CCLSortNumeric
        .AddColumn "K7", "Préstamo", ecgHdrTextALignRight, , 90, , , , , "###,###,###,###0.00", , CCLSortNumeric
        .AddColumn "K8", "Observaciones", ecgHdrTextALignLeft, , 150, , , , , , , CCLSortNumeric
        .Rows = 9
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtFolio_GotFocus()
    Seleccionar_Texto txtFolio
    Cambiar_Color True, txtFolio
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
Dim rcConsulta As New ADODB.Recordset
Dim Serie As Integer, Folio As Long
    
    KeyAscii = Solo_Numeros(KeyAscii)
    
    If chkAutomovil.Value = 0 Then
        
        Serie = SERIE_A
    Else
        
        Serie = SERIE_B
    End If

    If KeyAscii = vbKeyReturn And Trim(txtFolio.text) <> "" Then
            
            rcConsulta.Open "SELECT empeno.ID,empeno.Avaluo,empeno.Prestamo,clientes.Nombre,clientes.Apellido,clientes.Direccion,clientes.Colonia,clientes.Municipio,clientes.cp,clientes.Estado,clientes.Tel,clientes.Identificacion " _
                            & "FROM empeno LEFT JOIN clientes ON empeno.IDCliente=clientes.ID WHERE empeno.NumContrato=" & Val(txtFolio.text) & " AND Serie=" & Serie & " AND Pagado=0 AND Cancelado=0 AND Verificado=0 AND Origen=1", dbDatos, adOpenForwardOnly, adLockOptimistic
            If Not rcConsulta.BOF And Not rcConsulta.EOF Then
                
                With rcConsulta
                    
                    Folio = txtFolio.text
                    Limpiar "Lotes"
                    grdArticulos.ClearItems
                    txtFolio.text = Folio
                    txtFolio.Tag = !ID
                    lblNombre.Caption = !Nombre
                    lblApellido.Caption = !Apellido
                    lblDireccion.Caption = !Direccion
                    lblColonia.Caption = !Colonia
                    lblMunicipio.Caption = !Municipio
                    lblCP.Caption = !CP
                    lblEstado.Caption = !Estado
                    lblTelefono.Caption = !Tel
                    lblIdentificacion.Caption = !Identificacion
                    lblPrestamo.Caption = "$ " & Format(!Prestamo, "###,###,###,###0.00")
                    lblAvaluo.Caption = "$ " & Format(!Avaluo, "###,###,###,###0.00")
                    CargaPrendas !ID
                    
                End With
            
            Else
                
                MsgBox "No se encontró el contrato especificado !!", vbCritical, "Pago de Empeños"
                txtFolio.SetFocus
            
            End If
            rcConsulta.Close
            
    End If
    Set rcConsulta = Nothing
    
End Sub

Private Sub txtFolio_LostFocus()
    Cambiar_Color False, txtFolio
End Sub

Private Sub Limpiar(Contededor As String)
Dim ctrl As Control

    For Each ctrl In Controls
        
        On Error Resume Next
        If ctrl.Container.Caption = Contededor Then
            If TypeOf ctrl Is TextBox Then ctrl.text = "": ctrl.Tag = ""
            If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = ""
            If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
            On Error Resume Next
            ctrl.Tag = ""
        End If

    Next

End Sub

Function CargaPrendas(ID As Long)
Dim rcPrendas As New ADODB.Recordset
Dim i As Integer

On Error GoTo Error

    rcPrendas.Open "SELECT Tipo.Descripcion as TipoDescripcion,detallesempeno.IDEmpeno,detallesempeno.Cantidad,detallesempeno.Articulo,detallesempeno.Peso,kilatajes.Descripcion as Kilataje,detallesempeno.Avaluo,detallesempeno.Prestamo,detallesempeno.Estado,detallesempeno.Observaciones " _
                   & "FROM Tipo RIGHT JOIN detallesempeno ON tipo.ID=detallesempeno.Tipo LEFT JOIN kilatajes on detallesempeno.kilates=kilatajes.ID WHERE detallesempeno.IDEmpeno=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcPrendas.BOF And Not rcPrendas.EOF Then
        i = 1

        With grdArticulos
            While Not rcPrendas.EOF
                .CellText(i, 1) = rcPrendas!TipoDescripcion
            
                .CellText(i, 2) = rcPrendas!Cantidad
                .CellTextAlign(i, 2) = DT_RIGHT
            
                .CellText(i, 3) = rcPrendas!Articulo
    
                .CellText(i, 4) = rcPrendas!Peso
                .CellTextAlign(i, 4) = DT_RIGHT
            
                .CellText(i, 5) = rcPrendas!Kilataje
                .CellTextAlign(i, 5) = DT_RIGHT
            
                .CellText(i, 6) = rcPrendas!Avaluo
                .CellTextAlign(i, 6) = DT_RIGHT
            
                .CellText(i, 7) = rcPrendas!Prestamo
                .CellTextAlign(i, 7) = DT_RIGHT
            
                .CellText(i, 8) = rcPrendas!Observaciones
            i = i + 1
            rcPrendas.MoveNext
            Wend
        End With

    End If
    rcPrendas.Close
    
Error:
    Maneja_Error Err
    Set rcPrendas = Nothing
End Function

Private Function SacaIniciales() As String
Dim Cadena As String, Nombre As String, Apellidos As String
   
    Nombre = Trim(lblNombre.Caption)
    Apellidos = Trim(lblApellido.Caption)
   
    Cadena = Mid(Nombre, 1, 1)

    If InStr(1, Nombre, " ") <> 0 Then Cadena = Cadena & Mid(Nombre, InStr(1, Nombre, " ") + 1, 1)
   
    Cadena = Cadena & Mid(Apellidos, 1, 1)

    If InStr(1, Apellidos, " ") <> 0 Then Cadena = Cadena & Mid(Apellidos, InStr(1, Apellidos, " ") + 1, 1)
      
    SacaIniciales = Cadena
End Function
