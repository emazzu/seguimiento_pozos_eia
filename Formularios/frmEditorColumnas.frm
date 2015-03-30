VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEditorColumnas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edicion de campos"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMemo 
      Height          =   2295
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtMemo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   120
         MaxLength       =   32500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label lblMemo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese un valor alfanumerico"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2550
      End
   End
   Begin VB.CommandButton cmdGuardar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4335
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmEditorColumnas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2775
      UseMaskColor    =   -1  'True
      Width           =   285
   End
   Begin VB.CommandButton cmdReiniciar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4605
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmEditorColumnas.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2775
      UseMaskColor    =   -1  'True
      Width           =   285
   End
   Begin MSComctlLib.StatusBar info 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3210
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   530
            MinWidth        =   530
            Picture         =   "frmEditorColumnas.frx":0684
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   14993
            MinWidth        =   14993
            Text            =   "www.ConsultoresGIS.com"
            TextSave        =   "www.ConsultoresGIS.com"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraUbicacion 
      Height          =   1095
      Left            =   120
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.ComboBox cmbUbicacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEditorColumnas.frx":0C76
         Left            =   120
         List            =   "frmEditorColumnas.frx":0C78
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblUbicacion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione un nuevo valor"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.Frame fraBool 
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CheckBox chkBool 
         Caption         =   "No"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblBool 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese un nuevo valor"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1965
      End
   End
   Begin VB.Frame fraTexto 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtTexto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   255
         TabIndex        =   10
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese un nuevo valor alfanumerico"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3120
      End
   End
   Begin VB.Frame fraFecha 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   18219009
         CurrentDate     =   39804
      End
      Begin VB.Label lblFecha 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese una nueva Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2130
      End
   End
   Begin VB.Frame fraEntero 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtEntero 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   9
         TabIndex        =   5
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblEntero 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese un nuevo valor numerico entero"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3420
      End
   End
   Begin VB.Frame fraTipoYacimiento 
      Height          =   1455
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.OptionButton optDesarrollo 
         Caption         =   "Desarrollo"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optAvanzada 
         Caption         =   "De avanzada"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1065
         Width           =   1455
      End
      Begin VB.OptionButton optExploratorio 
         Caption         =   "Exploratorio"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   825
         Width           =   1335
      End
      Begin VB.Label lblTipoYacimiento 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione un valor"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Frame fraMoneda 
      Height          =   1095
      Left            =   120
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtMoneda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   9
         TabIndex        =   32
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblMoneda 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese un nuevo valor monetario"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraDecimal 
      Height          =   1095
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtDecimal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   9
         TabIndex        =   19
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblDecimal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese un nuevo valor numerico decimal"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3510
      End
   End
   Begin VB.Label etiCampo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Tag             =   "X"
      Top             =   120
      Width           =   3045
   End
   Begin VB.Label lblCampo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Edicion del campo:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Tag             =   "X"
      Top             =   120
      Width           =   1545
   End
End
Attribute VB_Name = "frmEditorColumnas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''Este form permite cambiar varios valores en simultaneo
Option Explicit

Public NombreColumna As String
Public Alias As String
Public Valor As Variant


Public Grilla As MSHFlexGrid
Private TipoDeDato As Long


Private Sub chkBool_Click()
'Cambia el texto del control chk para que sea mas claro apra el usuario
  On Error GoTo ErrorHandler
  
  chkBool.Caption = IIf(chkBool.Value = vbChecked, "Si", "No")
  chkBool.Refresh
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub cmdGuardar_Click()

'Guarda los cambios realizados para la seleccion de pozos
'  On Error GoTo ErrorHandler
  
  If VerificarCampos Then
  
    Select Case TipoDeDato
      Case enmTipoDeDato_TipoYacimiento
        BD.Execute "UPDATE POZOS SET " & NombreColumna & " = '" & ObtenerTipoYacimiento & "' WHERE IDPOZO IN (" & IDs & ")"
        
      Case enmTipoDeDato_Ubicacion
        BD.Execute "UPDATE POZOS SET " & NombreColumna & " = '" & cmbUbicacion & "', ORDENUBICACION = " & ObtenerOrdenUbicacion(cmbUbicacion) & " WHERE IDPOZO IN (" & IDs & ")"
      Case enmTipoDeDato_Entero
        BD.Execute "UPDATE POZOS SET " & NombreColumna & " = " & IIf(txtEntero = "", "NULL", txtEntero) & " WHERE IDPOZO IN (" & IDs & ")"
      Case enmTipoDeDato_Decimal
        BD.Execute "UPDATE POZOS SET " & NombreColumna & " = " & IIf(txtDecimal = "", "NULL", txtDecimal) & " WHERE IDPOZO IN (" & IDs & ")"
      Case enmTipoDeDato_Moneda
        BD.Execute "UPDATE POZOS SET " & NombreColumna & " = " & IIf(txtMoneda = "", "NULL", txtMoneda) & " WHERE IDPOZO IN (" & IDs & ")"
      Case enmTipoDeDato_Texto
        BD.Execute "UPDATE POZOS SET " & NombreColumna & " = " & IIf(txtTexto = "", "NULL", "'" & txtTexto & "'") & " WHERE IDPOZO IN (" & IDs & ")"
      Case enmTipoDeDato_Memo
        BD.Execute "UPDATE POZOS SET " & NombreColumna & " = " & IIf(txtMemo = "", "NULL", "'" & txtMemo & "'") & " WHERE IDPOZO IN (" & IDs & ")"
      Case enmTipoDeDato_Fecha
        BD.Execute "UPDATE POZOS SET " & NombreColumna & " = " & IIf(IsNull(dtpFecha.Value), "NULL", "#" & Format(dtpFecha.Value, "yyyy/mm/dd") & "#") & " WHERE IDPOZO IN (" & IDs & ")"
      Case enmTipoDeDato_Boolean
        BD.Execute "UPDATE POZOS SET " & NombreColumna & " = " & IIf(chkBool.Value = vbChecked, 1, 0) & " WHERE IDPOZO IN (" & IDs & ")"
    End Select
    
    'vergpTEST
    
'    MsgBox "11"
    
    BD.Execute "INSERT INTO HIST_POZOS (WELLID, Ubicacion, Equipo, MONOGRAFIA, ACUIFERO, Prognosis, Programa, Yacimiento, TIPOYACIMIENTO, Pozo, Prospect, FECHAPRIMERMONOGRAFIA, MONOGRAFIAS, FECHAULTIMAMONOGRAFIA, WELLINFORMED, INFORMEDBY, Definitiva, XPDC, YPDC, XPOS94, YPOS94, " & _
               " FECHASOLICITUD, FECHAPRIORIDAD, DOCUMENTOAPREPARAR, FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO, FIRSTPROD, " & _
               " TD, TOTDAYS, REMDAYS, Status, STARTDATE, ENDDATE, LANDOWNER, LANDOWNERPERMITDATE, " & _
               " Consult, Type, FechaPedidoEtia, FECHAESPERADAETIA, " & _
               " CONSULTANTRECOMENDATION, DMAPERMIT, Estado, " & _
               " IDMANIFIESTO, FECHAMANIFIESTO, FECHAENTREGAEIAXCONSULTORAOXY, EIAPRESENTADO, FECHAENVIOACS, FECHAPRESENTACIONDMA, FECHAPRESENTACIONSMA, PAGOTASAADMINISTRATIVA, FECHAINFOCOMPLEMENTARIA, TECHNICALREPORT, TASAADMINISTRATIVA, TASACONTRALOR, Estudio, Adenda, FECHAINICIODIA, FECHAFINDIA, INFORMEDEAVANCEDEOBRA50PORCIENTO, INFORMEDEAVANCEDEOBRA100PORCIENTO, INFORMEEVALUACIONARQUEOLOGICO, " & _
               " TIEMPOENTREPEDIDOETIAYRECEPCIONETIA, TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA, TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA, TIEMPOENTREPRESENTACIONANTEDMAYVISITA, TIEMPOENTREVISITAYAPROBACIONFINALDEDMA, TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA, TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA, TIEMPOENTREENTREGAEIADMAYAPROBACION, TipoMovimiento, XID, XFECHA, XUSUARIO)" & _
               " SELECT WELLID AS [Well ID], Ubicacion, Equipo, MONOGRAFIA AS [Monografía], ACUIFERO AS [Acuífero], Prognosis AS Prognosis, Programa AS Programa, Yacimiento, TIPOYACIMIENTO AS [Tipo Yacimiento], Pozo, Prospect, FECHAPRIMERMONOGRAFIA AS [Fecha Primera Monografia], MONOGRAFIAS AS [Monografías], FECHAULTIMAMONOGRAFIA AS [Fecha Ultima Monografía],  WELLINFORMED AS [Well Informed], INFORMEDBY AS [Informed By], Definitiva AS Definitiva, XPDC AS [X PDC], YPDC AS [Y PDC], XPOS94 AS [X POS94], YPOS94 AS [Y POS94], " & _
               " FECHASOLICITUD AS [Fecha Solicitud], FECHAPRIORIDAD as [Fecha Prioridad], DOCUMENTOAPREPARAR AS [Documento a Preparar], FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO AS [Fecha Entrega Dictamen Tco], FIRSTPROD AS [First Prod], " & _
               " TD, TOTDAYS AS [Tot Days], REMDAYS AS [Rem Days], Status, STARTDATE AS [Start Date], ENDDATE AS [End Date], LANDOWNER AS [Land Owner], LANDOWNERPERMITDATE AS [Land Owner Permit Date], " & _
               " Consult, Type, FechaPedidoEtia AS [Fecha Pedido ETIA], FechaEsperadaEtia AS [Fecha Esperada ETIA], " & _
               " CONSULTANTRECOMENDATION AS [Consultant Recomendation], DMAPERMIT AS [DMA Permit], Estado, " & _
               " IDMANIFIESTO AS [ID Manifiesto], FECHAMANIFIESTO AS [Fecha Manifiesto], FECHAENTREGAEIAXCONSULTORAOXY AS [Fecha Entrega EIA Por Consultor a OXY], EIAPRESENTADO AS [EIA Presentado], FECHAENVIOACS AS [Fecha Envio ACS], FECHAPRESENTACIONDMA AS [Fecha Presentación DMA], FECHAPRESENTACIONSMA AS [Fecha Presentación SMA], PAGOTASAADMINISTRATIVA AS [Pago Tasa Administrativa], FECHAINFOCOMPLEMENTARIA AS [Fecha Info Complementaria], TECHNICALREPORT AS [Technical Report], TASAADMINISTRATIVA AS [Tasa Administrativa], TASACONTRALOR AS [Tasa Contralor], Estudio, Adenda, FECHAINICIODIA AS [Fecha Inicio DIA], FECHAFINDIA AS [Fecha Fin DIA], INFORMEDEAVANCEDEOBRA50PORCIENTO AS [Informe de Avance de Obra 50%], INFORMEDEAVANCEDEOBRA100PORCIENTO AS [Informe de Avance de Obra 100%], INFORMEEVALUACIONARQUEOLOGICO AS [Informe Evaluacion Arqueologico], " & _
               " A.TIEMPOENTREPEDIDOETIAYRECEPCIONETIA AS [Tiempo Entre Pedido ETIA y Recepcion ETIA], A.TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA AS [Tiempo Entre Primer Monografía y Recepción ETIA], A.TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA AS [Tiempo Entre Recepcion ETIA y Presentación ante DMA], A.TIEMPOENTREPRESENTACIONANTEDMAYVISITA AS [Tiempo Entre Presentación Ante DMA y Visita], A.TIEMPOENTREVISITAYAPROBACIONFINALDEDMA AS [Tiempo Entre Visita y Aprobación Final De DMA], A.TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA AS [Tiempo Entre Presentación De Pozo y Aprobación Final DMA], A.TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA AS [Tiempo Entre Primer Monografia y Pedido EIA], A.TIEMPOENTREENTREGAEIADMAYAPROBACION AS [Tiempo Entre Entrega EIA a DMA y Aprobacion], 'M', IDPOZO, #" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "#,'" & InfoGlobal.Usuario & "' FROM POZOS A WHERE IDPOZO IN (" & IDs & ")"
    
    
'    Unload Me
'    frmPozosFuturos.cmdBuscar_Click
  
  End If
  
'ErrorHandler:
'  ErrorHandler
  
End Sub



Private Function IDs() As String
'Genera una lista con los IDs que se van a usar en el where de la modificacion
  On Error GoTo ErrorHandler
  
  Dim i As Long
  Dim aux As Long
  
  If Grilla.Row > Grilla.RowSel Then
    aux = Grilla.Row
    Grilla.Row = Grilla.RowSel
    Grilla.RowSel = aux
  End If
  For i = Grilla.Row To Grilla.RowSel
    If IsNumeric(Grilla.TextMatrix(i, 0)) Then
      IDs = IDs & "," & Grilla.TextMatrix(i, 0)
    End If
  Next i
  IDs = Mid(IDs, 2)
  
ErrorHandler:
  ErrorHandler
End Function

Private Function VerificarCampos() As Boolean
'Verifica segun el tipo de dato que se solicita que todo este correcto para guardar

  Select Case TipoDeDato
    Case enmTipoDeDato_Entero
      If txtEntero <> "" Then
        If IsNumeric(txtEntero) Then
          If InStr(1, txtEntero, ",") = 0 Then
            If CLng(txtEntero) > 0 Then
              VerificarCampos = True
            Else
              MsgBox "El campo debe ser mayor que cero", vbCritical + vbOKOnly
              txtEntero.SetFocus
            End If
          Else
            MsgBox "El campo debe ser un numero entero", vbCritical + vbOKOnly
            txtEntero.SetFocus
          End If
        Else
          MsgBox "El campo debe ser un numero entero", vbCritical + vbOKOnly
          txtEntero.SetFocus
        End If
      Else
        VerificarCampos = True
      End If
        
    Case enmTipoDeDato_Decimal
      If txtDecimal <> "" Then
        If IsNumeric(txtDecimal) Then
          If CDbl(txtDecimal) > 0 Then
            VerificarCampos = True
          Else
            MsgBox "El campo debe ser mayor que cero", vbCritical + vbOKOnly
            txtDecimal.SetFocus
          End If
        Else
          MsgBox "El campo debe ser un numero decimal", vbCritical + vbOKOnly
          txtDecimal.SetFocus
        End If
      Else
        VerificarCampos = True
      End If
      
    Case enmTipoDeDato_Moneda
      If txtMoneda <> "" Then
        If IsNumeric(txtMoneda) Then
          txtMoneda = CCur(txtMoneda)
          If CCur(txtMoneda) > 0 Then
            VerificarCampos = True
          Else
            MsgBox "El campo debe ser mayor que cero", vbCritical + vbOKOnly
            txtMoneda.SetFocus
          End If
        Else
          MsgBox "El campo debe ser un numero decimal", vbCritical + vbOKOnly
          txtMoneda.SetFocus
        End If
      Else
        VerificarCampos = True
      End If
      
    Case enmTipoDeDato_Texto, enmTipoDeDato_Memo, enmTipoDeDato_Fecha, enmTipoDeDato_Boolean, enmTipoDeDato_TipoYacimiento, enmTipoDeDato_Ubicacion
      VerificarCampos = True
  End Select
  
ErrorHandler:
  ErrorHandler
End Function

Private Sub cmdReiniciar_Click()
'Reinicia el formulario al hacer click en el boton de reiniciar
  On Error GoTo ErrorHandler
  
  ReiniciarFormulario Me
  
ErrorHandler:
  ErrorHandler
End Sub




Private Sub Form_Load()
'Prepara el form para su uso
  On Error GoTo ErrorHandler
  
  Center Me
  CargarCombos
  TipoDeDato = ObtenerTipoDeDato(NombreColumna)
  etiCampo = Alias
  Select Case TipoDeDato
    Case enmTipoDeDato_Ubicacion
      fraUbicacion.Visible = True
      cmdGuardar.Top = fraEntero.Top + fraEntero.Height + 100
      cmdReiniciar.Top = fraEntero.Top + fraEntero.Height + 100
      Me.Height = cmdReiniciar.Top + cmdReiniciar.Height + 800
    Case enmTipoDeDato_TipoYacimiento
      fraTipoYacimiento.Visible = True
      cmdGuardar.Top = fraTipoYacimiento.Top + fraTipoYacimiento.Height + 100
      cmdReiniciar.Top = fraTipoYacimiento.Top + fraTipoYacimiento.Height + 100
      Me.Height = cmdReiniciar.Top + cmdReiniciar.Height + 800
    Case enmTipoDeDato_Entero
      fraEntero.Visible = True
      cmdGuardar.Top = fraEntero.Top + fraEntero.Height + 100
      cmdReiniciar.Top = fraEntero.Top + fraEntero.Height + 100
      Me.Height = cmdReiniciar.Top + cmdReiniciar.Height + 800
    Case enmTipoDeDato_Decimal
      fraDecimal.Visible = True
      cmdGuardar.Top = fraEntero.Top + fraEntero.Height + 100
      cmdReiniciar.Top = fraEntero.Top + fraEntero.Height + 100
      Me.Height = cmdReiniciar.Top + cmdReiniciar.Height + 800
    Case enmTipoDeDato_Moneda
      fraMoneda.Visible = True
      cmdGuardar.Top = fraEntero.Top + fraEntero.Height + 100
      cmdReiniciar.Top = fraEntero.Top + fraEntero.Height + 100
      Me.Height = cmdReiniciar.Top + cmdReiniciar.Height + 800
    Case enmTipoDeDato_Texto
      fraTexto.Visible = True
      cmdGuardar.Top = fraEntero.Top + fraEntero.Height + 100
      cmdReiniciar.Top = fraEntero.Top + fraEntero.Height + 100
      Me.Height = cmdReiniciar.Top + cmdReiniciar.Height + 800
    Case enmTipoDeDato_Memo
      fraMemo.Visible = True
    Case enmTipoDeDato_Fecha
      fraFecha.Visible = True
      cmdGuardar.Top = fraEntero.Top + fraEntero.Height + 100
      cmdReiniciar.Top = fraEntero.Top + fraEntero.Height + 100
      Me.Height = cmdReiniciar.Top + cmdReiniciar.Height + 800
      dtpFecha.CheckBox = IIf(Valor = "", True, vbChecked)
      dtpFecha.Value = IIf(Valor = "", Date, Valor)
      dtpFecha.Value = IIf(Valor = "", Null, Valor)
    Case enmTipoDeDato_Boolean
      fraBool.Visible = True
      cmdGuardar.Top = fraEntero.Top + fraEntero.Height + 100
      cmdReiniciar.Top = fraEntero.Top + fraEntero.Height + 100
      Me.Height = cmdReiniciar.Top + cmdReiniciar.Height + 800
  End Select
  
ErrorHandler:
  ErrorHandler
End Sub

Private Function ObtenerTipoDeDato(NombreColumna As String) As Long
'Obtiene el tipo de dato de un atributo de la tabla
  On Error GoTo ErrorHandler
  Dim rs As New Recordset
  
  
  Select Case Trim(UCase(NombreColumna))
    Case "TIPOYACIMIENTO": ObtenerTipoDeDato = enmTipoDeDato_TipoYacimiento
    Case "UBICACION": ObtenerTipoDeDato = enmTipoDeDato_Ubicacion
    Case Else
      rs.CursorLocation = adUseClient
      rs.Open "SELECT TOP 1 " & NombreColumna & " FROM POZOS", BD, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs.EOF Then
        ObtenerTipoDeDato = rs.Fields(0).Type
      Else
        ObtenerTipoDeDato = -1
      End If
      rs.Close
  End Select
  
ErrorHandler:
  ErrorHandler
End Function

Private Sub CargarCombos()
'Carga el combo de ubicaciones por si ese campo se modificara
 On Error GoTo ErrorHandler
  Dim rs As New Recordset
  
  rs.Open "SELECT DISTINCT UBICACION FROM POZOS WHERE NOT UBICACION IS NULL", BD, adOpenDynamic, adLockOptimistic, adCmdText
  While Not rs.EOF
    cmbUbicacion.AddItem rs(0)
    rs.MoveNext
  Wend
  If cmbUbicacion.ListCount > 0 Then
    cmbUbicacion.ListIndex = 0
  Else
    cmbUbicacion.Text = ""
  End If
  rs.Close
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub txtDecimal_Change()
'Reemplaza el punto por la coma en txtDecimal
  On Error GoTo ErrorHandler
  
  txtDecimal = Replace(txtDecimal, ",", ".")
  
ErrorHandler:
  ErrorHandler
End Sub

Private Function ObtenerTipoYacimiento() As String
'Devuelve la letra correspondiente al tipo de yacimiento segun los opt
  On Error GoTo ErrorHandler
  
  If optDesarrollo.Value = True Then
    ObtenerTipoYacimiento = ""
  ElseIf optExploratorio.Value = True Then
    ObtenerTipoYacimiento = "X"
  ElseIf optAvanzada.Value = True Then
    ObtenerTipoYacimiento = "A"
  End If
  
ErrorHandler:
  ErrorHandler
End Function

