VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInformeHistoriales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historiales"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInformeHistoriales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   17400
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtWellID 
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
      Left            =   6180
      TabIndex        =   6
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton cmdWellID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10995
      Picture         =   "frmInformeHistoriales.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   97
      Width           =   330
   End
   Begin VB.ComboBox cmbHistoriales 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin MSComctlLib.StatusBar info 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   10125
      Width           =   17400
      _ExtentX        =   30692
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   530
            MinWidth        =   530
            Picture         =   "frmInformeHistoriales.frx":054E
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   35278
            MinWidth        =   35278
            Text            =   "www.ConsultoresGIS.com"
            TextSave        =   "www.ConsultoresGIS.com"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgItems 
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   5953
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgDetalleHistorial 
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   9340
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblWellID 
      Caption         =   "Well ID"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   150
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle de operaciones del item seleccionado"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   4140
   End
   Begin VB.Label lblItems 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Haga doble click en un item de la lista para ver su detalle"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5325
   End
   Begin VB.Label lblHistoriales 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ver historiales de "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   1680
   End
End
Attribute VB_Name = "frmInformeHistoriales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''
'Consultas de historiales'
''''''''''''''''''''''''''
Private Type TConsulta 'Tipo de dato que maneja la info de las consultas
  Nombre As String
  SQLSelect As String
  SQLFrom As String
End Type
Private arConsultas(0) As TConsulta

Private Sub cmdWellID_Click()
'Ejecuta la consulta
  On Error GoTo ErrorHandler
  
  EjecutarConsulta
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub Form_Load()
'Prepara el formulario
  On Error GoTo ErrorHandler
  
  Center Me
  CargarConsultasHistorial
  CargarComboHistoriales
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub CargarConsultasHistorial()
'Carga las consultas
  On Error GoTo ErrorHandler
  'RESPETAR QUE TODAS LAS TABLAS DEBEN LLEVAR ALIAS AUNQUE NO TENGAN JOINS, SI LOS TIENEN LOS JOINS SE DEBEN
  'HACER CONTRA LAS TABLAS DE HISTORIAL Y NO LAS ORIGINALES COMPARANDO EL CAMPO ID DE UNA CON EL XID DE OTRA
  
  arConsultas(0).Nombre = "Pozos"
  
  'vergpTEST
'  MsgBox "9"
  
  arConsultas(0).SQLSelect = "A.XID AS [Id], A.WELLID AS [Well ID], A.Ubicacion, A.Equipo, IIF(A.MONOGRAFIA,'Si','No') AS [Monografía], IIF(A.ACUIFERO,'Si','No') AS [Acuífero], IIF(A.Prognosis,'Si','No') AS [Plan de 8], IIF(A.Programa,'Si','No') AS [Plan de 13], A.Yacimiento, A.TIPOYACIMIENTO AS [Tipo Yacimiento], A.Pozo, A.Prospect, A.FECHAPRIMERMONOGRAFIA AS [Fecha Primera Monografia], A.MONOGRAFIAS AS [Monografías], A.FECHAULTIMAMONOGRAFIA AS [Fecha Ultima Monografía], A.WELLINFORMED AS [Well Informed], A.INFORMEDBY AS [Informed By], IIF(A.Definitiva,'Si','No') AS Definitiva, A.XPDC AS [X PDC], A.YPDC AS [Y PDC], A.XPOS94 AS [X POS94], A.YPOS94 AS [Y POS94], " & _
                            " A.FECHASOLICITUD AS [Fecha Solicitud], A.FECHAPRIORIDAD as [Fecha Prioridad], A.DOCUMENTOAPREPARAR AS [Documento a Preparar], A.FieldManifold, A.BatteryAssigned, A.FECHAENTREGADICTAMENTECNICO AS [Fecha Entrega Dictamen Tco], A.FIRSTPROD AS [First Prod], " & _
                            " A.TD, A.TOTDAYS AS [Tot Days], A.REMDAYS AS [Rem Days], A.Status, A.STARTDATE AS [Start Date], A.ENDDATE AS [End Date], A.LANDOWNER AS [Land Owner], A.LANDOWNERPERMITDATE AS [Land Owner Permit Date], " & _
                            " A.Consult, A.Type, A.FechaPedidoEtia AS [Fecha Pedido ETIA], A.FechaEsperadaEtia AS [Fecha Esperada ETIA], " & _
                            " A.CONSULTANTRECOMENDATION AS [Consultant Recomendation], A.DMAPERMIT AS [DMA Permit], A.Estado, " & _
                            " A.IDMANIFIESTO AS [ID Manifiesto], A.FECHAMANIFIESTO AS [Fecha Manifiesto], A.FECHAENTREGAEIAXCONSULTORAOXY AS [Fecha Entrega EIA Por Consultor a OXY], IIF(A.EIAPRESENTADO,'Si','No') AS [EIA Presentado], A.FECHAENVIOACS AS [Fecha Envio ACS], A.FECHAPRESENTACIONDMA AS [Fecha Presentación DMA], A.FECHAPRESENTACIONSMA AS [Fecha Presentación SMA], A.PAGOTASAADMINISTRATIVA AS [Pago Tasa Administrativa], A.FECHAINFOCOMPLEMENTARIA AS [Fecha Info Complementaria], A.TECHNICALREPORT AS [Technical Report], A.TASAADMINISTRATIVA AS [Tasa Administrativa], A.TASACONTRALOR AS [Tasa Contralor], A.Estudio, A.Adenda, A.FECHAINICIODIA AS [Fecha Inicio DIA], A.FECHAFINDIA AS [Fecha Fin DIA], IIF(A.INFORMEDEAVANCEDEOBRA50PORCIENTO,'Si','No') AS [Informe de Avance de Obra 50%], IIF(A.INFORMEDEAVANCEDEOBRA100PORCIENTO,'Si','No') AS [Informe de Avance de Obra 100%], IIF(A.INFORMEEVALUACIONARQUEOLOGICO,'Si','No') AS [Informe Evaluacion Arqueologico], " & _
                            " A.TIEMPOENTREPEDIDOETIAYRECEPCIONETIA AS [Tiempo Entre Pedido ETIA y Recepcion ETIA], A.TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA AS [Tiempo Entre Primer Monografía y Recepción ETIA], A.TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA AS [Tiempo Entre Recepcion ETIA y Presentación ante DMA], A.TIEMPOENTREPRESENTACIONANTEDMAYVISITA AS [Tiempo Entre Presentación Ante DMA y Visita], A.TIEMPOENTREVISITAYAPROBACIONFINALDEDMA AS [Tiempo Entre Visita y Aprobación Final De DMA], A.TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA AS [Tiempo Entre Presentación De Pozo y Aprobación Final DMA],  A.TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA AS [Tiempo Entre Primer Monografia y Pedido EIA], A.TIEMPOENTREENTREGAEIADMAYAPROBACION AS [Tiempo Entre Entrega EIA a DMA y Aprobacion]"
  arConsultas(0).SQLFrom = "HIST_POZOS A"
      
ErrorHandler:
  ErrorHandler
End Sub

Private Sub CargarComboHistoriales()
'Carga los nombres de las consultas en el combo
  On Error GoTo ErrorHandler
  
  Dim i As Long
  For i = LBound(arConsultas) To UBound(arConsultas)
    cmbHistoriales.AddItem arConsultas(i).Nombre
  Next i
  If cmbHistoriales.ListCount > 0 Then cmbHistoriales.ListIndex = 0
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub cmbHistoriales_Click()
'Carga los items de la seleccion del combo
  On Error GoTo ErrorHandler
  
  EjecutarConsulta
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub EjecutarConsulta()
'Carga los items de la seleccion del combo
  On Error GoTo ErrorHandler
  Dim strSQL As String
  
  mfgItems.Cols = 0
  mfgItems.Rows = 0
  mfgDetalleHistorial.Cols = 0
  mfgDetalleHistorial.Rows = 0
  strSQL = "SELECT " & arConsultas(cmbHistoriales.ListIndex).SQLSelect & " FROM " & arConsultas(cmbHistoriales.ListIndex).SQLFrom & " WHERE A.TIPOMOVIMIENTO = 'A' " & IIf(Trim(txtWellID) <> "", " AND UCASE(A.WELLID) LIKE '%" & UCase(txtWellID) & "%'", "")
  ProcesarBusqueda2 strSQL, mfgItems, "A.UBICACION"
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub mfgItems_DblClick()
'Carga el detalle del item elegido
  On Error GoTo ErrorHandler
  Dim strSQL As String
  
  mfgDetalleHistorial.Cols = 0
  mfgDetalleHistorial.Rows = 0
  If IsNumeric(mfgItems.TextMatrix(mfgItems.Row, 0)) Then
    strSQL = "SELECT " & arConsultas(cmbHistoriales.ListIndex).SQLSelect & ", IIF(A.TIPOMOVIMIENTO = 'A','Alta',IIF(A.TIPOMOVIMIENTO = 'M','Modificacion','Baja')) AS [Tipo Movimiento], A.XFECHA AS [Fecha Modificacion], A.XUSUARIO AS [Usuario Modificacion] FROM " & arConsultas(cmbHistoriales.ListIndex).SQLFrom & " WHERE A.XID = " & mfgItems.TextMatrix(mfgItems.Row, 0)
    ProcesarBusqueda2 strSQL, mfgDetalleHistorial, "A.XFECHA"
    MarcarCambios
    ReemplazarRetornosDeLinea
  End If
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub ReemplazarRetornosDeLinea()
'Reemplaza los <VBCRLF> por el verdadero caracter vbCrLf
  On Error GoTo ErrorHandler
  Dim i As Long, j As Long
  Dim AnchoColumna As Double
  
  For i = 0 To mfgDetalleHistorial.Cols - 1
    AnchoColumna = 0
    For j = 1 To mfgDetalleHistorial.Rows - 1
      mfgDetalleHistorial.TextMatrix(j, i) = Replace(mfgDetalleHistorial.TextMatrix(j, i), "<VBCRLF>", vbCrLf)
      If TextWidth(mfgDetalleHistorial.TextMatrix(j, i)) > AnchoColumna Then
       AnchoColumna = TextWidth(mfgDetalleHistorial.TextMatrix(j, i)) + 50
     End If
    Next j
  Next i
   
ErrorHandler:
  ErrorHandler
End Sub

Private Sub MarcarCambios()
'Resalta los cambios entre los registros
  On Error GoTo ErrorHandler
  Dim i As Long
  Dim j As Long
  
  For i = 2 To mfgDetalleHistorial.Rows - 1
    If mfgDetalleHistorial.TextMatrix(i, mfgDetalleHistorial.Cols - 3) <> "Baja" Then
      For j = 1 To mfgDetalleHistorial.Cols - 4
        If mfgDetalleHistorial.TextMatrix(i, j) <> mfgDetalleHistorial.TextMatrix(i - 1, j) Then
          mfgDetalleHistorial.Row = i
          mfgDetalleHistorial.Col = j
          mfgDetalleHistorial.CellBackColor = vbYellow
        End If
      Next j
    End If
  Next i
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub mfgItems_KeyDown(KeyCode As Integer, Shift As Integer)
'Busca una fila para seleccionarla como si hubiera sido el mouse
  On Error GoTo ErrorHandler
    Static Texto As String
    
    BuscarFila mfgItems, KeyCode, Texto
    If KeyCode = vbKeyReturn Then
      mfgItems_DblClick
    End If
    
ErrorHandler:
  ErrorHandler
End Sub

Private Sub mfgDetalleHistorial_KeyDown(KeyCode As Integer, Shift As Integer)
'Busca una fila para seleccionarla como si hubiera sido el mouse
  On Error GoTo ErrorHandler
    Static Texto As String
    
    BuscarFila mfgDetalleHistorial, KeyCode, Texto
       
ErrorHandler:
  ErrorHandler
End Sub

Private Sub txtWellID_KeyDown(KeyCode As Integer, Shift As Integer)
'Ejecuta la busqueda aplicando los filtros seleccionados
  On Error GoTo ErrorHandler
  
  If KeyCode = vbKeyReturn Then
    EjecutarConsulta
  End If
  
ErrorHandler:
  ErrorHandler
End Sub

