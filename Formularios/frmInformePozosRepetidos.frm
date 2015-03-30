VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInformePozosRepetidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe - Pozos Repetidos"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInformePozosRepetidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   17280
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
      Left            =   780
      TabIndex        =   3
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
      Left            =   5595
      Picture         =   "frmInformePozosRepetidos.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   97
      Width           =   330
   End
   Begin MSComctlLib.StatusBar info 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8715
      Width           =   17280
      _ExtentX        =   30480
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   530
            MinWidth        =   530
            Picture         =   "frmInformePozosRepetidos.frx":054E
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgPozosRepetidos 
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   6376
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgDetallePozos 
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   6376
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Detalles del pozo seleccionado"
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
      TabIndex        =   5
      Top             =   4680
      Width           =   2805
   End
   Begin VB.Label lblWellID 
      Caption         =   "Well ID"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   615
   End
   Begin VB.Label lblItems 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Los siguientes pozos contienen Well ID Repetidos"
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
      Top             =   600
      Width           =   4515
   End
End
Attribute VB_Name = "frmInformePozosRepetidos"
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
  
  CargarPozosRepetidos
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub txtWellID_KeyDown(KeyCode As Integer, Shift As Integer)
'Ejecuta la busqueda aplicando los filtros seleccionados
  On Error GoTo ErrorHandler

  If KeyCode = vbKeyReturn Then
    CargarPozosRepetidos
  End If

ErrorHandler:
  ErrorHandler
End Sub


Private Sub Form_Load()
'Prepara el formulario
  On Error GoTo ErrorHandler
  
  Center Me
  CargarPozosRepetidos
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub CargarPozosRepetidos()
'Carga los items de la seleccion del combo
  On Error GoTo ErrorHandler
  Dim strSQL As String
  
  mfgPozosRepetidos.Cols = 0
  mfgPozosRepetidos.Rows = 0
  mfgDetallePozos.Cols = 0
  mfgDetallePozos.Rows = 0
  strSQL = "SELECT '' AS ID, A.WELLID AS [Well ID], COUNT(*) AS Cantidad FROM POZOS A " & IIf(Trim(txtWellID) <> "", "WHERE A.WELLID LIKE '%" & Trim(txtWellID) & "%'", "") & " GROUP BY 1, A.WELLID HAVING COUNT(*) > 1"
  ProcesarBusqueda2 strSQL, mfgPozosRepetidos, "A.WELLID"
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub mfgPozosRepetidos_Click()
'Carga los items de la seleccion del combo
  On Error GoTo ErrorHandler
  Dim strSQL As String
  Const colWellID As Long = 1
  
  mfgDetallePozos.Cols = 0
  mfgDetallePozos.Rows = 0
  
  'vergpTEST
'  MsgBox "12"
  
  strSQL = "SELECT A.IDPOZO AS [Id], A.WELLID AS [Well ID], A.Ubicacion, A.Equipo, IIF(A.MONOGRAFIA,'Si','No') AS [Monografía], IIF(A.ACUIFERO,'Si','No') AS [Acuífero], IIF(A.Prognosis,'Si','No') AS [Plan de 8], IIF(A.Programa,'Si','No') AS [Plan de 13], A.Yacimiento, A.TIPOYACIMIENTO AS [Tipo Yacimiento], A.Pozo, A.Prospect, A.FECHAPRIMERMONOGRAFIA AS [Fecha Primera Monografia], A.MONOGRAFIAS AS [Monografías], A.FECHAULTIMAMONOGRAFIA AS [Fecha Ultima Monografía], A.WELLINFORMED AS [Well Informed], A.INFORMEDBY AS [Informed By], IIF(A.Definitiva,'Si','No') AS Definitiva, A.XPDC AS [X PDC], A.YPDC AS [Y PDC], A.XPOS94 AS [X POS94], A.YPOS94 AS [Y POS94], " & _
           " A.FECHASOLICITUD AS [Fecha Solicitud], A.FECHAPRIORIDAD as [Fecha Prioridad], A.DOCUMENTOAPREPARAR AS [Documento a Preparar], A.FieldManifold, A.BatteryAssigned, A.FECHAENTREGADICTAMENTECNICO AS [Fecha Entrega Dictamen Tco], A.FIRSTPROD AS [First Prod], " & _
           " A.TD, A.TOTDAYS AS [Tot Days], A.REMDAYS AS [Rem Days], A.Status, A.STARTDATE AS [Start Date], A.ENDDATE AS [End Date], A.LANDOWNER AS [Land Owner], A.LANDOWNERPERMITDATE AS [Land Owner Permit Date], " & _
           " A.Consult, A.Type, A.FechaPedidoEtia AS [Fecha Pedido ETIA], A.FechaEsperadaEtia AS [Fecha Esperada ETIA], " & _
           " (SELECT TOP 1 B.NUMEROVISITA & ' | ' & B.FECHA & ' | ' & Left(B.COMMENTS,255) & ' | ' & B.AUTOR FROM IMAGEINTERPRETATIONCOMMENTS B WHERE A.IDPOZO = B.IDPOZO ORDER BY B.NUMEROVISITA DESC) AS [Ultimo Image Interpretation Comments], A.CONSULTANTRECOMENDATION AS [Consultant Recomendation], (SELECT TOP 1 C.NUMEROVISITA & ' | ' & C.FECHA & ' | ' & C.NUMEROACTA & ' | ' & Left(C.COMMENTS,255) & ' | ' & C.AUTOR FROM SITEVISITCOMMENTS C WHERE A.IDPOZO = C.IDPOZO ORDER BY C.NUMEROVISITA DESC) AS [Ultimo Site Visit with DMA], A.DMAPERMIT AS [DMA Permit], A.Estado, " & _
           " A.IDMANIFIESTO AS [ID Manifiesto], A.FECHAMANIFIESTO AS [Fecha Manifiesto], A.FECHAENTREGAEIAXCONSULTORAOXY AS [Fecha Entrega EIA Por Consultor a OXY], IIF(A.EIAPRESENTADO,'Si','No') AS [EIA Presentado], A.FECHAENVIOACS AS [Fecha Envio ACS], A.FECHAPRESENTACIONDMA AS [Fecha Presentación DMA], A.FECHAPRESENTACIONSMA AS [Fecha Presentación SMA], A.PAGOTASAADMINISTRATIVA AS [Pago Tasa Administrativa], A.FECHAINFOCOMPLEMENTARIA AS [Fecha Info Complementaria], A.TECHNICALREPORT AS [Technical Report], A.TASAADMINISTRATIVA AS [Tasa Administrativa], A.TASACONTRALOR AS [Tasa Contralor], A.Estudio, A.Adenda, A.FECHAINICIODIA AS [Fecha Inicio DIA], A.FECHAFINDIA AS [Fecha Fin DIA], A.INFORMEDEAVANCEDEOBRA50PORCIENTO AS [Informe de Avance de Obra 50%], A.INFORMEDEAVANCEDEOBRA100PORCIENTO AS [Informe de Avance de Obra 100%], A.INFORMEEVALUACIONARQUEOLOGICO AS [Informe Evaluacion Arqueologico], " & _
           " A.TIEMPOENTREPEDIDOETIAYRECEPCIONETIA AS [Tiempo Entre Pedido ETIA y Recepcion ETIA], A.TIEMPOENTREPOZOINFORMADOYRECEPCIONETIA AS [Tiempo Entre Pozo Informado y Recepcion ETIA], A.TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA AS [Tiempo Entre Primer Monografía y Recepción ETIA], A.TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA AS [Tiempo Entre Recepcion ETIA y Presentación ante DMA], A.TIEMPOENTREPRESENTACIONANTEDMAYVISITA AS [Tiempo Entre Presentación Ante DMA y Visita], A.TIEMPOENTREVISITAYAPROBACIONFINALDEDMA AS [Tiempo Entre Visita y Aprobación Final De DMA], A.TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA AS [Tiempo Entre Presentación De Pozo y Aprobación Final DMA] FROM POZOS A"

  ProcesarBusqueda2 strSQL, mfgDetallePozos, "A.UBICACION", "Well ID (Texto)", "A.WELLID", mfgPozosRepetidos.TextMatrix(mfgPozosRepetidos.Row, colWellID)

ErrorHandler:
  ErrorHandler
End Sub

