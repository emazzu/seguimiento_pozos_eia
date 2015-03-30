VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmInformeGraficos 
   Caption         =   "Graficos"
   ClientHeight    =   12585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInformeGraficos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12585
   ScaleWidth      =   18465
   Begin VB.CheckBox chkLegenda 
      Caption         =   "Mostrar legenda"
      Height          =   195
      Left            =   14400
      TabIndex        =   6
      Top             =   180
      Width           =   2295
   End
   Begin VB.ComboBox cmbTipoGrafico 
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
      ItemData        =   "frmInformeGraficos.frx":038A
      Left            =   9840
      List            =   "frmInformeGraficos.frx":039A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   3375
   End
   Begin MSChart20Lib.MSChart Grafico 
      Height          =   11535
      Left            =   120
      OleObjectBlob   =   "frmInformeGraficos.frx":03C7
      TabIndex        =   3
      Top             =   600
      Width           =   18255
   End
   Begin VB.ComboBox cmbConsulta 
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
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin MSComctlLib.StatusBar info 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   12210
      Width           =   18465
      _ExtentX        =   32570
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Picture         =   "frmInformeGraficos.frx":285C
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   27519
            MinWidth        =   27519
            Key             =   "info"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   3969
            MinWidth        =   2
            Text            =   "www.ConsultoresGIS.com"
            TextSave        =   "www.ConsultoresGIS.com"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTipoGrafico 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar como"
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
      Left            =   8535
      TabIndex        =   5
      Top             =   150
      Width           =   1245
   End
   Begin VB.Label lblConsulta 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione una consulta"
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
      Width           =   2220
   End
End
Attribute VB_Name = "frmInformeGraficos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''
'Consultas de historiales'
''''''''''''''''''''''''''
Private Type TConsulta 'Tipo de dato que maneja la info de las consultas
  Nombre As String
  EjeX As String
  EjeY As String
  Sql As String
End Type
Private arConsultas(2) As TConsulta

Private Sub chkLegenda_Click()
  Grafico.ShowLegend = (chkLegenda = vbChecked)
End Sub

Private Sub cmbTipoGrafico_Click()
  Grafico.ChartType = cmbTipoGrafico.ItemData(cmbTipoGrafico.ListIndex)
End Sub

Private Sub Form_Load()
'Prepara el formulario
  On Error GoTo ErrorHandler
  
  Me.Width = 18500
  Me.Height = 13000
  Me.Top = 200
  Me.Left = 200
  CargarConsultas
  CargarComboConsultas
  cmbTipoGrafico.ListIndex = 1
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub CargarConsultas()
'Carga las consultas
  On Error GoTo ErrorHandler
  
  arConsultas(0).Nombre = "Well presented Vs Days to get the DMA Permit 2008 - 2009"
  arConsultas(0).EjeX = "Periodo"
  arConsultas(0).EjeY = "Days to get a Permit"
  arConsultas(0).Sql = "SELECT FECHAPRIMERMONOGRAFIA, TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA As TIEMPO From POZOS Where Not FECHAPRIMERMONOGRAFIA Is Null And FECHAPRIMERMONOGRAFIA >= #04/01/2008# AND NOT TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA IS NULL AND TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA >= 0 AND UBICACION IN ('GONE / DONE','RIGS SCHED.') ORDER BY 1 ASC"
  
  arConsultas(1).Nombre = "EIA Cost Vs Time"
  arConsultas(1).EjeX = "Periodo x Fecha primer Monografia"
  arConsultas(1).EjeY = "Cost (ARS)"
  arConsultas(1).Sql = "SELECT FECHAPRIMERMONOGRAFIA, ESTUDIO As MONTO From POZOS Where Not FECHAPRIMERMONOGRAFIA Is Null And FECHAPRIMERMONOGRAFIA >= #04/01/2008# AND NOT ESTUDIO IS NULL AND ESTUDIO >= 0 AND UBICACION IN ('GONE / DONE','RIGS SCHED.') ORDER BY 1 ASC"
  
  arConsultas(2).Nombre = "EIA Cost Vs Time"
  arConsultas(2).EjeX = "Periodo x Fecha Manifiesto"
  arConsultas(2).EjeY = "Cost (ARS)"
  arConsultas(2).Sql = "SELECT FECHAMANIFIESTO, SUM(ESTUDIO) AS MONTO From POZOS Where Not FECHAMANIFIESTO Is Null And FECHAMANIFIESTO >= #04/01/2008# AND NOT ESTUDIO IS NULL AND ESTUDIO >= 0 AND UBICACION IN ('GONE / DONE','RIGS SCHED.') GROUP BY FECHAMANIFIESTO ORDER BY 1 ASC"
    
ErrorHandler:
  ErrorHandler
End Sub

Private Sub CargarComboConsultas()
'Carga los nombres de las consultas en el combo
  On Error GoTo ErrorHandler
  
  Dim i As Long
  For i = LBound(arConsultas) To UBound(arConsultas)
    cmbConsulta.AddItem arConsultas(i).Nombre
  Next i
  If cmbConsulta.ListCount > 0 Then cmbConsulta.ListIndex = 0
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub cmbConsulta_Click()
'Carga los items de la seleccion del combo
  On Error GoTo ErrorHandler
  
  Select Case Index
    Case 0, 1, 2: EjecutarConsulta
  End Select
  
ErrorHandler:
  ErrorHandler
End Sub



Private Sub EjecutarConsulta()
'Carga los items de la seleccion del combo
  On Error GoTo ErrorHandler
  Dim i As Long
  Dim RS As New Recordset
  Dim MesAnterior As Long
  Dim AñoAnterior As Long
  Const colFecha = 0
  Const colTiempo = 0
  
  i = cmbConsulta.ListIndex
  RS.CursorLocation = adUseClient
  RS.Open arConsultas(i).Sql, BD, adOpenStatic, adLockReadOnly, adCmdText
  
  Grafico.ColumnCount = 0
  Grafico.RowCount = 0
  Grafico.Title.Text = arConsultas(i).Nombre
  Grafico.Plot.Axis(VtChAxisIdX).AxisTitle.Text = arConsultas(i).EjeX
  Grafico.Plot.Axis(VtChAxisIdY).AxisTitle.Text = arConsultas(i).EjeY
  While Not RS.EOF
      If Month(RS.Fields(colFecha)) <> MesAnterior Or Year(RS.Fields(colFecha)) <> AñoAnterior Then
        If Grafico.RowCount > 0 Then
          Grafico.RowLabel = Format("01/" & MesAnterior & "/" & AñoAnterior, "mm/yyyy")
        End If
        MesAnterior = Month(RS.Fields(colFecha))
        AñoAnterior = Year(RS.Fields(colFecha))
        Grafico.ColumnCount = 1
        Grafico.Column = 1
      End If
      Grafico.RowCount = Grafico.RowCount + 1
      Grafico.Row = Grafico.RowCount
      Grafico.RowLabel = ""
      Grafico.Data = RS.Fields(1)
    RS.MoveNext
  Wend
  Grafico.Refresh

ErrorHandler:
  ErrorHandler
End Sub

Private Sub Form_Resize()
  Grafico.Width = Me.Width - 100
  Grafico.Height = Me.Height - Grafico.Top - 850
End Sub

Private Sub Grafico_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
  Grafico.Row = DataPoint
  Grafico.ToolTipText = Grafico.Data
  info.Panels("info") = "Valor del punto seleccionado: " & Grafico.Data
End Sub
