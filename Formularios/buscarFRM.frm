VERSION 5.00
Begin VB.Form buscarFRM 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5265
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCadena 
      Caption         =   "Texto Completo"
      Height          =   300
      Left            =   90
      TabIndex        =   4
      Top             =   810
      Width           =   1860
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   3690
      TabIndex        =   3
      Top             =   810
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Default         =   -1  'True
      Height          =   315
      Left            =   2070
      TabIndex        =   2
      Top             =   810
      Width           =   1455
   End
   Begin VB.TextBox txtDato 
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese información a buscar"
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5025
   End
End
Attribute VB_Name = "buscarFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngUltimoEncontrado As Long
Dim m_gridFrm  As frmPozosFuturos

Public Property Set dsiGridFrm(frm As frmPozosFuturos)
  Set m_gridFrm = frm
End Property


Public Property Get dsiGridFrm() As frmPozosFuturos
  Set dsiGridFrm = m_gridFrm
End Property


Private Sub cmdBuscar_Click()
  
  Dim intRow As Long
  Dim intCadenaTot As Integer
  Dim blnB As Boolean
   
  If Me.txtDato = "" Then
    Me.txtDato.SetFocus
    Beep
    Exit Sub
  End If
   
  ' cambio puntero mouse
  Screen.MousePointer = vbHourglass
  buscarFRM.Caption = "Buscando..."
  
  If chkCadena.Value = 0 Then
    intCadenaTot = SearchFlagsPartialMatch
  Else
    intCadenaTot = SearchFlagsNone
  End If
  
  'recorro grilla
  intRow = lngUltimoEncontrado + 1
  Do While intRow <= Me.dsiGridFrm.spdCab.MaxRows
  
    'buscando en fila y todas las columnas
    intRes = Me.dsiGridFrm.spdCab.SearchRow(intRow, 0, -1, Me.txtDato, intCadenaTot)
    
    'si encontro
    If intRes <> -1 Then
                
      'SET  color default de grilla a celda para abandonar
      '     Actualmente no pinta celda
      Me.dsiGridFrm.spdCab.BlockMode = True
      Me.dsiGridFrm.spdCab.Row = Me.dsiGridFrm.spdCab.ActiveRow
      Me.dsiGridFrm.spdCab.Row2 = Me.dsiGridFrm.spdCab.ActiveRow
      Me.dsiGridFrm.spdCab.Col = Me.dsiGridFrm.spdCab.ActiveCol
      Me.dsiGridFrm.spdCab.Col2 = -1
'      Me.dsiGridFrm.spdCab.BackColor = RGB(245, 245, 245)
'      Me.dsiGridFrm.spdCab.ForeColor = RGB(60, 60, 60)
      
      'set activa celda
      Me.dsiGridFrm.spdCab.SetActiveCell intRes, intRow
      
      
      'set  color de seleccion a celda nueva
      '     Actualmente no pinta celda
      Me.dsiGridFrm.spdCab.BlockMode = True
      Me.dsiGridFrm.spdCab.Row = Me.dsiGridFrm.spdCab.ActiveRow
      Me.dsiGridFrm.spdCab.Row2 = Me.dsiGridFrm.spdCab.ActiveRow
      Me.dsiGridFrm.spdCab.Col2 = 1
      Me.dsiGridFrm.spdCab.Col2 = -1
'      Me.dsiGridFrm.spdCab.BackColor = RGB(220, 220, 220)
'      Me.dsiGridFrm.spdCab.ForeColor = RGB(60, 60, 60)
      Me.dsiGridFrm.spdCab.BlockMode = False
  
  
      'SET puntero interno en fila columna
      Me.dsiGridFrm.spdCab.Row = Me.dsiGridFrm.spdCab.ActiveRow
      Me.dsiGridFrm.spdCab.Col = Me.dsiGridFrm.spdCab.ActiveCol
      
'      'muestra en barra de estado de frm MDI
'      MainMDI.staBarra1.Panels(2) = " fila: " & Me.spdGrid.ActiveRow & " de " & Me.spdGrid.MaxRows & " , columna: " & Me.spdGrid.ActiveCol & " de " & Me.spdGrid.MaxCols & " , valor: " & Me.spdGrid.Text & " "
'      MainMDI.staBarra1.Panels(2).ToolTipText = " fila: " & Me.spdGrid.ActiveRow & " de " & Me.spdGrid.MaxRows & " , columna: " & Me.spdGrid.ActiveCol & " de " & Me.spdGrid.MaxCols & " , valor: " & Me.spdGrid.Text & " "
      
      
      lngUltimoEncontrado = intRow
      cmdBuscar.Caption = "&Buscar Siguiente"
      Exit Do
      
    End If
      
    intRow = intRow + 1
      
  Loop
  
  ' recupero puntero mouse
  Screen.MousePointer = vbDefault
  buscarFRM.Caption = "Buscar"
  
  'si no encontro nada
  If intRow >= Me.dsiGridFrm.spdCab.MaxRows Then
    intRes = MsgBox("Finalizó la búsqueda.", vbInformation + vbOKOnly, "atención...")
    cmdBuscar.Caption = "&Buscar"
    lngUltimoEncontrado = 0
  End If
  
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  
  'referencio grilla en donde buscar
  Set Me.dsiGridFrm = frmMenuPrincipal.ActiveForm
  
End Sub
