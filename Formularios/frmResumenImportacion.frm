VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResumenImportacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vintage Data: Resumen de Importacion "
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5850
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
   ScaleHeight     =   6030
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGenerarPozos 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5430
      Picture         =   "frmResumenImportacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Generar pozos seleccionados"
      Top             =   5310
      Width           =   330
   End
   Begin MSComctlLib.ListView lsvEncontrados 
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.StatusBar info 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5655
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   530
            MinWidth        =   530
            Picture         =   "frmResumenImportacion.frx":0312
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
   Begin MSComctlLib.ListView lsvNoEncontrados 
      Height          =   4935
      Left            =   3000
      TabIndex        =   4
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pozos NO encontrados"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblEncontrados 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pozos encontrados"
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
      Top             =   120
      Width           =   1725
   End
End
Attribute VB_Name = "frmResumenImportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private arEncontrados() As String
Private arNoEncontrados() As String


'   13/01/2012
'   NO SE USA MAS, lee directo de DSinfo
'
'
'Private Sub cmdGenerarPozos_Click()
''Genera los pozos seleccionados con los campos obtenidos del vintage data
'  On Error GoTo ErrorHandler
'  Dim IDs As String
'  Dim i As Long
'
'  IDs = ObtenerIDsSeleccionados
'  If IDs <> "" Then
'    If MsgBox("Se generaran los pozos seleccionados con los datos obtenidos de Vintage Data. ¿Confirma que desea continuar?", vbQuestion + vbYesNo, "Atención") = vbYes Then
'      ConectarBDVintageData
'      CrearPozosNoEncontrados IDs
'      DesconectarBDVintageData
'
'      For i = 1 To lsvNoEncontrados.ListItems.Count
'        If i > lsvNoEncontrados.ListItems.Count Then
'          Exit For
'        End If
'        If lsvNoEncontrados.ListItems(i).Checked Then
'          lsvNoEncontrados.ListItems.Remove i
'          i = 0
'        End If
'      Next i
'    End If
'  Else
'    MsgBox "Debe seleccionar los pozos que desea generar", vbOKOnly + vbInformation, "Atención"
'  End If
'
'ErrorHandler:
'  ErrorHandler
'End Sub



Private Function ObtenerIDsSeleccionados() As String
  Dim i As Long
  For i = 1 To lsvNoEncontrados.ListItems.Count
    If lsvNoEncontrados.ListItems(i).Checked Then
      ObtenerIDsSeleccionados = ObtenerIDsSeleccionados & ",'" & lsvNoEncontrados.ListItems(i) & "'"
    End If
  Next i
  ObtenerIDsSeleccionados = Mid(ObtenerIDsSeleccionados, 2)
End Function

Private Sub Form_Load()
'Prepara el form para usar
  On Error GoTo ErrorHandler
  
  Center Me
  OrdenarAlfabeticamente arEncontrados
  OrdenarAlfabeticamente arNoEncontrados
  AsignarArreglo lsvEncontrados, arEncontrados
  AsignarArreglo lsvNoEncontrados, arNoEncontrados
  
ErrorHandler:
  ErrorHandler
End Sub



Private Sub OrdenarAlfabeticamente(Ar() As String)
'Ordena alfabeticamente las listas
  On Error GoTo ErrorHandler
  
  Dim i As Long
  Dim j As Long
  Dim aux As String
  
  For i = 0 To UBound(Ar) - 1
    For j = i + 1 To UBound(Ar)
      If Ar(i) > Ar(j) Then
        aux = Ar(i)
        Ar(i) = Ar(j)
        Ar(j) = aux
      End If
    Next j
  Next i
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub AsignarArreglo(lsv As ListView, Ar() As String)
'Ordena alfabeticamente las listas
  On Error GoTo ErrorHandler
  Dim i As Long
  
  For i = 0 To UBound(Ar)
    lsv.ListItems.Add , , Ar(i)
  Next i
  
ErrorHandler:
  ErrorHandler
End Sub


Public Sub AsignarEncontrados(Texto As String)
  arEncontrados = Split(Texto, vbCrLf)
End Sub


Public Sub AsignarnoEncontrados(Texto As String)
  arNoEncontrados = Split(Texto, vbCrLf)
End Sub

