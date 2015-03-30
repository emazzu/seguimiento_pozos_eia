VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrdenadorCategorias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenamiento de Ubicaciones"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
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
   Icon            =   "frmOrdenadorCategorias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMemo 
      Height          =   3015
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   4815
      Begin MSComctlLib.ListView lsvCategorias 
         Height          =   2655
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   4683
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
            Text            =   "Categorias"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.CommandButton cmdSubir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmOrdenadorCategorias.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.CommandButton cmdBajar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmOrdenadorCategorias.frx":0254
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   300
      End
   End
   Begin VB.Frame fraUbicacion 
      Height          =   1095
      Left            =   120
      TabIndex        =   24
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
         ItemData        =   "frmOrdenadorCategorias.frx":0624
         Left            =   120
         List            =   "frmOrdenadorCategorias.frx":0626
         Style           =   2  'Dropdown List
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.Frame fraBool 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CheckBox chkBool 
         Caption         =   "No"
         Height          =   255
         Left            =   120
         TabIndex        =   12
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
         TabIndex        =   13
         Top             =   240
         Width           =   1965
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
      Left            =   4620
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmOrdenadorCategorias.frx":0628
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3435
      UseMaskColor    =   -1  'True
      Width           =   285
   End
   Begin VB.Frame fraTexto 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   240
         Width           =   3120
      End
   End
   Begin VB.Frame fraFecha 
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   59375617
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
         TabIndex        =   7
         Top             =   240
         Width           =   2130
      End
   End
   Begin VB.Frame fraEntero 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   3420
      End
   End
   Begin MSComctlLib.StatusBar info 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3795
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
            Picture         =   "frmOrdenadorCategorias.frx":096A
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
   Begin VB.Frame fraTipoYacimiento 
      Height          =   1455
      Left            =   120
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.OptionButton optDesarrollo 
         Caption         =   "Desarrollo"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optAvanzada 
         Caption         =   "De avanzada"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1065
         Width           =   1455
      End
      Begin VB.OptionButton optExploratorio 
         Caption         =   "Exploratorio"
         Height          =   255
         Left            =   240
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Frame fraDecimal 
      Height          =   1095
      Left            =   120
      TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   17
         Top             =   240
         Width           =   3510
      End
   End
   Begin VB.Label lblCampo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el orden en que se mostraran las categorias"
      Height          =   195
      Left            =   165
      TabIndex        =   1
      Tag             =   "X"
      Top             =   120
      Width           =   4665
   End
End
Attribute VB_Name = "frmOrdenadorCategorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''Este form permite cambiar el orden en el que el ABM de pozos muestra las categorias
Option Explicit



Private Sub cmdGuardar_Click()
'Modifica las ubicaciones
  On Error GoTo ErrorHandler
  
  Dim i As Long
  If lsvCategorias.ListItems.Count > 0 Then
    For i = 1 To lsvCategorias.ListItems.Count
      BD.Execute "UPDATE POZOS SET ORDENUBICACION = " & i & " WHERE UBICACION = '" & lsvCategorias.ListItems(i) & "'"
    Next i
    MsgBox "Guardado con exito", vbInformation + vbOKOnly, "Guardado"
  Else
    MsgBox "No hay ubicaciones cargadas", vbInformation + vbOKOnly, "Atencion"
  End If
ErrorHandler:
  ErrorHandler
End Sub

Private Sub cmdSubir_Click()
'Sube una linea
  On Error GoTo ErrorHandler
  
  If Not lsvCategorias.SelectedItem Is Nothing Then
    If Not lsvCategorias.SelectedItem.Index = 1 Then
      Set lsvCategorias.SelectedItem = lsvCategorias.ListItems.Add(lsvCategorias.SelectedItem.Index - 1, , lsvCategorias.SelectedItem.Text)
      lsvCategorias.ListItems.Remove lsvCategorias.SelectedItem.Index + 2
    End If
  End If
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub cmdBajar_Click()
'Baja una linea
  On Error GoTo ErrorHandler
  
  If Not lsvCategorias.SelectedItem Is Nothing Then
    If Not lsvCategorias.SelectedItem.Index = lsvCategorias.ListItems.Count Then
      Set lsvCategorias.SelectedItem = lsvCategorias.ListItems.Add(lsvCategorias.SelectedItem.Index + 2, , lsvCategorias.SelectedItem.Text)
      lsvCategorias.ListItems.Remove lsvCategorias.SelectedItem.Index - 2
    End If
  End If
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub Form_Load()
'Prepara el form para su uso
  On Error GoTo ErrorHandler
  
  Center Me
  CargarListaCategorias
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub CargarListaCategorias()
'Carga las ubicaciones
 On Error GoTo ErrorHandler
 
  Dim RS As New Recordset
  Dim i As Integer
 
  RS.Open "SELECT UBICACION FROM (SELECT DISTINCT ORDENUBICACION, UBICACION FROM POZOS WHERE NOT UBICACION IS NULL) ORDER BY ORDENUBICACION", BD, adOpenDynamic, adLockOptimistic, adCmdText
  While Not RS.EOF
    lsvCategorias.ListItems.Add , , RS(0)
    RS.MoveNext
  Wend
  RS.Close
  
ErrorHandler:
  ErrorHandler
End Sub

