VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ABM - Usuarios"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
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
      Left            =   3795
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmUsuarios.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Guardar"
      Top             =   1425
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
      Left            =   4065
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmUsuarios.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Reiniciar Formulario"
      Top             =   1425
      UseMaskColor    =   -1  'True
      Width           =   285
   End
   Begin VB.CommandButton cmdEliminar 
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
      Picture         =   "frmUsuarios.frx":0690
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Eliminar Registro"
      Top             =   1425
      UseMaskColor    =   -1  'True
      Width           =   285
   End
   Begin VB.ComboBox cmbTipoUsuario 
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
      ItemData        =   "frmUsuarios.frx":09D2
      Left            =   1200
      List            =   "frmUsuarios.frx":09DC
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.Frame fra 
      Caption         =   "Filtros de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   4575
      Begin VB.ComboBox cmb 
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
         ItemData        =   "frmUsuarios.frx":09F8
         Left            =   720
         List            =   "frmUsuarios.frx":0A05
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txt 
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
         Left            =   720
         TabIndex        =   7
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lbl 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   465
         Width           =   495
      End
      Begin VB.Label lbl 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   825
         Width           =   495
      End
   End
   Begin VB.TextBox txtPassword 
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
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtUsuario 
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
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin MSComctlLib.StatusBar info 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   5580
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   530
            MinWidth        =   530
            Picture         =   "frmUsuarios.frx":0A45
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfg 
      Height          =   2175
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3836
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo "
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
      Left            =   675
      TabIndex        =   13
      Top             =   975
      Width           =   465
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   870
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   825
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''
'ABM De Usuarios'
'''''''''''''''''
Option Explicit

Dim strSQL As String 'Query para buscar los contratistas
Dim arrCampos(1) As String 'Campos de los filtros en el mismo orden que aparecen en el combo
Dim IDModificacion As Long 'ID del registro que se esta modificando
Private Enum enmColumnas 'Columnas de la lista de Usuarios
  ColID
  ColUsuario
  ColTipoUsuario
End Enum


Private Sub Form_Load()
'Setea el query que se utilizara para buscar registros y prepara el formulario para su uso
  On Error GoTo ErrorHandler
  
  Center Me
  strSQL = "SELECT A.IDUSUARIO AS id, A.USUARIO AS Usuario, B.NOMBRE AS [Tipo Usuario] FROM USUARIOS A INNER JOIN TIPOUSUARIOS B ON A.IDTIPOUSUARIO = B.IDTIPOUSUARIO WHERE IDUSUARIO > 1"
  CargarFiltros
  cmdReiniciar_Click
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub CargarFiltros()
'Carga los filtros que se utilizaran para buscar registros
  On Error GoTo ErrorHandler
  
  arrCampos(0) = "USUARIO"
  arrCampos(1) = "NOMBRE"
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub cmb_Click()
'Pasa el foco al otro control, el cont es para evitar que se llame en el form_load
On Error Resume Next
  Static cont As Long
  If cont > 0 Then txt.SetFocus
  cont = cont + 1
End Sub

Public Sub cmdReiniciar_Click()
'Reinicia el formulario

  ReiniciarFormulario Me
  IDModificacion = 0
  ProcesarBusqueda2 strSQL, mfg, "A.USUARIO"
  
End Sub

Private Sub cmdEliminar_Click()
'Elimina el registro especificado mediante el ID de la base de datos
  On Error GoTo ErrorHandler
  
  If IDModificacion <> 1 Then
    EliminarRegistro Me, "USUARIOS", "IDUSUARIO", IDModificacion
  Else
    MsgBox "No se puede eliminar el usuario root", vbInformation + vbOKOnly, "Atencion"
  End If
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub cmdGuardar_Click()
'Realiza las acciones necesarias para guardar
  On Error GoTo ErrorHandler
  
  If VerificarCampos Then
    Guardar
    cmdReiniciar_Click
    txtUsuario.SetFocus
  End If
  
ErrorHandler:
  ErrorHandler
End Sub

Private Function VerificarCampos() As Boolean
'Verifica los datos
  On Error GoTo ErrorHandler
  
  If txtUsuario <> "" Then
    If txtPassword <> "" Then
      VerificarCampos = True
    Else
      MsgBox "El password es requerido", vbInformation + vbOKOnly, "Atención"
      txtPassword.SetFocus
    End If
  Else
    MsgBox "El usuario es requerido", vbInformation + vbOKOnly, "Atención"
    txtUsuario.SetFocus
  End If
  
ErrorHandler:
  ErrorHandler
End Function

Private Sub Guardar()
'Guarda los datos
  On Error GoTo ErrorHandler
  
  Dim RS As New Recordset
  Dim i As Long
  If IDModificacion = 0 Then
    RS.Open "USUARIOS", BD, adOpenDynamic, adLockOptimistic, adCmdTable
    RS.AddNew
      RS!Usuario = txtUsuario
      RS!CONTRASEÑA = txtPassword
      RS!IDTipoUsuario = cmbTipoUsuario.ItemData(cmbTipoUsuario.ListIndex)
    RS.Update
    RS.Close
    GuardarHistorial "USUARIOS", "IDUSUARIO", IDModificacion, "A"
  Else
     BD.Execute "UPDATE USUARIOS SET USUARIO = '" & txtUsuario & "', CONTRASEÑA = '" & txtPassword & "', IDTIPOUSUARIO = " & cmbTipoUsuario.ItemData(cmbTipoUsuario.ListIndex) & " WHERE IDUSUARIO = " & IDModificacion
     GuardarHistorial "USUARIOS", "IDUSUARIO", IDModificacion, "M"
  End If
    
ErrorHandler:
  ErrorHandler
End Sub


Private Sub mfg_DblClick()
'Carga los datos para modificar el registro en los campos superiores
  On Error GoTo ErrorHandler
  Dim i As Long

  IDModificacion = mfg.TextMatrix(mfg.Row, ColID)
  txtUsuario = mfg.TextMatrix(mfg.Row, ColUsuario)
  txtPassword = ObtenerValorCampo("USUARIOS", "CONTRASEÑA", "IDUSUARIO = " & IDModificacion)
  cmbTipoUsuario = mfg.TextMatrix(mfg.Row, ColTipoUsuario)
   
ErrorHandler:
  ErrorHandler
End Sub


Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
'Ejecuta la busqueda aplicando los filtros seleccionados
  On Error GoTo ErrorHandler
  
  If KeyCode = vbKeyReturn Then
    ProcesarBusqueda strSQL, mfg, "A.USUARIO", cmb, arrCampos(cmb.ListIndex), txt
  End If
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub mfg_KeyDown(KeyCode As Integer, Shift As Integer)
'Busca una fila para seleccionarla como si hubiera sido el mouse
  On Error GoTo ErrorHandler
    Static Texto As String
    
    BuscarFila mfg, KeyCode, Texto
    If KeyCode = vbKeyReturn Then
      mfg_DblClick
    End If
    
ErrorHandler:
  ErrorHandler
End Sub



