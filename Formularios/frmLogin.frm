VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4830
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
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContraseña 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   100
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      MaxLength       =   100
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin MSComctlLib.StatusBar info 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1035
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   530
            MinWidth        =   530
            Picture         =   "frmLogin.frx":0000
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
   Begin VB.Image imiLlave 
      Height          =   600
      Left            =   120
      Picture         =   "frmLogin.frx":05F2
      Top             =   240
      Width           =   600
   End
   Begin VB.Label lblContraseña 
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblUsuario 
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
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtContraseña_GotFocus()
'Cuando la contraseña recibe el foco selecciona todo el texto para reemplazarlo
  On Error GoTo ErrorHandler
  
  SeleccionTexto txtContraseña
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub txtContraseña_KeyPress(KeyAscii As Integer)
'Si se preciona enter sobre la contraseña se realiza la validacion de usuario
  On Error GoTo ErrorHandler
  
  If KeyAscii = vbKeyReturn Then
    If VerificarCampos Then
      VerificarUsuario
    End If
  End If
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub txtUsuario_GotFocus()
'Cuando el usuario recibe el foco selecciona todo el texto para reemplazarlo
  On Error GoTo ErrorHandler
  
  SeleccionTexto txtUsuario
  
ErrorHandler:
  ErrorHandler
End Sub

Public Sub SeleccionTexto(txt As TextBox)
'Selecciona el texto del control activo
  On Error GoTo ErrorHandler

  txt.SelStart = 0
  txt.SelLength = Len(txt.Text)
    
ErrorHandler:
  ErrorHandler
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
'Si preciona enter en el usuario se pasa el foco a contraseña
  On Error GoTo ErrorHandler

  If KeyAscii = vbKeyReturn Then
    SendKeys "{tab}"
  End If
  
ErrorHandler:
  ErrorHandler
End Sub

Private Function VerificarCampos() As Boolean
'Acciones del menu Archivo
  On Error GoTo ErrorHandler
  
  If txtUsuario <> "" Then
    If txtContraseña <> "" Then
      VerificarCampos = True
    Else
      MsgBox "La contraseña es requerida", vbInformation + vbOKOnly, "Atención"
      txtContraseña.SetFocus
    End If
  Else
    MsgBox "El usuario es requerido", vbInformation + vbOKOnly, "Atención"
    txtUsuario.SetFocus
  End If
  
ErrorHandler:
  ErrorHandler
End Function

Private Sub VerificarUsuario()
'Verifica que la combinacion usuario contraseña sea valida y carga la infoglobal
  On Error GoTo ErrorHandler
  
  Dim RS As New Recordset
  Dim strSQL As String
  
  strSQL = "SELECT A.IDUSUARIO, A.USUARIO, A.IDTIPOUSUARIO, B.NOMBRE FROM USUARIOS A LEFT JOIN TIPOUSUARIOS B ON A.IDTIPOUSUARIO = B.IDTIPOUSUARIO WHERE A.USUARIO = '" & txtUsuario & "' AND A.CONTRASEÑA = '" & txtContraseña & "'"
  RS.Open strSQL, BD, adOpenDynamic, adLockOptimistic, adCmdText
  If Not RS.EOF Then
    InfoGlobal.Usuario = RS!Usuario & ""
    InfoGlobal.IDUsuario = IIf(IsNull(RS!IDUsuario), 0, RS!IDUsuario)
    InfoGlobal.IDTipoUsuario = RS!IDTipoUsuario & ""
    InfoGlobal.TipoUsuario = RS!Nombre & ""
    frmMenuPrincipal.Show
    Unload Me
  Else
    MsgBox "Combinación usuario/contraseña incorrectas", vbExclamation + vbOKOnly, "Error"
  End If
  
ErrorHandler:
  ErrorHandler
End Sub
