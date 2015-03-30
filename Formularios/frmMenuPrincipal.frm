VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMenuPrincipal 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Seguimiento EIAs de Pozos (version 2014.10.03)"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13875
   Icon            =   "frmMenuPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar info 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6810
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   688
            MinWidth        =   530
            Picture         =   "frmMenuPrincipal.frx":2CFA
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Text            =   "Usuario:"
            TextSave        =   "Usuario:"
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Visible         =   0   'False
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Visible         =   0   'False
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   44106
            MinWidth        =   44097
            Text            =   "www.ConsultoresGIS.com      "
            TextSave        =   "www.ConsultoresGIS.com      "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu menPrincipal 
      Caption         =   "Datos"
      Index           =   1
      Begin VB.Menu menABM 
         Caption         =   "Pozos"
         Index           =   0
      End
      Begin VB.Menu menABM 
         Caption         =   "Ordenar Categorias"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu menABM 
         Caption         =   "Codigos Status"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu menPrincipal 
      Caption         =   "Informes"
      Index           =   2
      Begin VB.Menu menInformes 
         Caption         =   "Ver Historial"
         Index           =   0
      End
      Begin VB.Menu menInformes 
         Caption         =   "Pozos Repetidos"
         Index           =   1
      End
      Begin VB.Menu menInformes 
         Caption         =   "Graficos"
         Index           =   2
      End
   End
   Begin VB.Menu menPrincipal 
      Caption         =   "Seguridad"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu menSeguridad 
         Caption         =   "Usuarios"
         Index           =   0
      End
      Begin VB.Menu menSeguridad 
         Caption         =   "Administrador de Base de Datos"
         Index           =   1
      End
   End
   Begin VB.Menu menPrincipal 
      Caption         =   ""
      Index           =   4
      Visible         =   0   'False
      Begin VB.Menu menGrilla 
         Caption         =   "Editar xxx"
         Index           =   0
      End
   End
   Begin VB.Menu menPrincipal 
      Caption         =   ""
      Index           =   5
      Visible         =   0   'False
      Begin VB.Menu menExportar 
         Caption         =   "Exportar listado completo"
         Index           =   0
      End
      Begin VB.Menu menExportar 
         Caption         =   "Exportar estado de lineas"
         Index           =   1
      End
      Begin VB.Menu menExportar 
         Caption         =   "Save to TTM Database"
         Index           =   2
      End
      Begin VB.Menu menExportar 
         Caption         =   "Save permitting data"
         Index           =   3
      End
   End
   Begin VB.Menu menPrincipal 
      Caption         =   ""
      Index           =   6
      Visible         =   0   'False
      Begin VB.Menu menImageInterpretationComments 
         Caption         =   "Nuevo"
         Index           =   0
      End
      Begin VB.Menu menImageInterpretationComments 
         Caption         =   "Editar"
         Index           =   1
      End
      Begin VB.Menu menImageInterpretationComments 
         Caption         =   "Eliminar"
         Index           =   2
      End
   End
   Begin VB.Menu menPrincipal 
      Caption         =   ""
      Index           =   7
      Visible         =   0   'False
      Begin VB.Menu menSiteVisit 
         Caption         =   "Nuevo"
         Index           =   0
      End
      Begin VB.Menu menSiteVisit 
         Caption         =   "Editar"
         Index           =   1
      End
      Begin VB.Menu menSiteVisit 
         Caption         =   "Eliminar"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Enumeracion para el menu principal
Private Enum enmMenuPrincipal
  mPrincipal_Archivo
  mPrincipal_ABM
  mPrincipal_Informes
  mPrincipal_Seguridad
End Enum

'Enumeracion para el menu Archivo
Private Enum enmMenuArchivo
  mArchivo_Salir
  mArchivo_CerrarSesion
End Enum
'Enumeracion para el menu ABM
Private Enum enmMenuABM
  mABM_PozosFuturos
  mABM_OrdenarCategorias
  mABM_CodigosStatus
End Enum
'Enumeracion para el menu Informes
Private Enum enmMenuInformes
  mInformes_Historiales
  mInformes_PozosRepetidos
  mInformes_Graficos
End Enum
'Enumeracion para el menu Seguridad
Private Enum enmMenuSeguridad
  mSeguridad_Usuarios
  mSeguridad_AdministradorBD
End Enum

Private Sub info_PanelClick(ByVal Panel As MSComctlLib.Panel)
'Abre la web de consultoresGIS
  On Error GoTo ErrorHandler
  
  If Panel.Index = info.Panels.Count Then
    ShellExecute Me.hWnd, "open", Trim(Panel.Text), &O0, &O0, vbNormalFocus
  End If
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub MDIForm_Load()

  Dim rs As ADODB.Recordset
  Dim strT, strL As String
  Dim intI As Integer
  Dim blnB As Boolean
  
  'check configuracion regional
  If Not checkConfigRegional() Then
    blnB = MsgBox("El sistema detecto que la configuración regional no es correcta." & vbCrLf & vbCrLf & _
           "Configurar el formato para números de esta forma: 123,456,789.00." & vbCrLf & vbCrLf & _
           "Configurar el formato para fechas  de esta forma: dd/MM/yyyy.", vbCritical + vbOKOnly, "Atención...")
    End
  End If
  
  'get parametros de conexion
  blnB = SQLgetParam()
  
  'check si parametros ok, get menu
  If Not blnB Then
    
    blnB = MsgBox("Los parametros no son correctos.", vbCritical + vbOKOnly, "Atención...")
    
    'Cierro
    SQLclose
    
    'Exit
    Unload Me
           
  End If
  
  'cierro
  SQLclose
  
  'SHOW usuario que se acaba de conectar
  Me.info.Panels(2) = " " & SQLparam.UsuarioConectado & " "
  
'  'FILL tolltip
'  For intI = 1 To UBound(SQLparam.GrupoConectado)
'    Me.info.Panels(2).ToolTipText = Me.info.Panels(2).ToolTipText & SQLparam.GrupoConectado(intI) & " - "
'  Next
  
  
'   30/12/2011
'   No se usa mas, utilizo toda la conectividad de la VDS
'
''Prepara el MDI para usar
'  On Error GoTo ErrorHandler
'
'  info.Panels("Usuario") = "Usuario: " & InfoGlobal.Usuario
'  info.Panels("TipoUsuario") = "Tipo de Usuario: " & InfoGlobal.TipoUsuario
'  frmMenuPrincipal.WindowState = vbMaximized
'  VerificarPermisos
'
'
'ErrorHandler:
'  ErrorHandler
End Sub

Private Sub VerificarPermisos()
'Habilita las opciones de menu segun el tipo de usuario
  On Error GoTo ErrorHandler
  
  menPrincipal(mPrincipal_ABM).Enabled = False
  menABM(mABM_PozosFuturos).Enabled = False
  menABM(mABM_OrdenarCategorias).Enabled = False
  menPrincipal(mPrincipal_Informes).Enabled = False
  menInformes(mInformes_Historiales).Enabled = False
  menInformes(mInformes_PozosRepetidos).Enabled = False
  menInformes(mInformes_Graficos).Enabled = False
  menPrincipal(mPrincipal_Seguridad).Enabled = False
  menSeguridad(mSeguridad_Usuarios).Enabled = False
  menSeguridad(mSeguridad_AdministradorBD).Enabled = False
  menABM(mABM_CodigosStatus).Enabled = False
  
  Select Case InfoGlobal.IDTipoUsuario
    Case enmTipoUsuarioAdministrador
      menPrincipal(mPrincipal_ABM).Enabled = True
      menABM(mABM_PozosFuturos).Enabled = True
      menPrincipal(mPrincipal_Informes).Enabled = True
      menInformes(mInformes_Historiales).Enabled = True
      menInformes(mInformes_PozosRepetidos).Enabled = True
      menInformes(mInformes_Graficos).Enabled = True
      menPrincipal(mPrincipal_Seguridad).Enabled = True
      menSeguridad(mSeguridad_Usuarios).Enabled = True
      menSeguridad(mSeguridad_AdministradorBD).Enabled = True
      menABM(mABM_OrdenarCategorias).Enabled = True
      menABM(mABM_CodigosStatus).Enabled = True
    Case enmTipoUsuarioComun
      menPrincipal(mPrincipal_ABM).Enabled = True
      menABM(mABM_PozosFuturos).Enabled = True
      menPrincipal(mPrincipal_Informes).Enabled = True
      menInformes(mInformes_Historiales).Enabled = True
  End Select
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub menABM_Click(Index As Integer)
'Atiende el click de los subMenus de ABM
  On Error GoTo ErrorHandler

  Select Case Index
    Case mABM_PozosFuturos
      frmPozosFuturos.Show
    Case mABM_OrdenarCategorias
      frmOrdenadorCategorias.Show
    Case mABM_CodigosStatus
      frmCodigosStatus.Show
  End Select
  
ErrorHandler:
  ErrorHandler
End Sub


'   13/01/2012
'   NO SE USA MAS, lee directo de DSinfo
'
'
'Private Sub menExportar_Click(Index As Integer)
''Atiende los eventos del menu exportar (Aunque este menu se encuentre en este form corresponde al form de pozos)
'  On Error GoTo ErrorHandler
'
'  Exportar Index
'
'ErrorHandler:
'  ErrorHandler
'End Sub



'   13/01/2012
'   NO SE USA MAS, lee directo de DSinfo
'
'Private Sub menGrilla_Click(Index As Integer)
''Atiende el click del menu de la grilla de formpozos (aunque este menu se encuentre en este form corresponde al form de pozos)
'  On Error GoTo ErrorHandler
'
'  frmPozosFuturos.MenuClick
'
'ErrorHandler:
'  ErrorHandler
'End Sub



'   13/01/2012
'   NO SE USA MAS, lee directo de DSinfo
'
'Private Sub menImageInterpretationComments_Click(Index As Integer)
''Atiende el click del menu de la grilla de formpozos (aunque este menu se encuentre en este form corresponde al form de pozos)
'  On Error GoTo ErrorHandler
'
'  frmPozosFuturos.ImageInterpretationCommentsMenuClick Index
'
'ErrorHandler:
'  ErrorHandler
'End Sub


'   13/01/2012
'   NO SE USA MAS, lee directo de DSinfo
'
'
'Private Sub menSiteVisit_Click(Index As Integer)
''Atiende el click del menu de la grilla de formpozos (aunque este menu se encuentre en este form corresponde al form de pozos)
'  On Error GoTo ErrorHandler
'
'  frmPozosFuturos.SiteVisitCommentsMenuClick Index
'
'ErrorHandler:
'  ErrorHandler
'End Sub



Private Sub menInformes_Click(Index As Integer)
'Atiende los clicks de los submenus de informes
  On Error GoTo ErrorHandler
  
  Select Case Index
    Case mInformes_Historiales
      frmInformeHistoriales.Show
    Case mInformes_PozosRepetidos
      frmInformePozosRepetidos.Show
    Case mInformes_Graficos
      frmInformeGraficos.Show
  End Select
  
  
ErrorHandler:
  ErrorHandler
End Sub


'30/12/2011
'NO SE USA MAS
'
'Private Sub Salir()
''Solicita confirmacion y cierra el sistema
'  On Error GoTo ErrorHandler
'
'  If MsgBox("Cualquier cambio no guardado se perderá. ¿Confirma que desea salir de la aplicación?", vbQuestion + vbYesNo, "Confirmar salida") = vbYes Then
'    End
'  End If
'ErrorHandler:
'  ErrorHandler
'End Sub


'30/12/2011
'NO SE USA MAS
'
'Private Sub menSeguridad_Click(Index As Integer)
''Atiende los clicks de los submenus de seguridad
'  On Error GoTo ErrorHandler
'
'  Select Case Index
'    Case mSeguridad_Usuarios
'      frmUsuarios.Show
'    Case mSeguridad_AdministradorBD
'      frmAdministradorBD.Show
'  End Select
'
'ErrorHandler:
'  ErrorHandler
'End Sub

Private Sub mnu_salir_Click()
      
      'EXIT
      Unload Me
      
End Sub
