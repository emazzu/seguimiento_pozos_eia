VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAdministradorBD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrador de Base de Datos"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdministradorBD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMaxFila 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5700
      MaxLength       =   100
      TabIndex        =   4
      Top             =   2805
      Width           =   975
   End
   Begin MSComDlg.CommonDialog comDiag 
      Left            =   5640
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtContraseña 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   100
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton cmdArchivo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7305
      Picture         =   "frmAdministradorBD.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1530
      Width           =   330
   End
   Begin VB.TextBox txtArchivo 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   5895
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton cmdActualizar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6825
      Picture         =   "frmAdministradorBD.frx":054E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Copiar Datos"
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   810
   End
   Begin MSComctlLib.StatusBar Barra 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3120
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "frmAdministradorBD.frx":0CE6
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8821
            MinWidth        =   8821
            Key             =   "info"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   5106
            MinWidth        =   5115
            Text            =   "www.ConsultoresGIS.com"
            TextSave        =   "www.ConsultoresGIS.com"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMaxFila 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima fila"
      Height          =   270
      Left            =   4500
      TabIndex        =   13
      Top             =   2812
      Width           =   960
   End
   Begin VB.Label lblNota 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nota: Esta modificacion es irreversible"
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Tag             =   "X"
      Top             =   2812
      Width           =   3570
   End
   Begin VB.Label lblArchivo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo"
      Height          =   270
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   270
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label lblContraseña 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1035
   End
   Begin VB.Label lblIntroduccion2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Si desea continuar, seleccione un archivo Excel de seguimiento de EIAs y escriba su usuario y contraseña. Luego pulse el boton"
      Height          =   570
      Left            =   120
      TabIndex        =   7
      Tag             =   "X"
      Top             =   720
      Width           =   7530
   End
   Begin VB.Label lblIntroduccion1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Atencion: Esta modificacion eliminara todos los pozos existentes en la base de datos."
      Height          =   570
      Left            =   120
      TabIndex        =   6
      Tag             =   "X"
      Top             =   120
      Width           =   7290
   End
End
Attribute VB_Name = "frmAdministradorBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RutaExcel As String



Private Function VerificarUsuario() As Boolean
'Verifica que la combinacion usuario contraseña sea valida y carga la infoglobal
  On Error GoTo ErrorHandler
  
  Dim RS As New Recordset
  Dim strSQL As String
  
  strSQL = "SELECT * FROM USUARIOS WHERE USUARIO = '" & txtUsuario & "' AND CONTRASEÑA = '" & txtContraseña & "' AND IDTIPOUSUARIO = " & enmTipoUsuarioAdministrador
  RS.Open strSQL, BD, adOpenDynamic, adLockOptimistic, adCmdText
  If Not RS.EOF Then
    If MsgBox("Se reemplazarán todos los pozos existentes en la base de datos, esta modificación NO se podrá deshacer. ¿Confirma que desea continuar?", vbQuestion + vbYesNo, "Advertencia: Modificacion irreversible") = vbYes Then
      VerificarUsuario = True
    End If
  Else
    MsgBox "Combinación usuario/contraseña incorrectas", vbExclamation + vbOKOnly, "Error"
  End If
  
ErrorHandler:
  ErrorHandler
End Function


Private Sub cmdArchivo_Click()
'Permite buscar archivos excel para importar pozos
  On Error GoTo ErrorHandler
  
  Dim FSO As New FileSystemObject
  
  comDiag.Filter = "xls"
  comDiag.CancelError = True
  comDiag.DefaultExt = "xls"
  comDiag.DialogTitle = "Seleccione un archivo de pozos"
  comDiag.ShowOpen
  If UCase(FSO.GetExtensionName(comDiag.FileName)) = "XLS" Then
    RutaExcel = comDiag.FileName
    txtArchivo = comDiag.FileTitle
  Else
    RutaExcel = ""
    txtArchivo = ""
  End If
  
ErrorHandler:
  If Err.Number <> 0 And Err.Number <> 32755 Then
    ErrorHandler
  End If
End Sub

Private Sub Form_Load()
'Prepara el form para usar
  On Error GoTo ErrorHandler
  
  Center Me
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub cmdActualizar_Click()
'Verifica y llama a la rutina de actualizar
  On Error GoTo ErrorHandler
  
  BD.BeginTrans

  If RutaExcel <> "" And txtArchivo <> "" Then
    If IsNumeric(txtMaxFila) Then
      If CLng(txtMaxFila) > 12 Then
        If VerificarUsuario Then
          Me.Enabled = False
          frmMenuPrincipal.Enabled = False
          ReiniciarBD
          CargarDatosExcel
          Me.Enabled = True
          frmMenuPrincipal.Enabled = True
        End If
      Else
        MsgBox "El Numero de fila maximo debe ser mayor que 12 (Fila de inicio)", vbOKOnly + vbCritical, "Atencion"
      End If
    Else
      MsgBox "El Numero de fila maximo debe ser numerico", vbOKOnly + vbCritical, "Atencion"
    End If
   Else
    MsgBox "Debe seleccionar un archivo de datos", vbOKOnly + vbCritical, "Atencion"
  End If

  
ErrorHandler:
  If Err.Number <> 0 Then
    BD.RollbackTrans
    ErrorHandler
  Else
    BD.CommitTrans
  End If
  cmdActualizar.Enabled = True
End Sub


Private Sub ReiniciarBD()
'Reinicia la BD eliminando todas las tablas necesarias
  On Error GoTo ErrorHandler
  
  Barra.Panels("info") = "Eliminando datos anteriores"
  BD.Execute "DELETE FROM IMAGEINTERPRETATIONCOMMENTS"
  BD.Execute "DELETE FROM HIST_IMAGEINTERPRETATIONCOMMENTS"
  BD.Execute "DELETE FROM SITEVISITCOMMENTS"
  BD.Execute "DELETE FROM HIST_SITEVISITCOMMENTS"
  BD.Execute "DELETE FROM POZOS"
  BD.Execute "DELETE FROM HIST_POZOS"
  
ErrorHandler:
  If Err.Number <> 0 Then Err.Raise Err.Number
End Sub


Private Sub CargarDatosExcel()
'Graba en el access los datos del Excel
  On Error GoTo ErrorHandler
  
  Dim Excel As New Excel.Application
  Dim Libro As Excel.Workbook
  Set Libro = Excel.Workbooks.Open(RutaExcel)
  Dim Hoja As Excel.Worksheet
  Set Hoja = Libro.Sheets(1)
  Dim RS As New Recordset
  Dim RS2 As New Recordset
  Dim Fila As Long
  Dim Equipo As String
  Dim TipoYacimiento As String
  Dim OrdenUbicacion As Long
  Dim UbicacionAnterior As String
  
  OrdenUbicacion = 1
  For Fila = 12 To txtMaxFila
    Barra.Panels("info") = "Importando Fila " & Fila - 11 & " de " & txtMaxFila - 11
    RS.Open "POZOS", BD, adOpenDynamic, adLockOptimistic, adCmdTable
    If Trim(Hoja.Cells(Fila, "K")) <> "" Then  'Tiene WellID debe ser un pozo
      RS.AddNew
        RS!WellID = Hoja.Cells(Fila, "K")
        RS!Ubicacion = IIf(Trim(Hoja.Cells(Fila, "I")) = "", Null, Hoja.Cells(Fila, "I"))
        If UbicacionAnterior <> RS!Ubicacion Then
          UbicacionAnterior = RS!Ubicacion
          OrdenUbicacion = OrdenUbicacion + 1
        End If
        RS!OrdenUbicacion = OrdenUbicacion
        RS!Equipo = IIf(Hoja.Cells(Fila, "I") = "RIGS SCHED.", Equipo, "")
        RS!MONOGRAFIA = IIf(UCase(Trim(Hoja.Cells(Fila, "U"))) = "SI", True, False)
        RS!ACUIFERO = IIf(UCase(Trim(Hoja.Cells(Fila, "W"))) = "SI", True, False)
        RS!Prognosis = IIf(UCase(Trim(Hoja.Cells(Fila, "X"))) = "SI", True, False)
        RS!Programa = IIf(UCase(Trim(Hoja.Cells(Fila, "Y"))) = "SI", True, False)
        RS!Yacimiento = IIf(Trim(Hoja.Cells(Fila, "BM")) = "", Null, Hoja.Cells(Fila, "BM"))
        
        If Trim(Hoja.Cells(Fila, "BM")) <> Trim(Hoja.Cells(Fila, "BL")) Then
          TipoYacimiento = Replace(Trim(Hoja.Cells(Fila, "BL")), Trim(Hoja.Cells(Fila, "BM")), "")
          If UCase(TipoYacimiento) = "A" Then
            RS!TipoYacimiento = "a"
          ElseIf UCase(TipoYacimiento) = "X" Then
            RS!TipoYacimiento = "e"
          Else
            RS!TipoYacimiento = Null
          End If
        Else
          RS!TipoYacimiento = Null
        End If
        RS!POZO = IIf(Trim(Hoja.Cells(Fila, "BN")) = "", Null, Hoja.Cells(Fila, "BN"))
        RS!Prospect = IIf(Trim(Hoja.Cells(Fila, "BO")) = "", Null, Hoja.Cells(Fila, "BO"))
        
'        RS!FieldManifold = IIf(Trim(Hoja.Cells(Fila, "BO")) = "", Null, Hoja.Cells(Fila, "BO"))
'        RS!BatteryAssigned = IIf(Trim(Hoja.Cells(Fila, "BO")) = "", Null, Hoja.Cells(Fila, "BO"))
'

        RS!FECHAPRIMERMONOGRAFIA = IIf(Not IsDate(Hoja.Cells(Fila, "AN")), Null, Hoja.Cells(Fila, "AN"))
        RS!MONOGRAFIAS = IIf(Not IsNumeric(Hoja.Cells(Fila, "AO")), Null, Hoja.Cells(Fila, "AO"))
        RS!FECHAULTIMAMONOGRAFIA = IIf(Not IsDate(Hoja.Cells(Fila, "AP")), Null, Hoja.Cells(Fila, "AP"))
        RS!WellInformed = IIf(Not IsDate(Hoja.Cells(Fila, "BY")), Null, Hoja.Cells(Fila, "BY"))
        RS!INFORMEDBY = IIf(Trim(Hoja.Cells(Fila, "BZ")) = "", Null, Hoja.Cells(Fila, "BZ"))
        
        RS!Definitiva = IIf(UCase(Trim(Hoja.Cells(Fila, "BQ"))) = "SI", True, False)
        RS!XPDC = IIf(Not IsNumeric(Hoja.Cells(Fila, "BR")), Null, Hoja.Cells(Fila, "BR"))
        RS!YPDC = IIf(Not IsNumeric(Hoja.Cells(Fila, "BS")), Null, Hoja.Cells(Fila, "BS"))
        RS!XPOS94 = IIf(Not IsNumeric(Hoja.Cells(Fila, "BV")), Null, Hoja.Cells(Fila, "BV"))
        RS!YPOS94 = IIf(Not IsNumeric(Hoja.Cells(Fila, "BW")), Null, Hoja.Cells(Fila, "BW"))
        RS!FECHASOLICITUD = IIf(Not IsDate(Hoja.Cells(Fila, "AA")), Null, Hoja.Cells(Fila, "AA"))
        RS!FECHAPRIORIDAD = IIf(Not IsDate(Hoja.Cells(Fila, "AB")), Null, Hoja.Cells(Fila, "AB"))
        RS!DOCUMENTOAPREPARAR = IIf(Trim(Hoja.Cells(Fila, "AC")) = "", Null, Hoja.Cells(Fila, "AC"))
        
        'vergpOK
        'RS!FECHAENTREGAPLANES8PTS = IIf(Not IsDate(Hoja.Cells(Fila, "AD")), Null, Hoja.Cells(Fila, "AD"))
        'RS!FECHAENTREGAPLANES13PTS = IIf(Not IsDate(Hoja.Cells(Fila, "AE")), Null, Hoja.Cells(Fila, "AE
        
        RS!FieldManifold = IIf(Trim(Hoja.Cells(Fila, "AD")) = "", Null, Hoja.Cells(Fila, "AD"))
        RS!BatteryAssigned = IIf(Trim(Hoja.Cells(Fila, "AE")) = "", Null, Hoja.Cells(Fila, "AE"))

        
        RS!FECHAENTREGADICTAMENTECNICO = IIf(Not IsDate(Hoja.Cells(Fila, "AF")), Null, Hoja.Cells(Fila, "AF"))
        'RS!FIRSTPROD = IIf(Not IsDate(Hoja.Cells(Fila, "AG")), Null, Hoja.Cells(Fila, "AG"))
        RS!TD = IIf(Not IsNumeric(Hoja.Cells(Fila, "L")), Null, Hoja.Cells(Fila, "L"))
        RS!TOTDAYS = IIf(Not IsNumeric(Hoja.Cells(Fila, "AU")), Null, Hoja.Cells(Fila, "AU"))
        RS!REMDAYS = IIf(Not IsNumeric(Hoja.Cells(Fila, "AV")), Null, Hoja.Cells(Fila, "AV"))
        RS!Status = IIf(Trim(Hoja.Cells(Fila, "AW")) = "", Null, Hoja.Cells(Fila, "AW"))
        RS!STARTDATE = IIf(Not IsDate(Hoja.Cells(Fila, "AX")), Null, Hoja.Cells(Fila, "AX"))
        RS!ENDDATE = IIf(Not IsDate(Hoja.Cells(Fila, "AY")), Null, Hoja.Cells(Fila, "AY"))
        RS!LANDOWNER = IIf(Trim(Hoja.Cells(Fila, "AZ")) = "", Null, Hoja.Cells(Fila, "AZ"))
        RS!LANDOWNERPERMITDATE = IIf(Not IsDate(Hoja.Cells(Fila, "BA")), Null, Hoja.Cells(Fila, "BA"))
        RS!Consult = IIf(Trim(Hoja.Cells(Fila, "BD")) = "", Null, Hoja.Cells(Fila, "BD"))
        RS!Type = IIf(Trim(Hoja.Cells(Fila, "BE")) = "", Null, Hoja.Cells(Fila, "BE"))
        RS!FECHAPEDIDOETIA = Null
        RS!FECHAESPERADAETIA = IIf(Not IsDate(Hoja.Cells(Fila, "M")), Null, Hoja.Cells(Fila, "M"))
        RS!CONSULTANTRECOMENDATION = IIf(Trim(Hoja.Cells(Fila, "AJ")) = "", Null, Hoja.Cells(Fila, "AJ"))
        RS!DMAPERMIT = IIf(Not IsDate(Hoja.Cells(Fila, "AS")), Null, Hoja.Cells(Fila, "AS"))
        RS!Estado = IIf(Trim(Hoja.Cells(Fila, "AT")) = "", Null, Hoja.Cells(Fila, "AT"))
        RS!IDMANIFIESTO = IIf(Trim(Hoja.Cells(Fila, "CA")) = "", Null, Hoja.Cells(Fila, "CA"))
        RS!FECHAMANIFIESTO = IIf(Not IsDate(Hoja.Cells(Fila, "CB")), Null, Hoja.Cells(Fila, "CB"))
        RS!FECHAENTREGAEIAXCONSULTORAOXY = IIf(Not IsDate(Hoja.Cells(Fila, "CC")), Null, Hoja.Cells(Fila, "CC"))
        RS!EIAPRESENTADO = IIf(UCase(Trim(Hoja.Cells(Fila, "CD"))) = "SI", True, False)
        RS!FECHAENVIOACS = IIf(Not IsDate(Hoja.Cells(Fila, "CE")), Null, Hoja.Cells(Fila, "CE"))
        RS!FECHAPRESENTACIONDMA = IIf(Not IsDate(Hoja.Cells(Fila, "CF")), Null, Hoja.Cells(Fila, "CF"))
        RS!FECHAPRESENTACIONSMA = IIf(Not IsDate(Hoja.Cells(Fila, "CG")), Null, Hoja.Cells(Fila, "CG"))
        RS!PAGOTASAADMINISTRATIVA = IIf(Not IsDate(Hoja.Cells(Fila, "CH")), Null, Hoja.Cells(Fila, "CH"))
        RS!FECHAINFOCOMPLEMENTARIA = IIf(Not IsDate(Hoja.Cells(Fila, "CI")), Null, Hoja.Cells(Fila, "CI"))
        RS!TECHNICALREPORT = IIf(Trim(Hoja.Cells(Fila, "CJ")) = "", Null, Hoja.Cells(Fila, "CJ"))
        RS!TASAADMINISTRATIVA = IIf(Not IsNumeric(Hoja.Cells(Fila, "CK")), Null, Hoja.Cells(Fila, "CK"))
        RS!TASACONTRALOR = IIf(Not IsNumeric(Hoja.Cells(Fila, "CL")), Null, Hoja.Cells(Fila, "CL"))
        RS!ESTUDIO = IIf(Not IsNumeric(Hoja.Cells(Fila, "CM")), Null, Hoja.Cells(Fila, "CM"))
        RS!ADENDA = IIf(Trim(Hoja.Cells(Fila, "CN")) = "", Null, Hoja.Cells(Fila, "CN"))
'        RS!AFE = IIf(Not IsNumeric(Hoja.Cells(Fila, "CO")), Null, Hoja.Cells(Fila, "CO"))
'        RS!AFECERRADO = IIf(UCase(CStr(Hoja.Cells(Fila, "CP"))) = "SI", True, False)
'
        RS!FECHAINICIODIA = IIf(Not IsDate(Hoja.Cells(Fila, "CO")), Null, Hoja.Cells(Fila, "CO"))
        RS!FECHAFINDIA = IIf(Not IsDate(Hoja.Cells(Fila, "CP")), Null, Hoja.Cells(Fila, "CP"))
        
'    MsgBox "8"
        
        RS!TIEMPOENTREPEDIDOETIAYRECEPCIONETIA = Null
        RS!TIEMPOENTREPOZOINFORMADOYRECEPCIONETIA = IIf(Not IsNumeric(Hoja.Cells(Fila, "CQ")), Null, Hoja.Cells(Fila, "CQ"))
        RS!TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA = IIf(Not IsNumeric(Hoja.Cells(Fila, "CR")), Null, Hoja.Cells(Fila, "CR"))
        RS!TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA = IIf(Not IsNumeric(Hoja.Cells(Fila, "CS")), Null, Hoja.Cells(Fila, "CS"))
        RS!TIEMPOENTREPRESENTACIONANTEDMAYVISITA = IIf(Not IsNumeric(Hoja.Cells(Fila, "CT")), Null, Hoja.Cells(Fila, "CT"))
        RS!TIEMPOENTREVISITAYAPROBACIONFINALDEDMA = IIf(Not IsNumeric(Hoja.Cells(Fila, "CU")), Null, Hoja.Cells(Fila, "CU"))
        RS!TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA = IIf(Not IsNumeric(Hoja.Cells(Fila, "CV")), Null, Hoja.Cells(Fila, "CV"))
        RS!INFORMEDEAVANCEDEOBRA50PORCIENTO = IIf(Trim(Hoja.Cells(Fila, "CW")) = "", Null, Hoja.Cells(Fila, "CW"))
        RS!INFORMEDEAVANCEDEOBRA100PORCIENTO = IIf(Trim(Hoja.Cells(Fila, "CX")) = "", Null, Hoja.Cells(Fila, "CX"))
        RS!INFORMEEVALUACIONARQUEOLOGICO = IIf(Trim(Hoja.Cells(Fila, "CY")) = "", Null, Hoja.Cells(Fila, "CY"))
      RS.Update
      GuardarHistorial "POZOS", "IDPOZO", 0, "A"
      RS2.Open "IMAGEINTERPRETATIONCOMMENTS", BD, adOpenDynamic, adLockOptimistic, adCmdTable
      RS2.AddNew
        RS2!IDPozo = RS!IDPozo
        RS2!NumeroVisita = 1
        RS2!NumeroACTA = Null
        RS2!FECHA = Null
        RS2!Comments = IIf(Trim(Hoja.Cells(Fila, "AI")) = "", Null, Hoja.Cells(Fila, "AI"))
        RS2!AUTOR = Null
      RS2.Update
      GuardarHistorial "IMAGEINTERPRETATIONCOMMENTS", "IDIMAGEINTERPRETATIONCOMMENT", RS2!IDIMAGEINTERPRETATIONCOMMENT, "A"
      RS2.Close
      RS2.Open "SITEVISITCOMMENTS", BD, adOpenDynamic, adLockOptimistic, adCmdTable
      RS2.AddNew
        RS2!IDPozo = RS!IDPozo
        RS2!NumeroVisita = 1
        RS2!NumeroACTA = Null
        RS2!FECHA = IIf(Not IsDate(Hoja.Cells(Fila, "AK")), Null, Hoja.Cells(Fila, "AK"))
        RS2!Comments = IIf(Trim(Hoja.Cells(Fila, "AM")) = "", Null, Hoja.Cells(Fila, "AM"))
        RS2!AUTOR = IIf(Trim(Hoja.Cells(Fila, "AL")) = "", Null, Hoja.Cells(Fila, "AL"))
      RS2.Update
      GuardarHistorial "SITEVISITCOMMENTS", "IDSITEVISITCOMMENT", RS2!IDSITEVISITCOMMENT, "A"
      RS2.Close
    Else
      Equipo = IIf(Trim(Hoja.Cells(Fila, "I")) = "", Null, Hoja.Cells(Fila, "I"))
    End If
    RS.Close
    DoEvents
  Next Fila
  Barra.Panels("info") = "Proceso finalizado"
  
  Libro.Saved = True
  Libro.Close
  Set Hoja = Nothing
  Set Libro = Nothing
  Set Excel = Nothing

  
ErrorHandler:
  If Err.Number <> 0 Then Err.Raise Err.Number
End Sub


