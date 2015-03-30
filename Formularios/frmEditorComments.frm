VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEditorComments 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nuevo "
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7200
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
   ScaleHeight     =   3225
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpera 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   5070
      TabIndex        =   11
      Top             =   2790
      Width           =   945
   End
   Begin VB.Frame fraTexto 
      Height          =   3195
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   7065
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   6060
         TabIndex        =   10
         Top             =   2790
         Width           =   945
      End
      Begin VB.TextBox txtNUMEROACTA 
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
         Left            =   750
         MaxLength       =   255
         TabIndex        =   2
         Top             =   990
         Width           =   2985
      End
      Begin VB.TextBox txtComments 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   60
         MaxLength       =   32500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1620
         Width           =   6945
      End
      Begin VB.TextBox txtAutor 
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
         Left            =   750
         MaxLength       =   255
         TabIndex        =   1
         Top             =   660
         Width           =   2985
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   750
         TabIndex        =   0
         Top             =   240
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   18219009
         CurrentDate     =   39804
      End
      Begin VB.Label lblNUMEROACTA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Acta"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   1065
         Width           =   390
      End
      Begin VB.Label lblSiteVisitComments 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   1410
         Width           =   915
      End
      Begin VB.Label lblSiteVisitConductedBy 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Autor"
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   705
         Width           =   480
      End
      Begin VB.Label lblSiteVisitWithDMAConducted 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label etiCampo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Tag             =   "X"
      Top             =   120
      Width           =   3045
   End
End
Attribute VB_Name = "frmEditorComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''Este form permite cambiar varios valores en simultaneo
Option Explicit

Dim m_strOpera As String
Dim m_strDetalle As String
Dim m_gridFrm  As frmPozosFuturos

Public Property Set dsiGridFrm(frm As frmPozosFuturos)
  Set m_gridFrm = frm
End Property

Public Property Get dsiGridFrm() As frmPozosFuturos
  Set dsiGridFrm = m_gridFrm
End Property

Public Property Let dsiOpera(strT As String)
  m_strOpera = strT
End Property

Public Property Get dsiOpera() As String
  dsiOpera = m_strOpera
End Property

Public Property Let dsiDetalle(strT As String)
  m_strDetalle = strT
End Property

Public Property Get dsiDetalle() As String
  dsiDetalle = m_strDetalle
End Property

Private Sub cmdCancelar_Click()
  
  'CLOSE frm
  Unload Me
  
End Sub



Private Sub cmdOpera_Click()
  
  Dim strNumero As String
  Dim varDato As Variant
  
  'TRANSFER valor de grilla a controles
'  Me.dsiGridFrm.spdDet1.GetText Me.dsiGridFrm.spdDet1.GetColFromID("Numero"), Me.dsiGridFrm.spdDet1.ActiveRow, varDato
  
'  'CHECK si debo trabajar con Detalle 1 o 2
'  If Me.dsiDetalle = "LANDOWNER" Then
'    Set spd = Me.dsiGridFrm.spdDet1
'  Else
'    Set spd = Me.dsiGridFrm.spdDet2
'  End If
  
  'CHECK que operación esta realizando
  Select Case Me.dsiOpera
    
    Case "C"  'CREATE
      
      'CHECK si debo trabajar con Detalle 1 o 2
      If Me.dsiDetalle = "LANDOWNER" Then
        
        'CHECK si existe alguna fila, GENERO proximo número, sino, asigno un 1
        If Me.dsiGridFrm.spdDet1.DataRowCnt > 0 Then
          Me.dsiGridFrm.spdDet1.GetText Me.dsiGridFrm.spdDet1.GetColFromID("Numero"), Me.dsiGridFrm.spdDet1.DataRowCnt, varDato
          strNumero = varDato + 1
        Else
          strNumero = 1
        End If
        
        Me.dsiGridFrm.spdDet1.SetText Me.dsiGridFrm.spdDet1.GetColFromID("Numero"), Me.dsiGridFrm.spdDet1.DataRowCnt + 1, strNumero
        Me.dsiGridFrm.spdDet1.SetText Me.dsiGridFrm.spdDet1.GetColFromID("Fecha"), Me.dsiGridFrm.spdDet1.DataRowCnt, Me.dtpFecha.Value
        Me.dsiGridFrm.spdDet1.SetText Me.dsiGridFrm.spdDet1.GetColFromID("Comentario"), Me.dsiGridFrm.spdDet1.DataRowCnt, Me.txtComments
        Me.dsiGridFrm.spdDet1.SetText Me.dsiGridFrm.spdDet1.GetColFromID("Autor"), Me.dsiGridFrm.spdDet1.DataRowCnt, Me.txtAutor
        
      Else
      
        'CHECK si existe alguna fila, GENERO proximo número, sino, asigno un 1
        If Me.dsiGridFrm.spdDet2.DataRowCnt > 0 Then
          Me.dsiGridFrm.spdDet2.GetText Me.dsiGridFrm.spdDet2.GetColFromID("Numero"), Me.dsiGridFrm.spdDet2.DataRowCnt, varDato
          strNumero = varDato + 1
        Else
          strNumero = 1
        End If
      
        Me.dsiGridFrm.spdDet2.SetText Me.dsiGridFrm.spdDet2.GetColFromID("Numero"), Me.dsiGridFrm.spdDet2.DataRowCnt + 1, strNumero
        Me.dsiGridFrm.spdDet2.SetText Me.dsiGridFrm.spdDet2.GetColFromID("Fecha"), Me.dsiGridFrm.spdDet2.DataRowCnt, Me.dtpFecha.Value
        Me.dsiGridFrm.spdDet2.SetText Me.dsiGridFrm.spdDet2.GetColFromID("Comentario"), Me.dsiGridFrm.spdDet2.DataRowCnt, Me.txtComments
        Me.dsiGridFrm.spdDet2.SetText Me.dsiGridFrm.spdDet2.GetColFromID("Autor"), Me.dsiGridFrm.spdDet2.DataRowCnt, Me.txtAutor
        Me.dsiGridFrm.spdDet2.SetText Me.dsiGridFrm.spdDet2.GetColFromID("Acta"), Me.dsiGridFrm.spdDet2.DataRowCnt, Me.txtNUMEROACTA
        
      End If
      
    Case "R"  'READ
      
      'NO HAGO NADA, SOLO CONSULTA
      
    Case "U"  'UPDATE
      
      'CHECK si debo trabajar con Detalle 1 o 2
      If Me.dsiDetalle = "LANDOWNER" Then
        Me.dsiGridFrm.spdDet1.SetText Me.dsiGridFrm.spdDet1.GetColFromID("Fecha"), Me.dsiGridFrm.spdDet1.ActiveRow, Me.dtpFecha.Value
        Me.dsiGridFrm.spdDet1.SetText Me.dsiGridFrm.spdDet1.GetColFromID("Comentario"), Me.dsiGridFrm.spdDet1.ActiveRow, Me.txtComments
        Me.dsiGridFrm.spdDet1.SetText Me.dsiGridFrm.spdDet1.GetColFromID("Autor"), Me.dsiGridFrm.spdDet1.ActiveRow, Me.txtAutor
      Else
        Me.dsiGridFrm.spdDet2.SetText Me.dsiGridFrm.spdDet2.GetColFromID("Fecha"), Me.dsiGridFrm.spdDet2.ActiveRow, Me.dtpFecha.Value
        Me.dsiGridFrm.spdDet2.SetText Me.dsiGridFrm.spdDet2.GetColFromID("Comentario"), Me.dsiGridFrm.spdDet2.ActiveRow, Me.txtComments
        Me.dsiGridFrm.spdDet2.SetText Me.dsiGridFrm.spdDet2.GetColFromID("Autor"), Me.dsiGridFrm.spdDet2.ActiveRow, Me.txtAutor
        Me.dsiGridFrm.spdDet2.SetText Me.dsiGridFrm.spdDet2.GetColFromID("Acta"), Me.dsiGridFrm.spdDet2.ActiveRow, Me.txtNUMEROACTA
      End If
      
    Case "D"  'DELETE
      
      'CHECK si debo trabajar con Detalle 1 o 2
      'Para eliminar, inserto 2 y luego elimino 1, porque no escontre otra forma de hacerlo
      'Esto se da cuando la grilla esta enlazada a un recordset, es como que funciona en automático
      If Me.dsiDetalle = "LANDOWNER" Then
        Me.dsiGridFrm.spdDet1.InsertRows Me.dsiGridFrm.spdDet1.ActiveRow, 1
        Me.dsiGridFrm.spdDet1.DeleteRows Me.dsiGridFrm.spdDet1.ActiveRow, 2
      Else
        Me.dsiGridFrm.spdDet2.InsertRows Me.dsiGridFrm.spdDet2.ActiveRow, 1
        Me.dsiGridFrm.spdDet2.DeleteRows Me.dsiGridFrm.spdDet2.ActiveRow, 2
        
      End If
      
  End Select
  
  'CLOSE frm
  Unload Me
  
End Sub


Private Sub Form_Load()
  
  Dim varNum, varFecha, varComen, varAutor, varActa As Variant
  Dim strTema As String
  
  'SET REFERENCIA a formulario - Pozos futuros
  Set Me.dsiGridFrm = frmMenuPrincipal.ActiveForm
  
  'CHECK con que detalle se va a trabajar
  strTema = IIf(Me.dsiDetalle = "LANDOWNER", "Land Owner Comments", "Site Visit Comments")
  
  'SET titulos del formulario para que no haya confucion
  Select Case Me.dsiOpera
  Case "C"
    Me.Caption = strTema & " Detalle - Nuevo"
    Me.cmdOpera.Caption = "&Nuevo"
  Case "R"
    Me.Caption = strTema & " Detalle - Visualizar"
    Me.cmdOpera.Caption = "&Aceptar"
  Case "U"
    Me.Caption = strTema & " Detalle - Editar"
    Me.cmdOpera.Caption = "&Editar"
  Case "D"
    Me.Caption = strTema & " Detalle - Eliminar"
    Me.cmdOpera.Caption = "&Eliminar"
    
  End Select
  
  
  'CHECK con que detalle se va a trabajar. GET valores
  If Me.dsiDetalle = "LANDOWNER" Then
  
    'GET datos de detalle 1
    Me.dsiGridFrm.spdDet1.GetText Me.dsiGridFrm.spdDet1.GetColFromID("Numero"), Me.dsiGridFrm.spdDet1.ActiveRow, varNum
    Me.dsiGridFrm.spdDet1.GetText Me.dsiGridFrm.spdDet1.GetColFromID("Fecha"), Me.dsiGridFrm.spdDet1.ActiveRow, varFecha
    Me.dsiGridFrm.spdDet1.GetText Me.dsiGridFrm.spdDet1.GetColFromID("Comentario"), Me.dsiGridFrm.spdDet1.ActiveRow, varComen
    Me.dsiGridFrm.spdDet1.GetText Me.dsiGridFrm.spdDet1.GetColFromID("Autor"), Me.dsiGridFrm.spdDet1.ActiveRow, varAutor
  
  Else
  
    'GET datos de detalle 2
    Me.dsiGridFrm.spdDet2.GetText Me.dsiGridFrm.spdDet2.GetColFromID("Numero"), Me.dsiGridFrm.spdDet2.ActiveRow, varNum
    Me.dsiGridFrm.spdDet2.GetText Me.dsiGridFrm.spdDet2.GetColFromID("Fecha"), Me.dsiGridFrm.spdDet2.ActiveRow, varFecha
    Me.dsiGridFrm.spdDet2.GetText Me.dsiGridFrm.spdDet2.GetColFromID("Comentario"), Me.dsiGridFrm.spdDet2.ActiveRow, varComen
    Me.dsiGridFrm.spdDet2.GetText Me.dsiGridFrm.spdDet2.GetColFromID("Autor"), Me.dsiGridFrm.spdDet2.ActiveRow, varAutor
    Me.dsiGridFrm.spdDet2.GetText Me.dsiGridFrm.spdDet2.GetColFromID("Acta"), Me.dsiGridFrm.spdDet2.ActiveRow, varActa
  
  End If
  
  'CHECK si operacion <> Create, TRANSFER valores a controles
  If Me.dsiOpera <> "C" Then
    
    dtpFecha.Value = IIf(varFecha <> "", varFecha, Now())
    txtComments = varComen
    txtAutor = varAutor
    txtNUMEROACTA.Text = varActa
    
  End If
  
  'CHECK si operacion Read o Delete, BLOQUEO controles
  If Me.dsiOpera = "R" Or Me.dsiOpera = "D" Then
    
    Me.dtpFecha.Enabled = False
    Me.txtComments.Locked = True
    Me.txtAutor.Locked = True
    Me.txtNUMEROACTA.Locked = True
    
  End If
  
End Sub

