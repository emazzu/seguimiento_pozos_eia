VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCodigosStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Codigos Status"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCodigosStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMemo 
      Height          =   7695
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5535
      Begin MSComctlLib.ListView lsvStatus 
         Height          =   7335
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   12938
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Status"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
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
      Left            =   5355
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCodigosStatus.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
      UseMaskColor    =   -1  'True
      Width           =   285
   End
   Begin MSComctlLib.StatusBar info 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8505
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   530
            MinWidth        =   530
            Picture         =   "frmCodigosStatus.frx":034E
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
   Begin VB.Label lblCampo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Asigne a Cada posible valor de Status un codigo"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Tag             =   "X"
      Top             =   120
      Width           =   4020
   End
End
Attribute VB_Name = "frmCodigosStatus"
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
  If lsvStatus.ListItems.Count > 0 Then
    For i = 1 To lsvStatus.ListItems.Count
      BD.Execute "UPDATE CODIGOSTATUS SET STATUS = " & IIf(Trim(lsvStatus.ListItems(i)) = "", Null, "'" & Trim(lsvStatus.ListItems(i)) & "'") & ", CODIGO = " & IIf(Trim(lsvStatus.ListItems(i).SubItems(1)) = "", Null, "'" & Trim(lsvStatus.ListItems(i).SubItems(1)) & "'") & " WHERE STATUS = '" & lsvStatus.ListItems(i) & "'"
    Next i
    MsgBox "Guardado con exito", vbInformation + vbOKOnly, "Guardado"
  Else
    MsgBox "No hay status cargados", vbInformation + vbOKOnly, "Atencion"
  End If
ErrorHandler:
  ErrorHandler
End Sub


Private Sub Form_Load()
'Prepara el form para su uso
  On Error GoTo ErrorHandler
  
  Center Me
  CargarListaStatus
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub CargarListaStatus()
'Carga las ubicaciones
 On Error GoTo ErrorHandler
 
  Dim RS As New Recordset
  Dim i As Integer
  Dim NLI As ListItem
 
  RS.Open "SELECT STATUS, CODIGO FROM CODIGOSTATUS", BD, adOpenDynamic, adLockOptimistic, adCmdText
  While Not RS.EOF
    Set NLI = lsvStatus.ListItems.Add(, , RS(0))
    NLI.SubItems(1) = RS(1) & ""
    RS.MoveNext
  Wend
  RS.Close
  
ErrorHandler:
  ErrorHandler
End Sub

Private Sub lsvStatus_DblClick()
  Dim Codigo As String
  Codigo = InputBox("Ingrese el nuevo codigo para la categoria " & lsvStatus.SelectedItem, "Cambio de codigo", lsvStatus.SelectedItem.SubItems(1))
  lsvStatus.SelectedItem.SubItems(1) = Codigo
End Sub
