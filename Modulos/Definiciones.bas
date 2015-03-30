Attribute VB_Name = "Definiciones"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public BD As New ADODB.Connection 'Objeto de la BD
Public RutaBD As String 'Ruta de la BD Access

'Informacion sobre la sesion del usuario
Public InfoGlobal As TGlobal
Public Type TGlobal
  IDUsuario As Long
  Usuario As String
  IDTipoUsuario As Long
  TipoUsuario As String
End Type

'Tipos de usuario en el sistema
Public Enum enmTipoUsuario
  enmTipoUsuarioAdministrador = 1
  enmTipoUsuarioComun = 2
End Enum

'Variables para la Búsqueda
Public TipoBusqueda As String
Public FormBusqueda As Form
Public UltimoPozoSeleccionado As String


