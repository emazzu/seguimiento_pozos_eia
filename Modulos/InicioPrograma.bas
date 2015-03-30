Attribute VB_Name = "InicioPrograma"
Sub Main()
  
'  LeerIni
  
'   29/12/2011 - NO LO USO MAS
'
'   Me conecto en el momento que necesito realizar una operación
'   ConectarBD
  
'   29/12/2011 - NO LO USO MAS
'
'   Se conecta por seguridad integrada, no requiere mas un Login
'   frmLogin.Show
  
End Sub

'   29/12/2011 - NO LO USO MAS
'

'Private Sub ConectarBD()
'
'  'Conexion con la BD
'  On Error GoTo ErrorHandler
'    BD.ConnectionString = RutaBD
'    'BD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & RutaBD & ";"
'    BD.Open
'
'ErrorHandler:
'  If BD.State = adStateClosed Or Err.Number <> 0 Then
'    MsgBox "No se pudo conectar con la Base de datos. Compruebe la ruta en el archivo PozosFuturos.ini. Aplicación terminada", vbCritical + vbOKOnly, "Error"
'    End
'  End If
'
'End Sub

'Private Sub LeerIni()
'
'  Dim FSO As New FileSystemObject
'  Dim Arch As TextStream
'  Dim Linea As String
'
'  'Nombre del INI debe ser igual al de la aplicación
'  Set Arch = FSO.OpenTextFile(App.Path & "\" & App.EXEName & ".ini", ForReading)
'
'  While Not Arch.AtEndOfStream
'    Linea = Arch.ReadLine
'    If Mid(Linea, 1, 10) = "ConnString" Then
'      RutaBD = Linea
'    End If
'  Wend
  
'End Sub
