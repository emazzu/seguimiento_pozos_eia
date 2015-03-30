Attribute VB_Name = "FuncionesGenerales"
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub ErrorHandler()
  If Err.Number <> 0 Then
    MsgBox "Se ha producido el siguiente error: " & Chr(vbKeyReturn) & Err.Description, vbCritical + vbOKOnly, "Error"
  End If
End Sub

Public Function ObtenerOrdenUbicacion(Ubicacion As String) As Long
'Devuelve el orden de la ubicacion
  On Error GoTo ErrorHandler
  
  Dim RS As New Recordset
  RS.Open "SELECT TOP 1 ORDENUBICACION, COUNT(*) FROM POZOS WHERE UBICACION = '" & Ubicacion & "' GROUP BY ORDENUBICACION ORDER BY 2 DESC", BD, adOpenStatic, adLockOptimistic, adCmdText
  If Not RS.EOF Then
    ObtenerOrdenUbicacion = RS(0)
  Else
    ObtenerOrdenUbicacion = -1
  End If
  
ErrorHandler:
  ErrorHandler
End Function

Public Function Invertir(ByRef Valor1 As Variant, ByRef Valor2 As Variant)
'Invierte los valores de valor1 y valor2
  On Error GoTo ErrorHandler
  Dim aux As Variant
  
  aux = Valor1
  Valor1 = Valor2
  Valor2 = aux
  
ErrorHandler:
  ErrorHandler
End Function

Public Sub ReiniciarFormulario(Form As Form)
'Reinicia los controles del formulario que no tienen una X en el tag
  On Error Resume Next
  
  Dim Control As Control
  For Each Control In Form.Controls
    If TypeOf Control Is TextBox And Control.Tag <> "X" Then Control = ""
    If TypeOf Control Is ComboBox And Control.Tag <> "X" Then If Control.ListCount > 0 Then Control.ListIndex = 0
    If TypeOf Control Is ListView And Control.Tag <> "X" Then Control.ListItems.Clear
    If TypeOf Control Is CheckBox And Control.Tag <> "X" Then Control.Value = vbUnchecked
    If TypeOf Control Is DTPicker And Control.Tag <> "X" Then Control.Value = Date
    If TypeOf Control Is MSHFlexGrid And Control.Tag <> "X" Then Control.Rows = 1
    If TypeOf Control Is SSTab And Control.Tag <> "X" Then Control.Tab = 0
  Next

End Sub

Public Sub LogCrearArchivo(RutaDestino As String, NombreArchivo As String)
'Crea un nuevo archivo de log en la ruta especificada
  On Error GoTo ErrorHandler
  
  Dim FSO As New FileSystemObject
  DirectorioLog = RutaDestino
  NombreArchivoLog = NombreArchivo
  If Not FSO.FolderExists(RutaDestino) Then
    MsgBox "No se pudo crear el archivo " & NombreArchivo & ". El directorio especificado no existe", vbCritical + vbOKOnly, "Error"
    Exit Sub
  Else
    Set ArchivoLog = FSO.OpenTextFile(DirectorioLog & "\" & NombreArchivoLog, ForAppending, True)
  End If

ErrorHandler:
  ErrorHandler
End Sub


Public Sub LogEscribirLinea(Texto As String)
'Escribe una nueva linea de texto
  On Error Resume Next
  
  Dim FSO As New FileSystemObject
  If Not FSO.FileExists(DirectorioLog & "\" & NombreArchivoLog) Then
    MsgBox "No se pudo escribir el archivo " & NombreArchivoLog & ". Ha dejado de existir, fue movido, o renombrado", vbCritical + vbOKOnly, "Error"
    Exit Sub
  Else
    ArchivoLog.WriteLine Texto
  End If
End Sub

Public Sub LogEscribirLineasVacias(Cantidad As Long)
'Genera <Cantidad> lineas vacias
  On Error GoTo ErrorHandler
  
  Dim FSO As New FileSystemObject
  If Not FSO.FileExists(DirectorioLog & "\" & NombreArchivoLog) Then
    MsgBox "No se pudo escribir el archivo " & NombreArchivoLog & ". Ha dejado de existir, fue movido, o renombrado", vbCritical + vbOKOnly, "Error"
    Exit Sub
  Else
    ArchivoLog.WriteBlankLines Cantidad
  End If
  
ErrorHandler:
  ErrorHandler
End Sub


Public Sub LogEscribirAContinuacion(Texto As String)
'Escribe a continuacion en la linea actual del archivo
  On Error GoTo ErrorHandler
  
  Dim FSO As New FileSystemObject
  If Not FSO.FileExists(DirectorioLog & "\" & NombreArchivoLog) Then
    MsgBox "No se pudo escribir el archivo " & NombreArchivoLog & ". Ha dejado de existir, fue movido, o renombrado", vbCritical + vbOKOnly, "Error"
    Exit Sub
  Else
    ArchivoLog.Write Texto
  End If
  
ErrorHandler:
  ErrorHandler
End Sub


Public Sub LogCerrarArchivo()
'Cierra el archivo de log
  On Error GoTo ErrorHandler
  
  Dim FSO As New FileSystemObject
  If Not FSO.FileExists(DirectorioLog & "\" & NombreArchivoLog) Then
    MsgBox "No se pudo cerrar el archivo " & NombreArchivoLog & ". Ha dejado de existir, fue movido, o renombrado", vbCritical + vbOKOnly, "Error"
    Exit Sub
  Else
    LogEscribirLineasVacias 1
    ArchivoLog.Close
  End If
 
ErrorHandler:
  ErrorHandler
End Sub


Public Sub LogMostrar(Formulario As Form)
'Ejecuta el archivo log
  On Error GoTo ErrorHandler
  
  Const MostrarRestaurado = 1
  Dim FSO As New FileSystemObject
  If Not FSO.FileExists(DirectorioLog & "\" & NombreArchivoLog) Then
    MsgBox "No se pudo abrir el archivo " & NombreArchivoLog & ". Ha dejado de existir, fue movido, o renombrado", vbCritical + vbOKOnly, "Error"
    Exit Sub
  Else
    ShellExecute Formulario.hWnd, "open", DirectorioLog & "\" & NombreArchivoLog, vbNullString, vbNullString, MostrarRestaurado
  End If
    
ErrorHandler:
  ErrorHandler
End Sub


Public Function ObtenerValorCampo(Tabla As String, Campo As String, Optional Condicion As String = "") As Variant
'Obtiene el valor de un solo campo de una tabla que cumpla con una condicion
  On Error GoTo ErrorHandler
  Dim RS As New Recordset
  
  RS.Open "SELECT " & Campo & " FROM " & Tabla & IIf(Condicion = "", "", " WHERE " & Condicion), BD, adOpenDynamic, adLockOptimistic, adCmdText
  ObtenerValorCampo = Null
  If Not RS.EOF Then
    ObtenerValorCampo = RS(0)
  End If
  RS.Close
  
ErrorHandler:
  ErrorHandler
End Function


Public Sub GuardarHistorial(Tabla As String, CampoID As String, ID As Long, TipoMovimiento As String)
'Guarda un historial de todas las operaciones
  On Error GoTo ErrorHandler

  Dim RSOriginal As New Recordset
  Dim RSHistorial As New Recordset
  Dim i As Long
  Dim j As Long
  
  If ID = 0 Then 'se dio un alta
    RSOriginal.Open "SELECT TOP 1 * FROM " & Tabla & " ORDER BY 1 DESC", BD, adOpenDynamic, adLockOptimistic, adCmdText
    RSHistorial.Open "HIST_" & Tabla, BD, adOpenDynamic, adLockOptimistic, adCmdTable
  Else
    RSOriginal.Open "SELECT * FROM " & Tabla & " WHERE " & CampoID & " = " & ID, BD, adOpenDynamic, adLockOptimistic, adCmdText
    RSHistorial.Open "HIST_" & Tabla, BD, adOpenDynamic, adLockOptimistic, adCmdTable
  End If
  If Not RSOriginal.EOF Then
    'el FOR empieza desde uno para no actualizar el campo id
    RSHistorial.AddNew
    For i = 1 To RSOriginal.Fields.Count - 1
      For j = 1 To RSHistorial.Fields.Count - 1
        If RSOriginal.Fields(i).Name = RSHistorial.Fields(j).Name Then
          RSHistorial.Fields(j).Value = RSOriginal.Fields(i).Value
          Exit For
        End If
      Next j
    Next i
    RSHistorial!TipoMovimiento = TipoMovimiento
    If ID <> 0 Then
      RSHistorial!XID = ID
    Else
      RSHistorial!XID = RSOriginal.Fields(0).Value
    End If
    RSHistorial!XUSUARIO = InfoGlobal.Usuario
    RSHistorial!XFECHA = Now
    RSHistorial.Update
    RSHistorial.Close
    RSOriginal.Close
  End If
    
ErrorHandler:
  ErrorHandler
End Sub

Public Sub EliminarRegistro(Formulario As Form, Tabla As String, CampoID As String, ID As Long, Optional Reiniciar As Boolean = True, Optional PedirConfirmacion As Boolean = True)
'Elimina un registro de la BD, valida que pueda ser eliminado
  On Error GoTo ErrorHandler
 
  BD.BeginTrans
  If ID <> 0 Then
    If PedirConfirmacion Then
      If MsgBox("¿Confirma que desea eliminar el registro seleccionado?", vbYesNo + vbQuestion, "Confirmar eliminación") = vbYes Then
        GuardarHistorial Tabla, CampoID, ID, "B"
        
    
        BD.Execute "DELETE FROM " & Tabla & " WHERE " & CampoID & " = " & ID
        If Reiniciar Then
          UltimoPozoSeleccionado = ""
          Formulario.cmdReiniciar_Click
        End If
      End If
    Else
      GuardarHistorial Tabla, CampoID, ID, "B"
      
      BD.Execute "DELETE FROM " & Tabla & " WHERE " & CampoID & " = " & ID
        If Reiniciar Then
          UltimoPozoSeleccionado = ""
          Formulario.cmdReiniciar_Click
        End If
    End If
  Else
    MsgBox "Debe seleccionar un registro para eliminar", vbInformation + vbOKOnly, "Atención"
  End If

ErrorHandler:
  If Err.Number <> 0 Then
    BD.RollbackTrans
    MsgBox "El registro seleccionado no puede ser eliminado porque tiene registros relacionados", vbInformation + vbOKOnly, "Atención"
  Else
    BD.CommitTrans
  End If
End Sub


Public Sub FlexGridLlenar(FlexGrid As MSHFlexGrid, RS As Recordset)
'Vuelca el contenido del RS en el FlexGrid
  On Error GoTo ErrorHandler
  Dim AnchosColumna() As Double
  Dim i As Integer
  
  Screen.MousePointer = vbHourglass
  FlexGrid.Redraw = False 'Deshabilita el repintado del Flexgrid
  FlexGrid.Rows = RS.RecordCount + 1 'Agrega las filas necesarias en el FlexGRid
  If RS.RecordCount > 0 Then
    FlexGrid.FixedRows = 1
  End If
  FlexGrid.Enabled = RS.RecordCount > 0
  FlexGrid.Cols = RS.Fields.Count 'Agrega las columnas necesarias
  'Recorre los campos del recordset
  For i = 0 To RS.Fields.Count - 1
    FlexGrid.TextMatrix(0, i) = RS.Fields(i).Name 'Agrega los encabezados de columna
  Next
  'Selecciona
  If RS.RecordCount > 0 Then
    FlexGrid.Row = 1
    FlexGrid.Col = 0
    FlexGrid.RowSel = FlexGrid.Rows - 1
    FlexGrid.ColSel = FlexGrid.Cols - 1
    'Devuelve o establece el contenido de las celdas en una región de FlexGrid seleccionada.
    Set FlexGrid.Recordset = RS
    FlexGrid.Clip = RS.GetString(adClipString, -1, Chr(vbKeyTab), Chr(vbKeyReturn), vbNullString)
    FlexGrid.Row = 1
  End If
  'Asigna los anchos de las columnas
  For i = 0 To RS.Fields.Count - 1
    If (InStr(1, UCase(RS.Fields(i).Name), "ID") And Len(RS.Fields(i).Name) = 2) Or (InStr(1, UCase(RS.Fields(i).Name), "ID ") And Len(RS.Fields(i).Name) > 2) Or UCase(RS.Fields(i).Name) = "RIGORDERCHECKED" Then
      FlexGrid.ColWidth(i) = 0
    Else
      Select Case i
         Case colTD, colXPDC, colYPDC, colXPOS94, colYPOS94, colTotDays, colRemDays, colTiempoEntrePedidoETIAyRecepcionETIA, colTiempoEntrePozoInformadoYRecepcionETIA, colTiempoEntrePrimerMonografiaYRecepcionETIA, colTiempoEntreRecepcionETIAYPresentacionAnteDMA, colTIempoEntrePresentacionAnteDMAYVisita, colTiempoEntreVisitaYAprobacionFinalDeDMA, colTiempoEntrePresentacionDePozoYAprobacionFinalFMA: FlexGrid.ColWidth(i) = 1500: FlexGrid.ColAlignment(i) = flexAlignRightCenter
         Case colWellID, colUbicacion, colEquipo, colYacimiento, colTipoYacimiento, colPozo, colProspect, colFieldManifold, colBatteryAssigned, colInformedBy, colDocumentoAPreparar, colLandOwner, colConsult, colType: FlexGrid.ColWidth(i) = 5000: FlexGrid.ColAlignment(i) = flexAlignLeftCenter
         Case colSiteVisit, colImageInterpretationComments, colConsultantRecomendation: FlexGrid.ColWidth(i) = 10000: FlexGrid.ColAlignment(i) = flexAlignLeftCenter
         Case Else: FlexGrid.ColWidth(i) = 3000: FlexGrid.ColAlignment(i) = flexAlignCenterCenter
      End Select
    End If
  Next
  ' habilita nuevamente el Redraw en el control
  FlexGrid.Redraw = True
  Screen.MousePointer = vbDefault

ErrorHandler:
 ErrorHandler
End Sub

Public Sub ProcesarBusqueda(ByVal Sql As String, Grilla As MSHFlexGrid, UbicacionIN As String, Optional AliasFiltro As String = "", Optional CampoFiltro As String = "", Optional ValorFiltro As String = "", Optional FiltrarNulls As CheckBoxConstants = vbGrayed)
'Arma el query y abre el cursor que se usara para llenar la lista
  On Error GoTo ErrorHandler
  Dim RS As New Recordset
  Dim i As Long
  
  If ValorFiltro <> "" And CampoFiltro <> "" Then
    Select Case Mid(AliasFiltro, InStr(1, AliasFiltro, "("))
      Case "(Texto)"
        If InStr(InStr(1, Sql, "FROM POZOS"), UCase(Sql), "WHERE") = 0 Then
          Sql = Sql & " WHERE UCASE(" & CampoFiltro & ") LIKE '%" & UCase(ValorFiltro) & "%'"
        Else
          Sql = Sql & " AND UCASE(" & CampoFiltro & ") LIKE '%" & UCase(ValorFiltro) & "%'"
        End If
      Case "(Fecha)"
        If InStr(InStr(1, Sql, "FROM POZOS"), UCase(Sql), "WHERE") = 0 Then
          If IsNumeric(ValorFiltro) Then
            Sql = Sql & " WHERE MONTH(" & CampoFiltro & ") = " & ValorFiltro
          ElseIf IsDate(ValorFiltro) Then
            If ContarCaracter(ValorFiltro, "/") = 1 Then
              Sql = Sql & " WHERE MONTH(" & CampoFiltro & ") = " & Mid(ValorFiltro, 1, InStr(1, ValorFiltro, "/") - 1) & " AND YEAR(" & CampoFiltro & ") = " & Mid(ValorFiltro, InStr(1, ValorFiltro, "/") + 1)
            Else
              Sql = Sql & " WHERE " & CampoFiltro & " = #" & Format(ValorFiltro, "yyyy/mm/dd") & "#"
            End If
          Else
            MsgBox "Formato de fecha incorrecta. Ejemplos: " & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "02: Busca registros de febrero" & Chr(vbKeyReturn) & "02/2009: Busca registros de febrero de 2009" & Chr(vbKeyReturn) & "03/02/2009: Busca registros del 03 de febrero de 2009", vbOKOnly + vbInformation, "Atención"
            Exit Sub
          End If
        Else
          If IsNumeric(ValorFiltro) Then
            Sql = Sql & " AND MONTH(" & CampoFiltro & ") = " & ValorFiltro
          ElseIf IsDate(ValorFiltro) Then
            If ContarCaracter(ValorFiltro, "/") = 1 Then
              Sql = Sql & " AND MONTH(" & CampoFiltro & ") = " & Mid(ValorFiltro, 1, InStr(1, ValorFiltro, "/") - 1) & " AND YEAR(" & CampoFiltro & ") = " & Mid(ValorFiltro, InStr(1, ValorFiltro, "/") + 1)
            Else
              Sql = Sql & " AND " & CampoFiltro & " = #" & ValorFiltro & "#"
            End If
          Else
            MsgBox "Formato de fecha incorrecta. Ejemplos: " & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "02: Busca registros de febrero" & Chr(vbKeyReturn) & "02/2009: Busca registros de febrero de 2009" & Chr(vbKeyReturn) & "03/02/2009: Busca registros del 03 de febrero de 2009", vbOKOnly + vbInformation, "Atención"
            Exit Sub
          End If
        End If
      Case "(Número)"
        If IsNumeric(ValorFiltro) Then
          If InStr(InStr(1, Sql, "FROM POZOS"), UCase(Sql), "WHERE") = 0 Then
            Sql = Sql & " WHERE " & CampoFiltro & " = " & ValorFiltro
          Else
            Sql = Sql & " AND " & CampoFiltro & " = " & ValorFiltro
          End If
        Else
          MsgBox "Debe ingresar un numero para el filtro " & AliasFiltro, vbOKOnly + vbInformation, "Atención"
          Exit Sub
        End If
    End Select
  End If
  
  Select Case FiltrarNulls
    Case vbChecked
      If InStr(InStr(1, Sql, "FROM POZOS"), UCase(Sql), "WHERE") = 0 Then
        Sql = Sql & " WHERE NOT " & CampoFiltro & " IS NULL"
      Else
        Sql = Sql & " AND NOT " & CampoFiltro & " IS NULL"
      End If
    Case vbUnchecked
      If InStr(InStr(1, Sql, "FROM POZOS"), UCase(Sql), "WHERE") = 0 Then
        Sql = Sql & " WHERE " & CampoFiltro & " IS NULL"
      Else
        Sql = Sql & " AND " & CampoFiltro & " IS NULL"
      End If
    Case vbGrayed
      'NO HACE NADA
  End Select
  
  Dim RSCategorias As New Recordset
  Dim QueryUnion As String
  RSCategorias.Open "SELECT UBICACION FROM (SELECT DISTINCT ORDENUBICACION, UBICACION FROM POZOS WHERE UBICACION IN (" & UbicacionIN & ")) ORDER BY ORDENUBICACION", BD, adOpenDynamic, adLockReadOnly, adCmdText
  While Not RSCategorias.EOF
    Select Case RSCategorias(0)
      Case "RIGS SCHED."
        If InStr(1, UCase(UbicacionIN), RSCategorias(0)) > 0 Then
          If InStr(InStr(1, Sql, "FROM POZOS"), UCase(Sql), "WHERE") = 0 Then
            QueryUnion = Sql & " WHERE UBICACION = '" & RSCategorias(0) & "' ORDER BY A.EQUIPO, A.RIGORDER"
          Else
            QueryUnion = Sql & " AND UBICACION = '" & RSCategorias(0) & "' ORDER BY A.EQUIPO, A.RIGORDER"
          End If
        End If
      Case "DRILLING INV."
        If InStr(InStr(1, Sql, "FROM POZOS"), UCase(Sql), "WHERE") = 0 Then
          QueryUnion = Sql & " WHERE UBICACION = '" & RSCategorias(0) & "' ORDER BY A.YACIMIENTO, A.TIPOYACIMIENTO, A.WELLID"
        Else
          QueryUnion = Sql & " AND UBICACION = '" & RSCategorias(0) & "' ORDER BY A.YACIMIENTO, A.TIPOYACIMIENTO, A.WELLID"
        End If
      Case "GONE / DONE"
        If InStr(InStr(1, Sql, "FROM POZOS"), UCase(Sql), "WHERE") = 0 Then
          QueryUnion = Sql & " WHERE UBICACION = '" & RSCategorias(0) & "' ORDER BY A.ENDDATE DESC"
        Else
          QueryUnion = Sql & " AND UBICACION = '" & RSCategorias(0) & "' ORDER BY A.ENDDATE DESC"
        End If
      Case Else
        If InStr(InStr(1, Sql, "FROM POZOS"), UCase(Sql), "WHERE") = 0 Then
          QueryUnion = Sql & " WHERE UBICACION = '" & RSCategorias(0) & "'"
        Else
          QueryUnion = Sql & " AND UBICACION = '" & RSCategorias(0) & "'"
        End If
    End Select
    RS.CursorLocation = adUseClient
    RS.Open QueryUnion, BD, adOpenStatic, adLockReadOnly, adCmdText
    FlexGridLlenarPozos Grilla, RS
    RS.Close
    RSCategorias.MoveNext
  Wend
  RSCategorias.Close
  
  
  
ErrorHandler:
  ErrorHandler
End Sub

Public Sub Center(Form As Form)
  Form.Left = (frmMenuPrincipal.Width - Form.Width) / 2
  Form.Top = (frmMenuPrincipal.Height - Form.Height) / 2
End Sub

Public Function ContarCaracter(Texto As String, caracter As String) As Long
  Dim i As Long
  i = 1
  While i < Len(Texto)
    If Mid(Texto, i, Len(caracter)) = caracter Then
      ContarCaracter = ContarCaracter + 1
    End If
    i = i + 1
  Wend
End Function


Public Sub BuscarFila(mfg As MSHFlexGrid, KeyCode As Integer, Texto As String)
'Busca una fila para seleccionarla como si se hubiera hecho click en ella con el mouse
  On Error GoTo ErrorHandler
  
  Dim i As Long, j As Long
  
  Select Case KeyCode
    Case vbKeyBack, vbKeyDelete
        Texto = ""
    Case Else
      Texto = Texto & Chr(KeyCode)
      For i = 1 To mfg.Rows - 1
        For j = 0 To mfg.Cols - 1
          If InStr(1, UCase(mfg.TextMatrix(i, j)), UCase(Texto)) Then
            mfg.Row = i
            mfg.TopRow = i
            mfg.Col = 0
            mfg.ColSel = mfg.Cols - 1
            Exit Sub
          End If
        Next j
      Next i
  End Select
  
ErrorHandler:
  ErrorHandler
End Sub


Public Sub FlexGridLlenarPozos(FlexGrid As MSHFlexGrid, RS As Recordset)
'Vuelca el contenido del RS en el FlexGrid
  On Error GoTo ErrorHandler
  Dim AnchosColumna() As Double
  Dim UltimaFila As Long
  Dim i As Integer
  
  Screen.MousePointer = vbHourglass
  FlexGrid.Redraw = False 'Deshabilita el repintado del Flexgrid
  UltimaFila = FlexGrid.Rows
  FlexGrid.Rows = FlexGrid.Rows + RS.RecordCount  'Agrega las filas necesarias en el FlexGRid
  FlexGrid.Cols = RS.Fields.Count 'Agrega las columnas necesarias
  'Recorre los campos del recordset
  For i = 0 To RS.Fields.Count - 1
    FlexGrid.TextMatrix(0, i) = RS.Fields(i).Name 'Agrega los encabezados de columna
  Next
  If FlexGrid.Rows > 1 Then
    FlexGrid.FixedRows = 1
    FlexGrid.Enabled = True
  Else
    FlexGrid.Enabled = False
  End If
  'Selecciona
  If RS.RecordCount > 0 Then
    FlexGrid.Row = UltimaFila
    FlexGrid.Col = 0
    FlexGrid.RowSel = FlexGrid.Rows - 1
    FlexGrid.ColSel = FlexGrid.Cols - 1
    'Devuelve o establece el contenido de las celdas en una región de FlexGrid seleccionada.
    FlexGrid.Clip = RS.GetString(adClipString, -1, Chr(vbKeyTab), Chr(vbKeyReturn), vbNullString)
    FlexGrid.Row = 1
  End If
  'Asigna los anchos de las columnas
  For i = 0 To RS.Fields.Count - 1
    If (InStr(1, UCase(RS.Fields(i).Name), "ID") And Len(RS.Fields(i).Name) = 2) Or (InStr(1, UCase(RS.Fields(i).Name), "ID ") And Len(RS.Fields(i).Name) > 2) Or UCase(RS.Fields(i).Name) = "RIGORDERCHECKED" Then
      FlexGrid.ColWidth(i) = 0
    Else
      Select Case i
         Case colTD, colXPDC, colYPDC, colXPOS94, colYPOS94, colTotDays, colRemDays, colTiempoEntrePedidoETIAyRecepcionETIA, colTiempoEntrePozoInformadoYRecepcionETIA, colTiempoEntrePrimerMonografiaYRecepcionETIA, colTiempoEntreRecepcionETIAYPresentacionAnteDMA, colTIempoEntrePresentacionAnteDMAYVisita, colTiempoEntreVisitaYAprobacionFinalDeDMA, colTiempoEntrePresentacionDePozoYAprobacionFinalFMA: FlexGrid.ColWidth(i) = 1500: FlexGrid.ColAlignment(i) = flexAlignRightCenter
         Case colWellID, colUbicacion, colEquipo, colYacimiento, colTipoYacimiento, colPozo, colProspect, colFieldManifold, colBatteryAssigned, colInformedBy, colDocumentoAPreparar, colLandOwner, colConsult, colType: FlexGrid.ColWidth(i) = 5000: FlexGrid.ColAlignment(i) = flexAlignLeftCenter
         Case colSiteVisit, colImageInterpretationComments, colConsultantRecomendation: FlexGrid.ColWidth(i) = 10000: FlexGrid.ColAlignment(i) = flexAlignLeftCenter
         Case Else: FlexGrid.ColWidth(i) = 3000: FlexGrid.ColAlignment(i) = flexAlignCenterCenter
      End Select
    End If
  Next
  ' habilita nuevamente el Redraw en el control
  FlexGrid.Redraw = True
  Screen.MousePointer = vbDefault

ErrorHandler:
 ErrorHandler
End Sub




Public Sub ProcesarBusqueda2(ByVal Sql As String, Grilla As MSHFlexGrid, OrderBy As String, Optional AliasFiltro As String = "", Optional CampoFiltro As String = "", Optional ValorFiltro As String = "")
'Arma el query y abre el cursor que se usara para llenar la lista
  On Error GoTo ErrorHandler
  Dim RS As New Recordset
  Dim i As Long
  
  If ValorFiltro <> "" And CampoFiltro <> "" Then
    Select Case Mid(AliasFiltro, InStr(1, AliasFiltro, "("))
      Case "(Texto)"
        If InStr(1, UCase(Sql), "WHERE") = 0 Then
          Sql = Sql & " WHERE UPPER(" & CampoFiltro & ") LIKE '%" & UCase(ValorFiltro) & "%'"
        Else
          Sql = Sql & " AND UPPER(" & CampoFiltro & ") LIKE '%" & UCase(ValorFiltro) & "%'"
        End If
      Case "(Fecha)"
        If InStr(1, UCase(Sql), "WHERE") = 0 Then
          If IsNumeric(ValorFiltro) Then
            Sql = Sql & " WHERE MONTH(" & CampoFiltro & ") = " & ValorFiltro
          ElseIf IsDate(ValorFiltro) Then
            If ContarCaracter(ValorFiltro, "/") = 1 Then
              Sql = Sql & " WHERE MONTH(" & CampoFiltro & ") = " & Mid(ValorFiltro, 1, InStr(1, ValorFiltro, "/") - 1) & " AND YEAR(" & CampoFiltro & ") = " & Mid(ValorFiltro, InStr(1, ValorFiltro, "/") + 1)
            Else
              Sql = Sql & " WHERE " & CampoFiltro & " = CONVERT(DATETIME,'" & ValorFiltro & "')"
            End If
          Else
            MsgBox "Formato de fecha incorrecta. Ejemplos: " & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "02: Busca registros de febrero" & Chr(vbKeyReturn) & "02/2009: Busca registros de febrero de 2009" & Chr(vbKeyReturn) & "03/02/2009: Busca registros del 03 de febrero de 2009", vbOKOnly + vbInformation, "Atención"
            Exit Sub
          End If
        Else
          If IsNumeric(ValorFiltro) Then
            Sql = Sql & " AND MONTH(" & CampoFiltro & ") = " & ValorFiltro
          ElseIf IsDate(ValorFiltro) Then
            If ContarCaracter(ValorFiltro, "/") = 1 Then
              Sql = Sql & " AND MONTH(" & CampoFiltro & ") = " & Mid(ValorFiltro, 1, InStr(1, ValorFiltro, "/") - 1) & " AND YEAR(" & CampoFiltro & ") = " & Mid(ValorFiltro, InStr(1, ValorFiltro, "/") + 1)
            Else
              Sql = Sql & " AND " & CampoFiltro & " = CONVERT(DATETIME,'" & ValorFiltro & "')"
            End If
          Else
            MsgBox "Formato de fecha incorrecta. Ejemplos: " & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "02: Busca registros de febrero" & Chr(vbKeyReturn) & "02/2009: Busca registros de febrero de 2009" & Chr(vbKeyReturn) & "03/02/2009: Busca registros del 03 de febrero de 2009", vbOKOnly + vbInformation, "Atención"
            Exit Sub
          End If
        End If
      Case "(Número)"
        If IsNumeric(txt) Then
          If InStr(1, UCase(Sql), "WHERE") = 0 Then
            Sql = Sql & " WHERE " & CampoFiltro & " = " & ValorFiltro
          Else
            Sql = Sql & " AND " & CampoFiltro & " = " & ValorFiltro
          End If
        Else
          MsgBox "Debe ingresar un numero para este filtro", vbOKOnly + vbInformation, "Atención"
          Exit Sub
        End If
    End Select
  End If
  
  Sql = Sql & " ORDER BY " & OrderBy
  RS.CursorLocation = adUseClient
  RS.Open Sql, BD, adOpenStatic, adLockReadOnly, adCmdText
  FlexGridLlenar Grilla, RS

ErrorHandler:
  ErrorHandler
End Sub



'
'Private Sub mfg_KeyDown(KeyCode As Integer, Shift As Integer)
''Detecto si se presiona Ctrl o Shift
'  On Error GoTo ErrorHandler
'
'   If Shift = vbShiftMask Then
'      m_booKeyShift = True
'   End If
'   If Shift = vbCtrlMask Then
'      m_booKeyCtrl = True
'   End If
'
'ErrorHandler:
'  ErrorHandler
'End Sub
'
'
'Private Sub mfg_Keyup(KeyCode As Integer, Shift As Integer)
''Seteo que se soltaron las teclas
'  On Error GoTo ErrorHandler
'
'   m_booKeyCtrl = False
'   m_booKeyShift = False
'
'ErrorHandler:
'  ErrorHandler
'End Sub
'
'
'Private Sub mfg_RowColChange()
''Rutina de seleccion de filas
'  On Error GoTo ErrorHandler
'
'  Static Ocupado As Boolean
'  Dim ColumnaActual As Integer
'  Dim FilaActual As Integer
'
'  With mfg
'    If m_booKeyShift Or Ocupado Then
'      Exit Sub
'    Else
'      Ocupado = True
'    End If
'    ColumnaActual = .Col
'    FilaActual = .Row
'    LockWindowUpdate .hwnd
'    If m_booKeyCtrl Then
'      .Col = 1
'      .Row = FilaActual
'      .ColSel = .Cols - 1
'      .RowSel = FilaActual
'      If .CellBackColor = .BackColorSel Then
'         .CellBackColor = .BackColor
'         .CellForeColor = .ForeColor
'      Else
'         .CellBackColor = .BackColorSel
'         .CellForeColor = .ForeColorSel
'      End If
'    Else
'      .Col = 1
'      .Row = 1
'      .ColSel = .Cols - 1
'      .RowSel = .Rows - 1
'      .FillStyle = flexFillRepeat
'      .CellBackColor = .BackColor
'      .CellForeColor = .ForeColor
'      .Col = 1
'      .Row = FilaActual
'      .ColSel = .Cols - 1
'      .RowSel = FilaActual
'      .CellBackColor = .BackColorSel
'      .CellForeColor = .ForeColorSel
'    End If
'    .Col = ColumnaActual
'    .Row = FilaActual
'    LockWindowUpdate 0&
'    Ocupado = False
'  End With
'
'ErrorHandler:
'  ErrorHandler
'End Sub
'
'Private Sub mfg_SelChange()
''Rutina de seleccion de filas
'  On Error GoTo ErrorHandler
'  Dim i As Long
'  Dim ColumnaActual As Integer
'  Dim FilaActual As Integer
'  Dim SiguienteColumna As Integer
'  Dim SiguienteFila As Integer
'  With mfg
'    If Not m_booKeyShift Then
'      Exit Sub
'    End If
'   ColumnaActual = .Col
'   FilaActual = .Row
'   SiguienteColumna = .ColSel
'   SiguienteFila = .RowSel
'   LockWindowUpdate .hwnd
'   .Col = 1
'   '.Row = 1
'   .ColSel = .Cols - 1
'   .RowSel = .Rows - 1
'   .FillStyle = flexFillRepeat
'   .CellBackColor = .BackColor
'   .CellForeColor = .ForeColor
'
'   ' Update Multiline
'   .Col = 1
'   .Row = FilaActual
'   .ColSel = .Cols - 1
'   .RowSel = SiguienteFila
'   .FillStyle = flexFillRepeat
'   .CellBackColor = .BackColorSel
'   .CellForeColor = .ForeColorSel
'   LockWindowUpdate 0&
'   .Col = SiguienteColumna
'   .Row = SiguienteFila
'  End With
'
'ErrorHandler:
'  ErrorHandler
'End Sub

